#region Copyright Syncfusion Inc. 2001 - 2026
//
//  Copyright Syncfusion Inc. 2001 - 2026. All rights reserved.
//
//  Use of this code is subject to the terms of our license.
//  A copy of the current license can be obtained at any time by e-mailing
//  licensing@syncfusion.com. Re-distribution in any form is strictly
//  prohibited. Any infringement will be prosecuted under applicable laws. 
//
#endregion

using Syncfusion.AI.AgentTools.Core;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.Text.Json;

namespace Syncfusion.AI.AgentTools.Word
{
    /// <summary>
    /// Provides agent tools for managing form fields in Word documents.
    /// Handles form field data extraction, updates, and individual field operations.
    /// </summary>
    public class WordFormFieldAgentTools : AgentToolBase<WordDocument>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WordFormFieldAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The document manager for managing Word documents.</param>
        public WordFormFieldAgentTools(WordDocumentManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordFormFieldAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public WordFormFieldAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Retrieves all form field data as a dictionary.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// /// <returns>Result containing a dictionary of form field names and their values.</returns>
        [Tool(
            Name = "GetFormData",
            Description = "Retrieves all form field data as a dictionary. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult GetFormData(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                var formData = new Dictionary<string, object>();

                foreach (WSection section in document.Sections)
                {
                    ExtractFormFieldsFromBody(section.Body, formData);
                }

                // Convert all values to strings for safe JSON serialization.
                // System.Text.Json cannot reliably serialize Dictionary<string, object>
                // when values are mixed types (bool, string) inside an anonymous Data object.
                var formDataStrings = new Dictionary<string, string>(formData.Count);
                var sb = new System.Text.StringBuilder();
                sb.AppendLine($"Retrieved {formData.Count} form field(s):");
                foreach (var kvp in formData)
                {
                    string strVal = kvp.Value?.ToString() ?? string.Empty;
                    formDataStrings[kvp.Key] = strVal;
                    sb.AppendLine($"  {kvp.Key} = {strVal}");
                }

                return AgentToolResult.Ok(sb.ToString(), new FormDataResult
                {
                    Count = formData.Count,
                    FormData = formDataStrings
                });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to get form data: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets one or more form field values using a JSON string. Always use this tool to set any form field values.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="dataJson">A JSON object string containing all form field names and their corresponding values.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating whether the form fields were updated successfully.</returns>
        [Tool(
            Name = "SetFormFields",
            Description = "The only tool for setting form field values. Pass ALL field name/value pairs as a single JSON object string and call this tool exactly ONCE. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Example: '{\"TextField1\":\"Hello\",\"Check1\":true,\"DropDown1\":\"Option A\"}'")]
        public AgentToolResult SetFormFields(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "A JSON object string containing all field names and their values. Text fields accept strings, checkboxes accept true/false, drop-downs accept the item text or a zero-based index. Example: '{\"PatientName\":\"John\",\"Diabetes\":true,\"Gender\":\"Male\"}'")]
            string dataJson,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                // Deserialize the JSON string into a flat dictionary
                Dictionary<string, object> data;
                try
                {
                    var jsonDoc = JsonDocument.Parse(dataJson);
                    data = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    foreach (var prop in jsonDoc.RootElement.EnumerateObject())
                        data[prop.Name] = prop.Value.Clone(); // Clone preserves value after Dispose
                }
                catch (JsonException ex)
                {
                    return AgentToolResult.Fail($"Invalid JSON in dataJson: {ex.Message}");
                }

                // Track which keys were matched during the single document traversal
                var updatedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var setErrors = new List<string>();   // real exceptions while setting a field
                var notFound = new List<string>();    // keys supplied that don't exist in the doc

                foreach (WSection section in document.Sections)
                {
                    ApplyFormFieldsFromBody(section.Body, data, updatedFields, setErrors);
                }

                // Identify supplied keys that were never matched to any field in the document
                foreach (var key in data.Keys)
                {
                    if (!updatedFields.Contains(key))
                        notFound.Add(key);
                }

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_form_filled.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                var result = $"Updated {updatedFields.Count} of {data.Count} field(s) successfully into document {outputKey}.";
                if (setErrors.Count > 0)
                    result += $" {setErrors.Count} field(s) had errors.";
                if (notFound.Count > 0)
                    result += $" {notFound.Count} key(s) not found in document: {string.Join(", ", notFound)}.";

                return AgentToolResult.Ok(result, new
                {
                    UpdatedCount = updatedFields.Count,
                    TotalCount = data.Count,
                    SetErrors = setErrors.Count > 0 ? setErrors : null,
                    NotFoundKeys = notFound.Count > 0 ? notFound : null
                });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to set form fields: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets a specific form field value by field name.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="fieldName">The name of the form field to retrieve.</param>
        /// <returns>Result containing the value of the specified form field.</returns>
        [Tool(
            Name = "GetFormField",
            Description = "Gets a specific form field value by field name. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult GetFormField(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The name of the form field")]
            string fieldName)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                var formData = new Dictionary<string, object>();
                foreach (WSection section in document.Sections)
                {
                    ExtractFormFieldsFromBody(section.Body, formData);
                }

                if (formData.TryGetValue(fieldName, out object? value))
                {
                    return AgentToolResult.Ok($"Field Name: {fieldName}, Value: {value}", new
                    {
                        FieldName = fieldName,
                        Value = value
                    });
                }
                else
                {
                    return AgentToolResult.Fail($"Form field not found: {fieldName}");
                }
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to get form field: {ex.Message}");
            }
        }

        /// <summary>
        /// Recursively extracts form field data from document body.
        /// For drop-downs the selected item TEXT is returned (not the index) so the
        /// AI can echo back the same text when calling SetFormFields.
        /// </summary>
        private void ExtractFormFieldsFromBody(WTextBody body, Dictionary<string, object> formData)
        {
            foreach (var entity in body.ChildEntities)
            {
                if (entity is WParagraph paragraph)
                {
                    foreach (var item in paragraph.ChildEntities)
                    {
                        if (item is WTextFormField textField)
                        {
                            formData[textField.Name] = textField.TextRange.Text;
                        }
                        else if (item is WCheckBox checkBox)
                        {
                            formData[checkBox.Name] = checkBox.Checked;
                        }
                        else if (item is WDropDownFormField dropDown)
                        {
                            // Return the selected item text so the AI can pass the
                            // same text back when setting — not the raw index.
                            int idx = dropDown.DropDownSelectedIndex;
                            formData[dropDown.Name] = (idx >= 0 && idx < dropDown.DropDownItems.Count)
                                ? dropDown.DropDownItems[idx]
                                : idx.ToString();
                        }
                    }
                }
                else if (entity is WTable table)
                {
                    foreach (WTableRow row in table.Rows)
                    {
                        foreach (WTableCell cell in row.Cells)
                        {
                            ExtractFormFieldsFromBody(cell, formData);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Unwraps a JsonElement to its native CLR type so Convert.* calls succeed.
        /// </summary>
        private static object UnwrapJsonElement(object value)
        {
            if (value is System.Text.Json.JsonElement jsonElement)
            {
                return jsonElement.ValueKind switch
                {
                    System.Text.Json.JsonValueKind.True => true,
                    System.Text.Json.JsonValueKind.False => false,
                    System.Text.Json.JsonValueKind.Number => jsonElement.TryGetInt32(out int i) ? (object)i : jsonElement.GetDouble(),
                    System.Text.Json.JsonValueKind.String => jsonElement.GetString() ?? string.Empty,
                    _ => jsonElement.ToString()
                };
            }
            return value;
        }

        /// <summary>
        /// Converts a value to bool, accepting true/false, 1/0, yes/no, on/off strings.
        /// </summary>
        private static bool ParseBoolean(object value)
        {
            if (value is bool b) return b;
            string s = value?.ToString()?.Trim() ?? string.Empty;
            return s.ToLowerInvariant() switch
            {
                "true" or "1" or "yes" or "on" => true,
                "false" or "0" or "no" or "off" => false,
                _ => Convert.ToBoolean(value)   // fallback — may throw with a clear message
            };
        }

        /// <summary>
        /// Single-pass traversal: for every form field found, checks the dictionary and sets the value if matched.
        /// Dictionary lookup is case-insensitive (field names in documents are often PascalCase).
        /// </summary>
        private void ApplyFormFieldsFromBody(WTextBody body, Dictionary<string, object> data,
            HashSet<string> updatedFields, List<string> errors)
        {
            foreach (var entity in body.ChildEntities)
            {
                if (entity is WParagraph paragraph)
                {
                    for(int i = 0; i<paragraph.ChildEntities.Count; i++)
                    {
                        ParagraphItem paragraphItem = paragraph.ChildEntities[i] as ParagraphItem;
                        try
                        {
                            if (paragraphItem is WTextFormField textField && data.TryGetValue(textField.Name, out object? textValue))
                            {
                                textField.Text = UnwrapJsonElement(textValue)?.ToString() ?? string.Empty;
                                updatedFields.Add(textField.Name);
                            }
                            else if (paragraphItem is WCheckBox checkBox && data.TryGetValue(checkBox.Name, out object? checkValue))
                            {
                                checkBox.Checked = ParseBoolean(UnwrapJsonElement(checkValue));
                                updatedFields.Add(checkBox.Name);
                            }
                            else if (paragraphItem is WDropDownFormField dropDown && data.TryGetValue(dropDown.Name, out object? dropValue))
                            {
                                object unwrapped = UnwrapJsonElement(dropValue);
                                // If the value is numeric, use it directly as the index;
                                // otherwise find the matching item in DropDownItems by text.
                                if (unwrapped is int || unwrapped is long || unwrapped is double ||
                                    (unwrapped is string s && int.TryParse(s, out _)))
                                {
                                    dropDown.DropDownSelectedIndex = Convert.ToInt32(unwrapped);
                                }
                                else
                                {
                                    string itemText = unwrapped?.ToString() ?? string.Empty;
                                    int matchIndex = -1;
                                    for (int j = 0; j < dropDown.DropDownItems.Count; j++)
                                    {
                                        if (string.Equals(dropDown.DropDownItems[j].Text, itemText, StringComparison.OrdinalIgnoreCase))
                                        {
                                            matchIndex = j;
                                            break;
                                        }
                                    }
                                    if (matchIndex == -1)
                                        throw new ArgumentException($"Drop-down item '{itemText}' not found in field '{dropDown.Name}'.");
                                    dropDown.DropDownSelectedIndex = matchIndex;
                                }
                                updatedFields.Add(dropDown.Name);
                            }
                        }
                        catch (Exception ex)
                        {
                            errors.Add($"Error setting field: {ex.Message}");
                        }
                    }
                }
                else if (entity is WTable table)
                {
                    foreach (WTableRow row in table.Rows)
                    {
                        foreach (WTableCell cell in row.Cells)
                        {
                            ApplyFormFieldsFromBody(cell, data, updatedFields, errors);
                        }
                    }
                }
            }
        }

        // Concrete serializable result type — anonymous types with mixed-value
        // Dictionary<string, object> cannot be reliably serialized by System.Text.Json
        // when used as the Data payload returned through AIFunctionFactory.
        private sealed class FormDataResult
        {
            public int Count { get; set; }
            public Dictionary<string, string> FormData { get; set; } = new();
        }
    }
}
