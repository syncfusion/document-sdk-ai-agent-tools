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
using System.Data;
using System.IO;
using System.Net.Http;
using System.Text.Json;

namespace Syncfusion.AI.AgentTools.Word
{
    /// <summary>
    /// Provides agent tools for mail merge operations in Word documents.
    /// Handles executing mail merge with various data sources and merge options.
    /// </summary>
    public class WordMailMergeAgentTools : AgentToolBase<WordDocument>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WordMailMergeAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The document manager for managing Word documents.</param>
        public WordMailMergeAgentTools(WordDocumentManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordMailMergeAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public WordMailMergeAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Word) { }

        #region MergeImageField event

        /// <summary>
        /// Handles the <c>MergeImageField</c> event raised by DocIO for every merge field
        /// whose name starts with <c>"Image:"</c>.
        /// <para>
        /// The data value stored in <see cref="MergeImageFieldEventArgs.ImageFileName"/> is
        /// interpreted as follows (in order):
        /// <list type="number">
        ///   <item><description>
        ///     <b>HTTP / HTTPS URL</b> — the image is downloaded with <see cref="HttpClient"/>
        ///     and the response bytes are wrapped in a <see cref="MemoryStream"/>.
        ///   </description></item>
        ///   <item><description>
        ///     <b>Local file path (InMemory mode)</b> — the file is opened as a read-only
        ///     <see cref="FileStream"/> if it exists on disk.
        ///   </description></item>
        ///   <item><description>
        ///     <b>Storage path (DocumentStorage mode)</b> — the file is retrieved from the
        ///     storage manager if it exists.
        ///   </description></item>
        ///   <item><description>
        ///     <b>Anything else</b> — the field is left unmerged (no image is set) so that
        ///     the mail-merge engine can apply its normal <c>ClearFields</c> / keep-as-is
        ///     behaviour.
        ///   </description></item>
        /// </list>
        /// </para>
        /// </summary>
        private void MergeImageFieldHandler(object sender, MergeImageFieldEventArgs args)
        {
            string? imageSource = args.FieldValue?.ToString();

            if (string.IsNullOrEmpty(imageSource))
                return;

            try
            {
                // ── URL (http / https) ────────────────────────────────────────────────
                if (imageSource.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                    imageSource.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                {
                    using var httpClient = new HttpClient();
                    byte[] imageBytes = httpClient.GetByteArrayAsync(imageSource).GetAwaiter().GetResult();
                    args.ImageStream = new MemoryStream(imageBytes);
                    return;
                }

                // ── File path based on mode ───────────────────────────────────────────
                if (Mode == DocumentManagerMode.InMemory)
                {
                    // InMemory mode: check local file system
                    if (File.Exists(imageSource))
                    {
                        args.ImageStream = new FileStream(imageSource, FileMode.Open, FileAccess.Read, FileShare.Read);
                        return;
                    }
                }
                else
                {
                    // DocumentStorage mode: check storage manager
                    if (StorageManager!.HasDocument(imageSource))
                    {
                        args.ImageStream = StorageManager.GetDocumentStream(imageSource);
                        return;
                    }
                }

                // Source is neither a reachable URL nor an existing file path —
                // leave args.ImageStream as null so DocIO respects ClearFields / keep-as-is.
            }
            catch
            {
                // On any error leave the image field unmerged rather than aborting the
                // entire mail-merge operation.
            }
        }

        #endregion

        #region Execute simple mail merge
        /// <summary>
        /// Extended mail merge with field mappings and multiple document output.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the template document.</param>
        /// <param name="dataSourceJson">JSON data source provided as a file path or a raw JSON string.</param>
        /// <param name="removeEmptyParagraphs">Specifies whether empty paragraphs created during mail merge should be removed.</param>
        /// <param name="clearFields">Specifies whether unmerged merge fields should be removed after mail merge.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure of the mail merge operation, along with details of the generated document.</returns>
        [Tool(
            Name = "ExecuteMailMerge",
            Description = "Extended mail merge with field mappings and output options for batch generation. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Returns document IDs of generated documents.")]
        public AgentToolResult ExecuteMailMerge(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the template document")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "Either a file path to a JSON file (e.g. 'C:\\data\\records.json') or a raw JSON string containing the data source")]
            string dataSourceJson,
            [ToolParameter(Description = "Whether to remove the empty paragraphs when the paragraph has only a merge field item, without any data during Mail merge process. (true/false)")]
            bool removeEmptyParagraphs = true,
            [ToolParameter(Description = "Whether to remove unmerged merge fields from the document after mail merge. true = remove unmerged fields (default); false = keep unmerged fields as-is in the output document.")]
            bool clearFields = true,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                // Resolve: if dataSourceJson is a file path, read the file content
                string? jsonContent = ResolveJsonContent(dataSourceJson, out string? resolveError);
                if (resolveError != null)
                    return AgentToolResult.Fail(resolveError);

                // Parse JSON to DataTable
                DataTable dataTable = ParseJsonToDataTable(jsonContent!);

                var generatedDocumentIds = new List<string>();

                // Configure mail merge options
                document.MailMerge.RemoveEmptyParagraphs = removeEmptyParagraphs;
                document.MailMerge.ClearFields = clearFields;

                // Subscribe image event to handle both local file paths and URLs.
                document.MailMerge.MergeImageField += MergeImageFieldHandler;
                try
                {
                    // Merge all into the same document
                    document.MailMerge.Execute(dataTable);
                }
                finally
                {
                    document.MailMerge.MergeImageField -= MergeImageFieldHandler;
                }

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_mail_merged.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Mail merge executed successfully with {dataTable.Rows.Count} record(s) into document {outputKey}", new
                {
                    DocumentId = outputKey,
                    GroupName = dataTable.TableName,
                    RecordCount = dataTable.Rows.Count
                });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to execute mail merge: {ex.Message}");
            }
        }

        /// <summary>
        /// Resolves the input as either a file path or a raw JSON string.
        /// Returns content on success and sets out parameter `error` on failure — never throws.
        /// </summary>
        private string? ResolveJsonContent(string input, out string? error)
        {
            error = null;
            try
            {
                if (string.IsNullOrWhiteSpace(input))
                {
                    error = "Data source cannot be null or empty.";
                    return null;
                }

                string trimmed = input.Trim();

                // Detect file path: does not start with '{' or '[', and the file exists on disk
                if (!trimmed.StartsWith("{") && !trimmed.StartsWith("["))
                {
                    if (Mode == DocumentManagerMode.InMemory)
                    {
                        if (System.IO.File.Exists(trimmed))
                            return System.IO.File.ReadAllText(trimmed);
                    }
                    else
                    {
                        if (StorageManager!.HasDocument(trimmed))
                        {
                            using var stream = StorageManager.GetDocumentStream(trimmed);
                            if (stream != null)
                            {
                                using var reader = new StreamReader(stream, System.Text.Encoding.UTF8);
                                return reader.ReadToEnd();
                            }
                        }
                    }

                    error = $"Input does not appear to be valid JSON and the file path was not found: {trimmed}";
                    return null;
                }

                return trimmed;
            }
            catch (Exception ex)
            {
                error = $"Failed to resolve data source: {ex.Message}";
                return null;
            }
        }

        /// <summary>
        /// Parses JSON string to DataTable.
        /// Supported JSON formats:
        /// 1. Array of objects:  [{"Name":"John","Email":"john@example.com"}, ...]
        /// 2. Object wrapper:    { "TableName": [{"Name":"John",...}, ...] }
        /// 3. Columns/Rows:      { "Columns": ["Name","Email"], "Rows": [["John","john@example.com"],...] }
        /// </summary>
        private static DataTable ParseJsonToDataTable(string json)
        {
            try
            {
                using JsonDocument doc = JsonDocument.Parse(json);
                JsonElement root = doc.RootElement;

                // Format 1: Direct array of objects -> [{"Key":"Value"}, ...]
                if (root.ValueKind == JsonValueKind.Array)
                {
                    return ParseArrayOfObjects(root);
                }

                if (root.ValueKind == JsonValueKind.Object)
                {
                    // Format 3: Columns/Rows legacy format
                    if (root.TryGetProperty("Columns", out JsonElement columnsEl) &&
                        root.TryGetProperty("Rows", out JsonElement rowsEl))
                    {
                        var dataTable = new DataTable();
                        foreach (var col in columnsEl.EnumerateArray())
                            dataTable.Columns.Add(col.GetString());

                        foreach (var rowEl in rowsEl.EnumerateArray())
                        {
                            var rowValues = new List<object>();
                            foreach (var cell in rowEl.EnumerateArray())
                                rowValues.Add(GetJsonValue(cell));
                            dataTable.Rows.Add(rowValues.ToArray());
                        }
                        return dataTable;
                    }

                    // Format 2: Object wrapper -> { "TableName": [{...}, ...] }
                    foreach (JsonProperty property in root.EnumerateObject())
                    {
                        if (property.Value.ValueKind == JsonValueKind.Array)
                        {
                            DataTable dataTable = ParseArrayOfObjects(property.Value);
                            dataTable.TableName = property.Name;
                            return dataTable;
                        }
                    }
                }

                throw new ArgumentException("Unsupported JSON format. Expected an array of objects, an object wrapper, or a Columns/Rows structure.");
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Failed to parse JSON to DataTable: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Helper class for JSON deserialization.
        /// </summary>
        private class MailMergeData
        {
            public string[] Columns { get; set; } = Array.Empty<string>();
            public string[][] Rows { get; set; } = Array.Empty<string[]>();
        }
        #endregion

        #region ExecuteGroup
        /// <summary>
        /// Executes a mail merge operation in the Word document using a data table.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="dataTableJson">JSON string representing the data table used for mail merge.</param>
        /// <param name="groupName">The group name defined in the mail merge template.</param>
        /// <param name="removeEmptyFields">Specifies whether empty merge fields should be removed.</param>
        /// <param name="removeEmptyGroup">Specifies whether empty mail merge groups should be removed.</param>
        /// <param name="clearFields">Specifies whether unmerged merge fields should be removed after mail merge.</param>
        /// <param name="removeEmptyParagraphs">Specifies whether empty paragraphs created during mail merge should be removed.</param>
        /// <param name="startAtNewPage">Specifies whether each group record should start on a new page.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure of the mail merge operation, along with details of the generated document.</returns>
        [Tool(
            Name = "ExecuteGroupMailMerge",
            Description = "Executes a mail merge operation in the Word document using a DataTable. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Options: RemoveEmptyFields, RemoveEmptyGroup, RestartNumberingInLists, InsertAsNewRow.")]
        public AgentToolResult ExecuteGroupMailMerge(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "JSON representation of the data table with rows and columns")]
            string dataTableJson,
            [ToolParameter(Description = "The group name used in the mail merge template (e.g. 'Employees'). Required when the JSON is a plain array; ignored when the JSON object already contains a named property.")]
            string groupName = "",
            [ToolParameter(Description = "Whether to remove empty merge fields (true/false)")]
            bool removeEmptyFields = true,
            [ToolParameter(Description = "Whether to remove empty groups (true/false)")]
            bool removeEmptyGroup = true,
            [ToolParameter(Description = "Whether to remove unmerged merge fields from the document after mail merge. true = remove unmerged fields (default); false = keep unmerged fields as-is in the output document.")]
            bool clearFields = true,
            [ToolParameter(Description = "Whether to remove empty paragraphs that contain only a merge field with no data during mail merge. (true/false)")]
            bool removeEmptyParagraphs = true,
            [ToolParameter(Description = "Whether to start each group record on a new page. Valid only when group start and end are in the text body (not in tables, headers, or footers). (true/false)")]
            bool startAtNewPage = false,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                // Resolve: if dataSourceJson is a file path, read the file content
                string? jsonContent = ResolveJsonContent(dataTableJson, out string? resolveError);
                if (resolveError != null)
                    return AgentToolResult.Fail(resolveError);

                DataTable? dataTable = ParseJsonToDataTableForGroup(jsonContent, out string? parseError, groupName);
                if (parseError != null)
                    return AgentToolResult.Fail(parseError);

                document.MailMerge.RemoveEmptyParagraphs = removeEmptyParagraphs;
                document.MailMerge.RemoveEmptyGroup = removeEmptyGroup;
                document.MailMerge.StartAtNewPage = startAtNewPage;
                document.MailMerge.ClearFields = clearFields;

                // Subscribe image event to handle both local file paths and URLs.
                document.MailMerge.MergeImageField += MergeImageFieldHandler;
                try
                {
                    document.MailMerge.ExecuteGroup(dataTable);
                }
                finally
                {
                    document.MailMerge.MergeImageField -= MergeImageFieldHandler;
                }

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_group_mail_merged.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Group mail merge executed successfully with {dataTable.Rows.Count} record(s) into document {outputKey}", new
                {
                    DocumentId = outputKey,
                    GroupName = dataTable.TableName,
                    RecordCount = dataTable.Rows.Count
                });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to execute mail merge: {ex.Message}");
            }
        }
        // -------------------------------
        // GROUP JSON → DATATABLE
        // -------------------------------
        private static DataTable? ParseJsonToDataTableForGroup(string json, out string? error, string groupName = "")
        {
            error = null;
            try
            {
                using JsonDocument doc = JsonDocument.Parse(json);
                JsonElement root = doc.RootElement;

                if (root.ValueKind == JsonValueKind.Array)
                {
                    DataTable dataTable = ParseArrayOfObjects(root);
                    // TableName MUST be set for DocIO group mail merge to work
                    if (!string.IsNullOrEmpty(groupName))
                        dataTable.TableName = groupName;
                    else
                    {
                        error = "groupName parameter is required for array-type JSON in group mail merge.";
                        return null;
                    }

                    return dataTable;
                }
                if (root.ValueKind == JsonValueKind.Object)
                {
                    foreach (JsonProperty property in root.EnumerateObject())
                    {
                        if (property.Value.ValueKind == JsonValueKind.Array)
                        {
                            DataTable dataTable = ParseArrayOfObjects(property.Value);
                            dataTable.TableName = property.Name; // e.g., "Customers"
                            return dataTable;
                        }
                    }
                }

                error = "Group JSON must be an array of objects";
                return null;
            }
            catch (Exception ex)
            {
                error = $"Failed to parse group JSON: {ex.Message}";
                return null;
            }
        }

        /// <summary>Converts a JSON array of objects into a DataTable.</summary>
        private static DataTable ParseArrayOfObjects(JsonElement array)
        {
            DataTable table = new DataTable();
            HashSet<string> columns = new HashSet<string>();

            foreach (var item in array.EnumerateArray())
            {
                foreach (var prop in item.EnumerateObject())
                {
                    if (prop.Value.ValueKind != JsonValueKind.Object &&
                        prop.Value.ValueKind != JsonValueKind.Array)
                    {
                        if (!table.Columns.Contains(prop.Name))
                            table.Columns.Add(prop.Name);
                    }
                }
            }

            foreach (var item in array.EnumerateArray())
            {
                DataRow row = table.NewRow();
                foreach (var prop in item.EnumerateObject())
                {
                    if (table.Columns.Contains(prop.Name))
                        row[prop.Name] = GetJsonValue(prop.Value);
                }
                table.Rows.Add(row);
            }

            return table;
        }

        /// <summary>Extracts the string representation of a JsonElement value.</summary>
        private static string GetJsonValue(JsonElement value)
        {
            return value.ValueKind switch
            {
                JsonValueKind.String => value.GetString() ?? "",
                JsonValueKind.Number => value.GetRawText(),
                JsonValueKind.True => "true",
                JsonValueKind.False => "false",
                _ => ""
            };
        }
        #endregion ExecuteGroupMailMerge

        #region ExecuteNestedGroup
        /// <summary>
        /// Executes a mail merge operation in the Word document using a data table.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="dataTableJson">JSON string representing the data table used for nested group mail merge.</param>
        /// <param name="groupName">The group name defined in the mail merge template.</param>
        /// <param name="removeEmptyFields">Specifies whether empty merge fields should be removed.</param>
        /// <param name="removeEmptyGroup">Specifies whether empty groups should be removed.</param>
        /// <param name="clearFields">Specifies whether unmerged merge fields should be removed after mail merge.</param>
        /// <param name="removeEmptyParagraphs">Specifies whether empty paragraphs created during mail merge should be removed.</param>
        /// <param name="startAtNewPage">Specifies whether each group record should start on a new page.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating whether the nested group mail merge operation succeeded.</returns>
        [Tool(
            Name = "ExecuteNestedGroupMailMerge",
            Description = "Executes a nested group mail merge operation in the Word document using a DataTable. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Options: RemoveEmptyFields, RemoveEmptyGroup, RestartNumberingInLists, InsertAsNewRow.")]
        public AgentToolResult ExecuteNestedGroupMailMerge(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "JSON representation of the data table with rows and columns")]
            string dataTableJson,
            [ToolParameter(Description = "The group name used in the mail merge template (e.g. 'Employees'). Required when the JSON is a plain array; ignored when the JSON object already contains a named property.")]
            string groupName = "",
            [ToolParameter(Description = "Whether to remove empty merge fields (true/false)")]
            bool removeEmptyFields = true,
            [ToolParameter(Description = "Whether to remove empty groups (true/false)")]
            bool removeEmptyGroup = true,
            [ToolParameter(Description = "Whether to remove unmerged merge fields from the document after mail merge. true = remove unmerged fields (default); false = keep unmerged fields as-is in the output document.")]
            bool clearFields = true,
            [ToolParameter(Description = "Whether to remove empty paragraphs that contain only a merge field with no data during mail merge. (true/false)")]
            bool removeEmptyParagraphs = true,
            [ToolParameter(Description = "Whether to start each group record on a new page. Valid only when group start and end are in the text body (not in tables, headers, or footers). (true/false)")]
            bool startAtNewPage = false,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                // Resolve: if dataSourceJson is a file path, read the file content
                string? jsonContent = ResolveJsonContent(dataTableJson, out string? resolveError);
                if (resolveError != null)
                    return AgentToolResult.Fail(resolveError);

                // Resolve masterTableName from JSON if it's an object wrapper
                string resolvedTableName = groupName;
                using (var doc = System.Text.Json.JsonDocument.Parse(jsonContent))
                {
                    var root = doc.RootElement;
                    if (root.ValueKind == System.Text.Json.JsonValueKind.Object)
                    {
                        foreach (var property in root.EnumerateObject())
                        {
                            if (property.Value.ValueKind == System.Text.Json.JsonValueKind.Array)
                            {
                                // Auto-resolve table name from JSON key e.g., "Organizations"
                                resolvedTableName = property.Name;
                                break;
                            }
                        }
                    }
                }

                // Parse nested JSON to List<object>
                var jsonData = ParseNestedJson(jsonContent);

                document.MailMerge.RemoveEmptyParagraphs = removeEmptyParagraphs;
                document.MailMerge.StartAtNewPage = startAtNewPage;
                document.MailMerge.ClearFields = clearFields;

                // Create MailMergeDataTable for nested merge
                var mailMergeDataTable = new MailMergeDataTable(resolvedTableName, jsonData);

                // Subscribe image event to handle both local file paths and URLs.
                document.MailMerge.MergeImageField += MergeImageFieldHandler;
                try
                {
                    document.MailMerge.ExecuteNestedGroup(mailMergeDataTable);
                }
                finally
                {
                    document.MailMerge.MergeImageField -= MergeImageFieldHandler;
                }

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_nested_group_mail_merged.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Nested group mail merge executed successfully with {jsonData.Count} record(s) into document {outputKey}", new
                {
                    DocumentId = outputKey,
                    GroupName = resolvedTableName,
                    RecordCount = jsonData.Count
                });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to execute mail merge: {ex.Message}");
            }
        }
        // -------------------------------
        // NESTED JSON → OBJECT GRAPH
        // -------------------------------
        private static List<object> ParseNestedJson(string json)
        {
            try
            {
                var result = new List<object>();

                using var doc = System.Text.Json.JsonDocument.Parse(json);
                var root = doc.RootElement;

                // Case 1: Direct array -> [{...}, {...}]
                if (root.ValueKind == System.Text.Json.JsonValueKind.Array)
                {
                    foreach (var item in root.EnumerateArray())
                    {
                        var record = ConvertJsonElement(item);
                        if (record != null)
                            result.Add(record);
                    }
                    return result;
                }

                // Case 2: Object with nested array property -> { "Organizations": [{...}] }
                if (root.ValueKind == System.Text.Json.JsonValueKind.Object)
                {
                    foreach (var property in root.EnumerateObject())
                    {
                        if (property.Value.ValueKind == System.Text.Json.JsonValueKind.Array)
                        {
                            foreach (var item in property.Value.EnumerateArray())
                            {
                                var record = ConvertJsonElement(item);
                                if (record != null)
                                    result.Add(record);
                            }
                            return result;
                        }
                    }
                }

                throw new ArgumentException("JSON must be a direct array or an object containing an array property.");
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Failed to parse nested JSON: {ex.Message}", ex);
            }
        }

        /// <summary>Converts a JsonElement into dictionaries, lists, or primitive .NET values.</summary>
        private static object ConvertJsonElement(JsonElement element)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    var dict = new Dictionary<string, object>();
                    foreach (var p in element.EnumerateObject())
                        dict[p.Name] = ConvertJsonElement(p.Value);
                    return dict;

                case JsonValueKind.Array:
                    var list = new List<object>();
                    foreach (var v in element.EnumerateArray())
                        list.Add(ConvertJsonElement(v));
                    return list;

                case JsonValueKind.String:
                    return element.GetString() ?? "";

                case JsonValueKind.Number:
                    return element.GetRawText();

                case JsonValueKind.True:
                    return true;

                case JsonValueKind.False:
                    return false;

                default:
                    return "";
            }
        }
        #endregion

    }
}
