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

using System;
using System.Text.RegularExpressions;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.Presentation;

namespace Syncfusion.AI.AgentTools.PowerPoint
{
    /// <summary>
    /// Provides AI agent tools for PowerPoint presentation find and replace operations.
    /// Uses the Syncfusion Presentation library's built-in FindAll API for
    /// reliable text searching and replacement across all slide elements.
    /// </summary>
    public class PresentationFindAndReplaceAgentTools : AgentToolBase<IPresentation>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationFindAndReplaceAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The presentation manager for managing PowerPoint presentations.</param>
        public PresentationFindAndReplaceAgentTools(PresentationManager manager)
            : base(manager, DocumentType.PowerPoint) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationFindAndReplaceAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public PresentationFindAndReplaceAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.PowerPoint) { }

        /// <summary>
        /// Finds all occurrences of one or more specified texts in the presentation and replaces each with the corresponding replacement text in a single pass.
        /// </summary>
        /// <param name="documentIdOrFilePath">The ID of the presentation or input file path.</param>
        /// <param name="findTexts">The array of texts to find.</param>
        /// <param name="replaceTexts">The array of replacement texts corresponding to each find text.</param>
        /// <param name="matchCase">Whether to match case.</param>
        /// <param name="wholeWord">Whether to match whole words only.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the count of replaced occurrences per find text.</returns>
        [Tool(Name = "FindAndReplace", Description = "Finds all occurrences of one or more specified texts in the PowerPoint presentation and replaces each with the corresponding replacement text in a single pass on the same document. Accepts arrays so multiple placeholders can be replaced at once without reopening the file. Searches across all slides and elements such as shapes, textboxes, tables, SmartArt, etc. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult FindAndReplace(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath,
            [ToolParameter(Description = "Array of texts to find ")] string[] findTexts,
            [ToolParameter(Description = "Array of replacement texts corresponding to each find text")] string[] replaceTexts,
            [ToolParameter(Description = "Whether to match case")] bool matchCase = false,
            [ToolParameter(Description = "Whether to match whole words only")] bool wholeWord = false,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);
                ArgumentNullException.ThrowIfNull(findTexts);
                ArgumentNullException.ThrowIfNull(replaceTexts);

                if (findTexts.Length != replaceTexts.Length)
                    return AgentToolResult.Fail($"findTexts and replaceTexts arrays must have the same length. Got {findTexts.Length} find texts and {replaceTexts.Length} replace texts.");

                if (findTexts.Length == 0)
                    return AgentToolResult.Fail("findTexts array must not be empty.");

                var presentation = OpenDocument(documentIdOrFilePath);
                if (presentation == null)
                    return AgentToolResult.Fail($"Presentation not found: {documentIdOrFilePath}");

                var replacementSummary = new System.Collections.Generic.Dictionary<string, int>();

                // Apply all find-replace pairs on the same document instance
                for (int i = 0; i < findTexts.Length; i++)
                {
                    string findWhat = findTexts[i];
                    string replaceWith = replaceTexts[i];

                    ITextSelection[] textSelections = presentation.FindAll(findWhat, matchCase, wholeWord);

                    int count = 0;
                    if (textSelections != null && textSelections.Length > 0)
                    {
                        foreach (ITextSelection textSelection in textSelections)
                        {
                            ITextPart textPart = textSelection.GetAsOneTextPart();
                            textPart.Text = replaceWith;
                        }
                        count = textSelections.Length;
                    }

                    replacementSummary[findWhat] = count;
                }

                // ── Save once after all replacements ────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage) 
                    outputFilePath = "output_replaced.pptx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, presentation);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath;

                int totalReplaced = 0;
                var summaryLines = new System.Text.StringBuilder();
                foreach (var kvp in replacementSummary)
                {
                    totalReplaced += kvp.Value;
                    summaryLines.AppendLine($"  '{kvp.Key}': {kvp.Value} occurrence(s) replaced");
                }

                return AgentToolResult.Ok(
                    $"Completed {findTexts.Length} find-and-replace operation(s) with {totalReplaced} total replacement(s) into document '{outputKey}':\n{summaryLines}",
                    new { TotalReplacedCount = totalReplaced, ReplacementSummary = replacementSummary });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to find and replace text in PowerPoint presentation: {ex.Message}");
            }
        }

        /// <summary>
        /// Finds all occurrences of text matching the specified regex pattern in the presentation and replaces them with the given replacement text.
        /// </summary>
        /// <param name="documentId">The ID of the presentation.</param>
        /// <param name="regexPattern">The regex pattern to match text.</param>
        /// <param name="replaceText">The replacement text.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the count of replaced occurrences.</returns>
        [Tool(Name = "FindAndReplaceByPattern", Description = "Finds all occurrences of text matching the specified regex pattern in the PowerPoint presentation and replaces them with the given replacement text. Searches across all slides and elements such as shapes, textboxes, tables, SmartArt, etc. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult FindAndReplaceByPattern(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath,
            [ToolParameter(Description = "The regex pattern to match text (e.g., '{[A-Za-z]+}' to match placeholders like {Name}, {Date})")] string regexPattern,
            [ToolParameter(Description = "The replacement text")] string replaceText,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);
                ArgumentNullException.ThrowIfNull(regexPattern);
                ArgumentNullException.ThrowIfNull(replaceText);

                var presentation = OpenDocument(documentIdOrFilePath);
                if (presentation == null)
                    return AgentToolResult.Fail($"Presentation not found: {documentIdOrFilePath}");

                // Find all occurrences matching the regex pattern using the Syncfusion Presentation API
                Regex pattern = new Regex(regexPattern);
                ITextSelection[] textSelections = presentation.FindAll(pattern);

                if (textSelections == null || textSelections.Length == 0)
                {
                    return AgentToolResult.Ok($"No occurrence matching pattern '{regexPattern}' found", new { ReplacedCount = 0 });
                }

                // Replace each found occurrence
                foreach (ITextSelection textSelection in textSelections)
                {
                    ITextPart textPart = textSelection.GetAsOneTextPart();
                    textPart.Text = replaceText;
                }

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_pattern_replaced.pptx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, presentation);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Replaced {textSelections.Length} occurrence(s) matching pattern '{regexPattern}' with '{replaceText}' into document {outputKey}",
                    new { ReplacedCount = textSelections.Length });
            }
            catch (RegexParseException ex)
            {
                return AgentToolResult.Fail($"Invalid regex pattern '{regexPattern}': {ex.Message}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to find and replace text by pattern in PowerPoint presentation: {ex.Message}");
            }
        }
    }
}
