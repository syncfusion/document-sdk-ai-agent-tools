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
using System.Text.RegularExpressions;

namespace Syncfusion.AI.AgentTools.Word
{
    /// <summary>
    /// Provides agent tools for text search and replacement operations in Word documents.
    /// Handles finding and replacing text with support for case sensitivity and whole word matching.
    /// </summary>
    public class WordFindAndReplaceAgentTools : AgentToolBase<WordDocument>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WordFindAndReplaceAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The document manager for managing Word documents.</param>
        public WordFindAndReplaceAgentTools(WordDocumentManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordFindAndReplaceAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public WordFindAndReplaceAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Finds and replaces all occurrences of one or more specified texts in the Word document.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="findTexts">Array of texts to search for.</param>
        /// <param name="replaceTexts">Array of replacement texts corresponding to each find text.</param>
        /// <param name="matchCase">Indicates whether the search should be case-sensitive.</param>
        /// <param name="wholeWord">Indicates whether only whole words should be matched.</param>
        /// <param name="replaceFirst">Indicates whether to replace only the first occurrence.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating whether the find-and-replace operation succeeded or failed.</returns>
        [Tool(
            Name = "FindAndReplace",
            Description = "Finds all occurrences of one or more specified texts in the Word document and replaces each with the corresponding replacement text in a single pass on the same document. Accepts arrays so multiple placeholders can be replaced at once without reopening the file. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult FindAndReplace(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "Array of texts to search for")]
            string[] findTexts,
            [ToolParameter(Description = "Array of replacement texts corresponding to each find text")]
            string[] replaceTexts,
            [ToolParameter(Description = "Whether to match case (true/false)")]
            bool matchCase = false,
            [ToolParameter(Description = "Whether to match whole words only (true/false)")]
            bool wholeWord = false,
            [ToolParameter(Description = "Whether to find and replace first occurrence or all occurrences (true/false)")]
            bool replaceFirst = false,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
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

                // ── Open ──────────────────────────────────────────────────────────
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                var replacementSummary = new Dictionary<string, int>();

                // Apply all find-replace pairs on the same document instance
                for (int i = 0; i < findTexts.Length; i++)
                {
                    string findWhat = findTexts[i];
                    string replaceWith = replaceTexts[i];

                    document.ReplaceFirst = replaceFirst;
                    int count = document.Replace(findWhat, replaceWith, matchCase, wholeWord);
                    replacementSummary[findWhat] = count;
                }

                // ── Save once after all replacements ────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_replaced.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

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
                return AgentToolResult.Fail($"Failed to replace text: {ex.Message}");
            }
        }

        /// <summary>
        /// Finds and replaces occurrences of the specified regex pattern in the Word document.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="findWhat">The regular expression pattern used to match text (for example, '{[A-Za-z]+}' to match placeholders like {Name} or {Date}).</param>
        /// <param name="replaceText">The text used to replace the matched content.</param>
        /// <param name="replaceFirst">Specifies whether to replace only the first match.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating the number of replacements performed.</returns>
        [Tool(
            Name = "FindAndReplaceByPattern",
            Description = "Finds and replaces first occurrence of the specified pattern in the Word document. Returns a count of replaced text.")]
        public AgentToolResult FindAndReplaceByPattern(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The regex pattern to match text (e.g., '{[A-Za-z]+}' to match placeholders like {Name}, {Date})")]
            string findWhat,
            [ToolParameter(Description = "The replacement text")]
            string replaceText,
            [ToolParameter(Description = "Whether to find and replace first occurence or all occurence (true/false)")]
            bool replaceFirst = false,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                // ── Open ────────────────────────────────────────────────────────
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                // ── Core logic ──────────────────────────────────
                Regex regex = new Regex(findWhat);
                document.ReplaceFirst = replaceFirst;
                int replacedCount = document.ReplaceSingleLine(regex, replaceText);

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_replaced.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Replaced {replacedCount} occurrence(s) into document {outputKey}", new { ReplacedCount = replacedCount });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to replace all text: {ex.Message}");
            }
        }
    }
}
