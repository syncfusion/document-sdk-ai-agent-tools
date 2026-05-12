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

namespace Syncfusion.AI.AgentTools.Word
{
    /// <summary>
    /// Provides agent tools for managing tracked changes (revisions) in Word documents.
    /// Handles accepting, rejecting, and querying document revisions.
    /// </summary>
    public class WordRevisionAgentTools : AgentToolBase<WordDocument>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WordRevisionAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The document manager for managing Word documents.</param>
        public WordRevisionAgentTools(WordDocumentManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordRevisionAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public WordRevisionAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Accepts all revisions made by a specific author.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document containing the TOC to update.</param>
        /// <param name="author">The name of the author whose tracked revisions should be accepted.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating whether the revision acceptance was completed successfully.</returns>
        [Tool(
            Name = "AcceptRevisionsByAuthor",
            Description = "Accepts all tracked change revisions made by a specific author in the Word document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult AcceptRevisionsByAuthor(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The name of the author whose revisions should be accepted")]
            string author,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                int acceptedCount = 0;
                //Iterate revisions in reverse order to safely accept while modifying the collection.
                for (int i = document.Revisions.Count - 1; i >= 0; i--)
                {
                    //Check the author of the current revision and accept it.
                    if (document.Revisions[i].Author == author)
                    {
                        document.Revisions[i].Accept();
                        acceptedCount++;
                    }
                    //Reset to last item when accepting moving-related revisions
                    //which can shift the collection size.
                    if (i > document.Revisions.Count - 1)
                        i = document.Revisions.Count;
                }

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_revisions_accepted.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                if (acceptedCount == 0)
                    return AgentToolResult.Ok($"No revisions found for author '{author}'", new { AcceptedCount = 0, Author = author, DocumentId = outputKey });

                return AgentToolResult.Ok($"Accepted {acceptedCount} revision(s) by '{author}' into document {outputKey}", new { AcceptedCount = acceptedCount, Author = author, DocumentId = outputKey });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to accept revisions: {ex.Message}");
            }
        }

        /// <summary>
        /// Rejects all revisions made by a specific author.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="author">The name of the author whose tracked revisions should be rejected.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating whether the revision rejection was completed successfully.</returns>
        [Tool(
            Name = "RejectRevision",
            Description = "Rejects all tracked change revisions made by a specific author in the Word document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult RejectRevision(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The name of the author whose revisions should be rejected")]
            string author,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                int rejectedCount = 0;
                //Iterate revisions in reverse order to safely reject while modifying the collection.
                for (int i = document.Revisions.Count - 1; i >= 0; i--)
                {
                    //Check the author of the current revision and reject it.
                    if (document.Revisions[i].Author == author)
                    {
                        document.Revisions[i].Reject();
                        rejectedCount++;
                    }
                    //Reset to last item when rejecting moving-related revisions
                    //which can shift the collection size.
                    if (i > document.Revisions.Count - 1)
                        i = document.Revisions.Count;
                }

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_revisions_rejected.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                if (rejectedCount == 0)
                    return AgentToolResult.Ok($"No revisions found for author '{author}'", new { RejectedCount = 0, Author = author, DocumentId = outputKey });

                return AgentToolResult.Ok($"Rejected {rejectedCount} revision(s) by '{author}' into document {outputKey}", new { RejectedCount = rejectedCount, Author = author, DocumentId = outputKey });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to reject revisions: {ex.Message}");
            }
        }

        /// <summary>
        /// Accepts all revisions in a Word document.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document containing the TOC to update.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating whether the revision acceptance was completed successfully.</returns>
        [Tool(
            Name = "AcceptAllRevisions",
            Description = "Accepts all revisions in the document and returns count.")]
        public AgentToolResult AcceptAllRevisions(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                int count = document.Revisions.Count;
                document.Revisions.AcceptAll();

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_all_accepted.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Accepted all {count} revision(s) into document {outputKey}", new { AcceptedCount = count, DocumentId = outputKey });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to accept all revisions: {ex.Message}");
            }
        }

        /// <summary>
        /// Rejects all revisions in a Word document.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document containing the TOC to update.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating whether the revision rejection was completed successfully.</returns>
        [Tool(
            Name = "RejectAllRevisions",
            Description = "Rejects all revisions in the document and returns count.")]
        public AgentToolResult RejectAllRevisions(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                int count = document.Revisions.Count;
                document.Revisions.RejectAll();

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_all_rejected.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Rejected all {count} revision(s) into document {outputKey}", new { RejectedCount = count, DocumentId = outputKey });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to reject all revisions: {ex.Message}");
            }
        }
    }
}
