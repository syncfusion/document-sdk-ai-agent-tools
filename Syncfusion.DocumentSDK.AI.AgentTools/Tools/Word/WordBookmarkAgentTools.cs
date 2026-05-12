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

namespace Syncfusion.AI.AgentTools.Word
{
    /// <summary>
    /// Provides agent tools for bookmark management in Word documents.
    /// Handles bookmark operations including content extraction, replacement, and deletion.
    /// </summary>
    public class WordBookmarkAgentTools : AgentToolBase<WordDocument>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WordBookmarkAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The document manager for managing Word documents.</param>
        public WordBookmarkAgentTools(WordDocumentManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordBookmarkAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public WordBookmarkAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Gets all bookmark names from the Word document.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or the input presentation file path (DocumentStorage mode).</param>
        /// <returns>A collection of all bookmark names found in the document.</returns>
        [Tool(
            Name = "GetBookmarks",
            Description = "Gets all bookmark names from the Word document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Returns the collection of all bookmark names.")]
        public AgentToolResult GetBookmarks(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                var bookmarkNames = new string[document.Bookmarks.Count];
                for (int i = 0; i < document.Bookmarks.Count; i++)
                {
                    bookmarkNames[i] = document.Bookmarks[i].Name;
                }

                string bookmarkList = bookmarkNames.Length > 0
                    ? string.Join(", ", bookmarkNames)
                    : "No bookmarks found";

                return AgentToolResult.Ok($"Found {bookmarkNames.Length} bookmark(s): {bookmarkList}", new
                {
                    Count = bookmarkNames.Length,
                    BookmarkNames = bookmarkNames
                });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to get bookmarks: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets content from the specified bookmark and creates a new document with that content.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or the input presentation file path (DocumentStorage mode).</param>
        /// <param name="bookmarkName">The name of the bookmark from which content will be extracted.</param>
        /// <returns>Result containing the document ID of the newly created document with the extracted bookmark content.</returns>
        [Tool(
            Name = "GetContent",
            Description = "Get content from the bookmark. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Returns the document id of new document created for bookmark content.")]
        public AgentToolResult GetContent(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The name of the bookmark")]
            string bookmarkName)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                // Navigate to bookmark and copy content
                BookmarksNavigator navigator = new BookmarksNavigator(document);
                navigator.MoveToBookmark(bookmarkName);
                
                // Get the bookmark content as Word document and add to memory
                WordDocumentPart bookmarkContent = navigator.GetContent();
                WordDocument newDocument = bookmarkContent.GetAsWordDocument();
                
                string newDocumentId;
                if (Mode == DocumentManagerMode.InMemory)
                {
                    ((WordDocumentManager)InMemoryManager!).AddDocument(newDocument);
                    newDocumentId = ((WordDocumentManager)InMemoryManager).ActiveDocumentId!;
                }
                else
                {
                    // For storage mode, generate a unique ID for the new document
                    newDocumentId = $"bookmark_content_{bookmarkName}_{Guid.NewGuid():N}.docx";
                    SaveDocument(newDocumentId, newDocument);
                }

                return AgentToolResult.Ok($"Bookmark content extracted to new document with ID: {newDocumentId}", new
                {
                    NewDocumentId = newDocumentId,
                });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to get bookmark content: {ex.Message}");
            }
        }

        /// <summary>
        /// Replaces the existing bookmark content with content from another document.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or the input presentation file path (DocumentStorage mode).</param>
        /// <param name="bookmarkName">The name of the bookmark whose content will be replaced.</param>
        /// <param name="replaceDocumentIdOrFilePath">The document ID or file path containing the replacement content.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure of the bookmark content replacement operation.</returns>
        [Tool(
            Name = "ReplaceContent",
            Description = "Replaces the existing bookmark content with content from another document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ReplaceContent(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) containing the bookmark")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The name of the bookmark")]
            string bookmarkName,
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) to replace with")]
            string replaceDocumentIdOrFilePath,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                var replaceDocument = OpenDocument(replaceDocumentIdOrFilePath);
                if (replaceDocument == null)
                    return AgentToolResult.Fail($"Replace document not found: {replaceDocumentIdOrFilePath}");

                // Navigate to bookmark
                BookmarksNavigator navigator = new BookmarksNavigator(document);
                navigator.MoveToBookmark(bookmarkName);

                // Replace bookmark content
                WordDocumentPart replacementContent = new WordDocumentPart(replaceDocument);
                navigator.ReplaceContent(replacementContent);

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_bookmark_replaced.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Bookmark '{bookmarkName}' content replaced successfully into document {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to replace bookmark content: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes the content of the specified bookmark in the Word document.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or the input presentation file path (DocumentStorage mode).</param>
        /// <param name="bookmarkName"> The bookmark for which content will be removed.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure of the bookmark content removal operation.</returns>
        [Tool(
            Name = "RemoveContent",
            Description = "Removes the content of the specified bookmark in the Word document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult RemoveContent(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The name of the bookmark")]
            string bookmarkName,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                // Navigate to bookmark and delete content
                BookmarksNavigator navigator = new BookmarksNavigator(document);
                navigator.MoveToBookmark(bookmarkName);
                navigator.DeleteBookmarkContent(false);

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_bookmark_content_removed.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Content of bookmark '{bookmarkName}' removed successfully into document {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to remove bookmark content: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes the specified bookmark from the Word document.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or the input presentation file path (DocumentStorage mode).</param>
        /// <param name="bookmarkName">The name of the bookmark to be removed from the document.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure of the bookmark removal operation.</returns>
        [Tool(
            Name = "RemoveBookmark",
            Description = "Removes the specified bookmark from the Word document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult RemoveBookmark(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The name of the bookmark")]
            string bookmarkName,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                // Check if bookmark exists
                Bookmark bookmark = document.Bookmarks.FindByName(bookmarkName);
                if (bookmark == null)
                    return AgentToolResult.Fail($"Bookmark not found: {bookmarkName}");

                // Remove the bookmark
               document.Bookmarks.Remove(bookmark);

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_bookmark_removed.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Bookmark '{bookmarkName}' removed successfully into document {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to remove bookmark: {ex.Message}");
            }
        }
    }
}
