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
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;

namespace Syncfusion.AI.AgentTools.Word
{
    /// <summary>
    /// Provides agent tools for core document lifecycle management and file I/O operations.
    /// Handles document creation, loading, export, and basic manager.
    /// </summary>
    public class WordDocumentAgentTools : AgentToolBase
    {
        private readonly WordDocumentManager _manager;
        private readonly string _outputDirectory;

        /// <summary>
        /// Initializes a new instance of the <see cref="WordDocumentAgentTools"/> class.
        /// </summary>
        /// <param name="manager">The document manager for managing Word documents.</param>
        /// <param name="outputDirectory">The default output directory for exporting documents. If null or empty, uses current directory.</param>
        public WordDocumentAgentTools(WordDocumentManager manager, string? outputDirectory = null)
        {
            ArgumentNullException.ThrowIfNull(manager);
            
            _manager = manager;
            _outputDirectory = string.IsNullOrEmpty(outputDirectory) 
                ? Environment.CurrentDirectory 
                : outputDirectory;

            // Ensure output directory exists
            if (!Directory.Exists(_outputDirectory))
                Directory.CreateDirectory(_outputDirectory);
        }

        /// <summary>
        /// Creates a new Word document in the manager or loads from file.
        /// </summary>
        /// <param name="filePath">Optional file path to load an existing document. If null, creates an empty document.</param>
        /// <param name="password">Optional password if the document is encrypted.</param>
        /// <returns>Result containing the document ID.</returns>
        [Tool(
            Name = "CreateDocument",
            Description = "Create or load a Word document in memory. filePath - Use the path to create a WordDocument instance; if it is null, create an empty document. password - If the document is encrypted, provide the password. Returns the documentid of a newly created document.")]
        public AgentToolResult CreateDocument(
            [ToolParameter(Description = "Use the path to create a WordDocument instance; if it is null, create an empty document")]
            string? filePath = null,
            [ToolParameter(Description = "If the document is encrypted, provide the password")]
            string? password = null)
        {
            try
            {
                WordDocument document;
                string documentId;

                if (string.IsNullOrEmpty(filePath))
                {
                    document = _manager.CreateDocument();
                    documentId = _manager.ActiveDocumentId!;
                }
                else
                {
                    // Load existing word document
                    if (!string.IsNullOrEmpty(password))
                    {
                        document = _manager.ImportDocument(filePath, password);
                    }
                    else
                    {
                        document = _manager.ImportDocument(filePath);
                    }
                    documentId = _manager.ActiveDocumentId!;
                }

                return AgentToolResult.Ok(
                    $"Document created/loaded successfully with ID: {documentId}",
                    new { DocumentId = documentId });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to create document: {ex.Message}");
            }
        }

        /// <summary>
        /// Returns all document IDs in memory.
        /// </summary>
        /// <returns>Result containing array of document IDs.</returns>
        [Tool(
            Name = "GetAllDocuments",
            Description = "Returns all documentid of the Word document instance that are available in main memory.")]
        public AgentToolResult GetAllDocuments()
        {
            try
            {
                var documentIds = _manager.GetAllDocumentIds();

                return AgentToolResult.Ok(
                    $"Found {documentIds.Count} document(s) in memory",
                    new { DocumentIds = documentIds.ToArray(), Count = documentIds.Count });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to retrieve documents: {ex.Message}");
            }
        }

        /// <summary>
        /// Exports the document to the file system in the specified format.
        /// </summary>
        /// <param name="documentId">The ID of the document to be exported.</param>
        /// <param name="filePath">The file name or full path where the exported document will be saved</param>
        /// <param name="formatType">The export format: 'Docx', 'Doc', 'Rtf', 'Html', 'Md' or 'Txt'. Default is 'Docx'.</param>
        /// <returns>Result indicating whether the export operation succeeded or failed.</returns>
        [Tool(
            Name = "ExportDocument",
            Description = "Exports the document to the file system in the specified format. Supported formats: DOCX, DOC, RTF, HTML, TXT, MD.")]
        public AgentToolResult ExportDocument(
            [ToolParameter(Description = "The ID of the document to export")]
            string documentId,
            [ToolParameter(Description = "The file name or full path to export to. If only filename is provided, uses the output directory.")]
            string filePath,
            [ToolParameter(Description = "The format: Docx, Doc, Rtf, Html, Txt, Md")]
            string? formatType = "Docx")
        {
            try
            {
                // Determine the full path
                string fullPath;
                if (Path.IsPathRooted(filePath))
                {
                    fullPath = filePath;
                }
                else
                {
                    fullPath = Path.Combine(_outputDirectory, filePath);
                }

                _manager.ExportDocument(fullPath, documentId);

                return AgentToolResult.Ok($"Document {documentId} exported successfully to {fullPath}", new { FilePath = fullPath });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to export document: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes a document from memory.
        /// </summary>
        /// <param name="documentId">The ID of the document to be removed.</param>
        /// <returns>Result indicating whether the document was removed successfully or if it was not found.</returns>
        [Tool(
            Name = "RemoveDocument",
            Description = "Removes a specific document from memory by ID. Returns true if the document was removed, false otherwise.")]
        public AgentToolResult RemoveDocument(
            [ToolParameter(Description = "The ID of the document to remove")]
            string documentId)
        {
            try
            {
                bool removed = _manager.RemoveDocument(documentId);

                if (removed)
                {
                    return AgentToolResult.Ok($"Document {documentId} removed successfully from memory");
                }
                else
                {
                    return AgentToolResult.Fail($"Document not found: {documentId}");
                }
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to remove document: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets the active document context.
        /// </summary>
        /// <param name="documentId">The ID of the document to set as active.</param>
        /// <returns>Result indicating whether the active document was set successfully or if it was not found.</returns>
        [Tool(
            Name = "SetActiveDocument",
            Description = "Changes the active document context by ID. The active document is used as the default when no document ID is specified in other operations.")]
        public AgentToolResult SetActiveDocument(
            [ToolParameter(Description = "The ID of the document to set as active")]
            string documentId)
        {
            try
            {
                _manager.SetActiveDocument(documentId);

                return AgentToolResult.Ok($"Document {documentId} set as active document");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to set active document: {ex.Message}");
            }
        }

      
    }
}
