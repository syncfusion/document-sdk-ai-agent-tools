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
using System.IO;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.Pdf;

namespace Syncfusion.AI.AgentTools.PDF
{
    /// <summary>
    /// Provides AI agent tools for PDF document lifecycle management and I/O operations.
    /// Handles document creation, import, export, and manager.
    /// </summary>
    public class PdfDocumentAgentTools : AgentToolBase
    {
        private readonly PdfDocumentManager _manager;
        private readonly string? _outputDirectory;

        /// <summary>
        /// Initializes a new instance of the <see cref="PdfDocumentAgentTools"/> class.
        /// </summary>
        /// <param name="manager">The PDF document manager.</param>
        /// <param name="outputDirectory">Optional output directory for file operations.</param>
        public PdfDocumentAgentTools(PdfDocumentManager manager, string? outputDirectory = null)
        {
            _manager = manager ?? throw new ArgumentNullException(nameof(manager));
            _outputDirectory = outputDirectory;
        }

        /// <summary>
        /// Creates a PDF document in memory or loads an existing PDF from a file.
        /// </summary>
        /// <param name="filePath">Optional path to an existing PDF file. If null, creates a new empty PDF document.</param>
        /// <param name="password">Optional password for encrypted PDF files.</param>
        /// <returns>Result containing the document ID of the created or loaded PDF document.</returns>
        [Tool(Name = "CreatePdfDocument", Description = "Create a PDF file in memory. filePath: Path of an existing PDF to load. If null, creates a new PdfDocument. password: Password for encrypted PDF. Returns the documentId.")]
        public AgentToolResult CreatePdfDocument(
            [ToolParameter(Description = "Path of an existing PDF to load, or null to create a new document")] string? filePath = null,
            [ToolParameter(Description = "Password for encrypted PDF")] string? password = null)
        {
            try
            {
                string documentId;
                if (string.IsNullOrEmpty(filePath))
                {
                    // Create new empty PDF document
                    var document = _manager.CreateDocument();
                    documentId = _manager.ActiveDocumentId ?? throw new InvalidOperationException("Failed to create document");
                }
                else
                {
                    // Load existing PDF document
                    if (!string.IsNullOrEmpty(password))
                    {
                        _manager.ImportDocument(filePath, password);
                    }
                    else
                    {
                        _manager.ImportDocument(filePath);
                    }
                    documentId = _manager.ActiveDocumentId ?? throw new InvalidOperationException("Failed to import document");
                }
                
                return AgentToolResult.Ok(
                    $"PDF document created/loaded successfully with ID: {documentId}",
                    new { DocumentId = documentId });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to create/load PDF document: {ex.Message}");
            }
        }

        /// <summary>
        /// Retrieves all PDF document IDs currently in memory.
        /// </summary>
        /// <returns>Result containing array of all PDF document IDs.</returns>
        [Tool(Name = "GetAllPDFDocuments", Description = "Returns all documentid of the PDF document instance that are available in main memory.")]
        public AgentToolResult GetAllPDFDocuments()
        {
            try
            {
                var documentIds = _manager.GetAllDocumentIds();
                
                return AgentToolResult.Ok(
                    $"Found {documentIds.Count} PDF document(s) in memory",
                    new { DocumentIds = documentIds.ToArray(), Count = documentIds.Count });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to retrieve PDF documents: {ex.Message}");
            }
        }
        

        /// <summary>
        /// Exports a PDF document from memory to the file system.
        /// </summary>
        /// <param name="documentId">The ID of the PDF document to export.</param>
        /// <param name="filePath">The file path where the PDF should be saved.</param>
        /// <returns>Result containing the export file path.</returns>
        [Tool(Name = "ExportPDFDocument", Description = "Exports the PDF document to the file system.")]
        public AgentToolResult ExportPDFDocument(
            [ToolParameter(Description = "The ID of the document to export")] string documentId,
            [ToolParameter(Description = "The file path to export to")] string filePath)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentId);
                ArgumentNullException.ThrowIfNull(filePath);

                // Use output directory if provided and filePath is relative
                string fullPath = filePath;
                if (!Path.IsPathRooted(filePath) && !string.IsNullOrEmpty(_outputDirectory))
                {
                    fullPath = Path.Combine(_outputDirectory, filePath);
                }

                _manager.ExportDocument(fullPath, documentId);
                
                return AgentToolResult.Ok(
                    $"PDF document {documentId} exported successfully to {fullPath}",
                    new { FilePath = fullPath });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to export PDF document: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes a specific PDF document from memory by its ID.
        /// </summary>
        /// <param name="documentId">The ID of the document to remove.</param>
        /// <returns>Result indicating whether the document was removed successfully.</returns>
        [Tool(Name = "RemovePdfDocument", Description = "Removes a specific PDF document from memory by ID.")]
        public AgentToolResult RemovePdfDocument(
            [ToolParameter(Description = "The ID of the document to remove")] string documentId)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentId);
                bool removed = _manager.RemoveDocument(documentId);
                
                if (removed)
                {
                    return AgentToolResult.Ok($"PDF document {documentId} removed successfully from memory");
                }
                else
                {
                    return AgentToolResult.Fail($"PDF document not found: {documentId}");
                }
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to remove PDF document: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets the active PDF document context by ID.
        /// </summary>
        /// <param name="documentId">The ID of the document to set as active.</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "SetActivePdfDocument", Description = "Changes the active PDF document context by ID.")]
        public AgentToolResult SetActivePdfDocument(
            [ToolParameter(Description = "The ID of the document to set as active")] string documentId)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentId);
                _manager.SetActiveDocument(documentId);
                
                return AgentToolResult.Ok($"PDF document {documentId} set as active document");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to set active PDF document: {ex.Message}");
            }
        }
    }
}
