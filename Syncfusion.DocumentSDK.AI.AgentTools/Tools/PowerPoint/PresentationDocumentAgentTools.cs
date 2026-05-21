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
using System.Collections.Generic;
using System.IO;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace Syncfusion.AI.AgentTools.PowerPoint
{
    /// <summary>
    /// Provides AI agent tools for PowerPoint presentation lifecycle management and I/O operations.
    /// Handles presentation creation, import, export, and manager.
    /// </summary>
    public class PresentationDocumentAgentTools : AgentToolBase
    {
        private readonly PresentationManager _manager;
        private readonly string? _outputDirectory;

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationDocumentAgentTools"/> class.
        /// </summary>
        /// <param name="manager">The presentation manager.</param>
        /// <param name="outputDirectory">Optional output directory for file operations.</param>
        public PresentationDocumentAgentTools(PresentationManager manager, string? outputDirectory = null)
        {
            _manager = manager ?? throw new ArgumentNullException(nameof(manager));
            _outputDirectory = outputDirectory;
        }

        /// <summary>
        /// Creates a empty PowerPoint presentation instance in memory or loads an existing presentation from a file.
        /// </summary>
        /// <param name="filePath">Optional path to an existing PowerPoint file. If null, creates a new empty presentation.</param>
        /// <param name="password">Optional password for encrypted presentations.</param>
        /// <returns>Result containing the document ID of the created or loaded presentation.</returns>
        [Tool(Name = "CreatePresentation", Description = "Create a empty PowerPoint Presentation instance in memory or load an existing one. FilePath - Use the path to load an existing Presentation; if null, creates an empty presentation. Returns the documentid.")]
        public AgentToolResult CreatePresentation(
            [ToolParameter(Description = "Path to an existing PowerPoint file, or null to create a new presentation")] string? filePath = null,
            [ToolParameter(Description = "Password for encrypted presentation")] string? password = null)
        {
            try
            {
                string documentId;
                if (string.IsNullOrEmpty(filePath))
                {
                    // Create new empty presentation
                    var presentation = _manager.CreateDocument();
                    documentId = _manager.ActivePresentationId ?? throw new InvalidOperationException("Failed to create presentation");
                }
                else
                {
                    // Load existing presentation
                    if (!string.IsNullOrEmpty(password))
                    {
                        _manager.ImportDocument(filePath, password);
                    }
                    else
                    {
                        _manager.ImportDocument(filePath);
                    }
                    documentId = _manager.ActivePresentationId ?? throw new InvalidOperationException("Failed to import presentation");
                }
                
                return AgentToolResult.Ok(
                    $"PowerPoint presentation created/loaded successfully with ID: {documentId}",
                    new { DocumentId = documentId });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to create/load PowerPoint presentation: {ex.Message}");
            }
        }

        /// <summary>
        /// Retrieves all presentation document IDs currently in memory.
        /// </summary>
        /// <returns>Result containing array of all presentation document IDs.</returns>
        [Tool(Name = "GetAllPresentations", Description = "Returns all documentid of the PowerPoint presentation instances that are available in main memory.")]
        public AgentToolResult GetAllPresentations()
        {
            try
            {
                var documentIds = _manager.GetAllDocumentIds();
                string message = documentIds.Count == 0
                    ? "No PowerPoint presentations found in memory"
                    : $"Found {documentIds.Count} PowerPoint presentation(s) in memory: {string.Join(", ", documentIds)}";

                return AgentToolResult.Ok(message, new { DocumentIds = documentIds.ToArray(), Count = documentIds.Count });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to retrieve PowerPoint presentations: {ex.Message}");
            }
        }

        /// <summary>
        /// Exports a PowerPoint presentation from memory to the file system.
        /// </summary>
        /// <param name="documentId">The ID of the presentation to export.</param>
        /// <param name="filePath">The file path where the presentation should be saved.</param>
        /// <param name="format">The export format (PPTX, PDF, Image).</param>
        /// <returns>Result containing the export file path.</returns>
        [Tool(Name = "ExportPresentation", Description = "Exports the PowerPoint presentation to the file system.")]
        public AgentToolResult ExportPresentation(
            [ToolParameter(Description = "The ID of the presentation to export")] string documentId,
            [ToolParameter(Description = "The file name or full path to export to. If only filename is provided, uses the output directory.")]
            string filePath,
            [ToolParameter(Description = "Export format: PPTX only")] string format = "PPTX")
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentId);
                ArgumentNullException.ThrowIfNull(filePath);

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
                
                return AgentToolResult.Ok(
                    $"PowerPoint presentation {documentId} exported successfully to {fullPath}",
                    new { FilePath = fullPath, Format = format });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to export PowerPoint presentation: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes a specific presentation from memory by its ID.
        /// </summary>
        /// <param name="documentId">The ID of the presentation to remove.</param>
        /// <returns>Result indicating whether the presentation was removed successfully.</returns>
        [Tool(Name = "RemovePresentation", Description = "Removes a specific PowerPoint presentation from memory by ID.")]
        public AgentToolResult RemovePresentation(
            [ToolParameter(Description = "The ID of the presentation to remove")] string documentId)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentId);
                bool removed = _manager.RemoveDocument(documentId);
                
                if (removed)
                {
                    return AgentToolResult.Ok($"PowerPoint presentation {documentId} removed successfully from memory");
                }
                else
                {
                    return AgentToolResult.Fail($"PowerPoint presentation not found: {documentId}");
                }
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to remove PowerPoint presentation: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets the active presentation context by ID.
        /// </summary>
        /// <param name="documentId">The ID of the presentation to set as active.</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "SetActivePresentation", Description = "Changes the active PowerPoint presentation context by ID.")]
        public AgentToolResult SetActivePresentation(
            [ToolParameter(Description = "The ID of the presentation to set as active")] string documentId)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentId);
                _manager.SetActiveDocument(documentId);
                
                return AgentToolResult.Ok($"PowerPoint presentation {documentId} set as active presentation");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to set active PowerPoint presentation: {ex.Message}");
            }
        }
    }
}
