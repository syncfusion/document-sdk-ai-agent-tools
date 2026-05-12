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
using Syncfusion.XlsIO;

namespace Syncfusion.AI.AgentTools.Excel
{
    /// <summary>
    /// Provides AI agent tools for Excel workbook lifecycle management and I/O operations.
    /// Handles workbook creation, import, export, and manager.
    /// </summary>
    public class ExcelWorkbookAgentTools : AgentToolBase
    {
        private readonly ExcelWorkbookManager _manager;
        private readonly string? _outputDirectory;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWorkbookAgentTools"/> class.
        /// </summary>
        /// <param name="manager">The Excel workbook manager.</param>
        /// <param name="outputDirectory">Optional output directory for file operations.</param>
        public ExcelWorkbookAgentTools(ExcelWorkbookManager manager, string? outputDirectory = null)
        {
            _manager = manager ?? throw new ArgumentNullException(nameof(manager));
            _outputDirectory = outputDirectory;
        }

        /// <summary>
        /// Sets the active workbook context by ID.
        /// </summary>
        /// <param name="workbookId">The ID of the workbook to set as active.</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "SetActiveWorkbook", Description = "Changes the active workbook context by ID.")]
        public AgentToolResult SetActiveWorkbook(
            [ToolParameter(Description = "The ID of the workbook to set as active")] string workbookId)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookId);
                _manager.SetActiveDocument(workbookId);

                return AgentToolResult.Ok($"Excel workbook {workbookId} set as active workbook");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to set active Excel workbook: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes a specific workbook from memory by its ID.
        /// </summary>
        /// <param name="workbookId">The ID of the workbook to remove.</param>
        /// <returns>Result indicating whether the workbook was removed successfully.</returns>
        [Tool(Name = "RemoveWorkbook", Description = "Removes a specific workbook from memory by ID.")]
        public AgentToolResult RemoveWorkbook(
            [ToolParameter(Description = "The ID of the workbook to remove")] string workbookId)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookId);
                bool removed = _manager.RemoveDocument(workbookId);

                if (removed)
                {
                    return AgentToolResult.Ok($"Excel workbook {workbookId} removed successfully from memory");
                }
                else
                {
                    return AgentToolResult.Fail($"Excel workbook not found: {workbookId}");
                }
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to remove Excel workbook: {ex.Message}");
            }
        }

        /// <summary>
        /// Retrieves all workbook IDs currently in memory.
        /// </summary>
        /// <returns>Result containing array of all workbook IDs.</returns>
        [Tool(Name = "GetAllWorkbooks", Description = "Returns all workbook IDs in memory.")]
        public AgentToolResult GetAllWorkbooks()
        {
            try
            {
                var workbookIds = _manager.GetAllDocumentIds();

                return AgentToolResult.Ok(
                    $"Found {workbookIds.Count} Excel workbook(s) in memory",
                    new { WorkbookIds = workbookIds.ToArray(), Count = workbookIds.Count });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to retrieve Excel workbooks: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates an Excel workbook in memory or loads an existing workbook from a file.
        /// </summary>
        /// <param name="filePath">Optional path to an existing Excel file. If null, creates a new blank workbook.</param>
        /// <param name="password">Optional password for encrypted workbooks.</param>
        /// <returns>Result containing the workbook ID of the created or loaded workbook.</returns>
        [Tool(Name = "CreateWorkbook", Description = "Creates an Excel workbook instance in memory. If filePath is null, creates a blank workbook; otherwise loads from the specified path.")]
        public AgentToolResult CreateWorkbook(
            [ToolParameter(Description = "Path to an existing Excel file, or null to create a new workbook")] string? filePath = null,
            [ToolParameter(Description = "Password for encrypted workbook")] string? password = null)
        {
            try
            {
                string workbookId;
                if (string.IsNullOrEmpty(filePath))
                {
                    // Create new blank workbook
                    var workbook = _manager.CreateDocument();
                    workbookId = _manager.ActiveWorkbookId ?? throw new InvalidOperationException("Failed to create workbook");
                }
                else
                {
                    // Load existing workbook
                    if (!string.IsNullOrEmpty(password))
                    {
                        _manager.ImportDocument(filePath, password);
                    }
                    else
                    {
                        _manager.ImportDocument(filePath);
                    }
                    workbookId = _manager.ActiveWorkbookId ?? throw new InvalidOperationException("Failed to import workbook");
                }
                
                return AgentToolResult.Ok(
                    $"Excel workbook created/loaded successfully with ID: {workbookId}",
                    new { WorkbookId = workbookId });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to create/load Excel workbook: {ex.Message}");
            }
        }

        /// <summary>
        /// Exports an Excel workbook from memory to the file system.
        /// </summary>
        /// <param name="workbookId">The ID of the workbook to export.</param>
        /// <param name="filePath">The file path where the workbook should be saved.</param>
        /// <param name="version">The Excel version format (Excel97to2003, Xlsx, Excel2016).</param>
        /// <returns>Result containing the export file path.</returns>
        [Tool(Name = "ExportWorkbook", Description = "Exports the workbook to the file system in the specified format.")]
        public AgentToolResult ExportWorkbook(
            [ToolParameter(Description = "The ID of the workbook to export")] string workbookId,
            [ToolParameter(Description = "The file path to export to")] string filePath,
            [ToolParameter(Description = "Excel version: XLS, XLSX, XLSM (default: XLSX)")] string version = "XLSX")
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookId);
                ArgumentNullException.ThrowIfNull(filePath);

                // Use output directory if provided and filePath is relative
                string fullPath = filePath;
                if (!Path.IsPathRooted(filePath) && !string.IsNullOrEmpty(_outputDirectory))
                {
                    fullPath = Path.Combine(_outputDirectory, filePath);
                }

                // Ensure correct file extension based on version
                string extension = version.ToUpperInvariant() switch
                {
                    "XLS" => ".xls",
                    "XLSX" => ".xlsx",
                    "XLSM" => ".xlsm",
                    "CSV" => ".csv",
                    _ => ".xlsx"
                };

                if (!fullPath.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
                {
                    fullPath = Path.ChangeExtension(fullPath, extension);
                }

                _manager.ExportDocument(fullPath, workbookId);
                
                return AgentToolResult.Ok(
                    $"Excel workbook {workbookId} exported successfully to {fullPath}",
                    new { FilePath = fullPath, Version = version });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to export Excel workbook: {ex.Message}");
            }
        }
    }
}
