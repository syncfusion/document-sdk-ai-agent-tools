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
using System.Linq;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.XlsIO;

namespace Syncfusion.AI.AgentTools.Excel
{
    /// <summary>
    /// Provides AI agent tools for Excel worksheet management operations.
    /// Handles worksheet creation, renaming, deletion, and listing.
    /// </summary>
    public class ExcelWorksheetAgentTools : AgentToolBase<IWorkbook>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWorksheetAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The Excel workbook manager.</param>
        public ExcelWorksheetAgentTools(ExcelWorkbookManager manager)
            : base(manager, DocumentType.Excel) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWorksheetAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public ExcelWorksheetAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Excel) { }

        /// <summary>
        /// Creates a new worksheet in the specified workbook.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="sheetName">Optional name for the new worksheet. If null, a default name is used.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the name of the created worksheet.</returns>
        [Tool(Name = "CreateWorksheet", Description = "Creates a worksheet inside the specified workbook. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode). If sheetName is null, a default name is used.")]
        public AgentToolResult CreateWorksheet(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "Optional name for the worksheet")] string? sheetName = null,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Create new worksheet
                IWorksheet worksheet;
                if (string.IsNullOrEmpty(sheetName))
                {
                    // Add with default name
                    worksheet = workbook.Worksheets.Create();
                }
                else
                {
                    // Add with specified name
                    worksheet = workbook.Worksheets.Create(sheetName);
                }
                worksheet.Activate();

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_worksheet_created.xlsx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Worksheet '{worksheet.Name}' created successfully in workbook {outputKey}",
                    new { WorksheetName = worksheet.Name, WorkbookId = outputKey });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to create worksheet: {ex.Message}");
            }
        }

        /// <summary>
        /// Deletes a worksheet from the workbook.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet to delete.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "DeleteWorksheet", Description = "Deletes a worksheet from the workbook. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult DeleteWorksheet(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet to delete")] string worksheetName,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Find the worksheet by name
                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                // Cannot delete if it's the only worksheet
                if (workbook.Worksheets.Count <= 1)
                    return AgentToolResult.Fail("Cannot delete the only worksheet in the workbook");

                // Delete the worksheet
                worksheet.Remove();

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_worksheet_deleted.xlsx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID
                
                return AgentToolResult.Ok($"Worksheet '{worksheetName}' deleted successfully from workbook {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to delete worksheet: {ex.Message}");
            }
        }
    }
}
