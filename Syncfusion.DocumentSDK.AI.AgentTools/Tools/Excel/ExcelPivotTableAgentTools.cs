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
using System.Linq;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.PivotTables;

namespace Syncfusion.AI.AgentTools.Excel
{
    /// <summary>
    /// Provides AI agent tools for Excel pivot table operations.
    /// Handles creating, editing, removing, styling, sorting, filtering,
    /// refreshing, and laying out pivot tables using Syncfusion XlsIO.
    /// Pivot table creation and manipulation is supported only in XLSX format (Excel 2007+).
    /// </summary>
    public class ExcelPivotTableAgentTools : AgentToolBase<IWorkbook>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelPivotTableAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The Excel workbook manager.</param>
        public ExcelPivotTableAgentTools(ExcelWorkbookManager manager)
            : base(manager, DocumentType.Excel) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelPivotTableAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public ExcelPivotTableAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Excel) { }

        // ─────────────────────────────────────────────────────────────────────
        // CREATE / EDIT / REMOVE
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>
        /// Creates a pivot table in a worksheet using a data range from a source worksheet.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="dataWorksheetName">The name of the worksheet that contains the source data.</param>
        /// <param name="dataRange">The cell range address of the source data (e.g., "A1:H50").</param>
        /// <param name="pivotWorksheetName">The name of the worksheet where the pivot table will be placed.</param>
        /// <param name="pivotTableName">The name for the new pivot table (e.g., "PivotTable1").</param>
        /// <param name="pivotLocation">The top-left cell address in the pivot worksheet where the pivot table will start (e.g., "A1").</param>
        /// <param name="rowFieldIndices">Comma-separated zero-based field indices to add as row fields (e.g., "2,6").</param>
        /// <param name="columnFieldIndices">Comma-separated zero-based field indices to add as column fields (e.g., "3").</param>
        /// <param name="dataFieldIndex">Zero-based field index to use as the data (values) field.</param>
        /// <param name="dataFieldCaption">Caption label for the data field (e.g., "Sum of Units").</param>
        /// <param name="subtotalType">Aggregation function for the data field: Sum, Count, Average, Max, Min, Product, CountNums, StdDev, StdDevP, Var, VarP (default: Sum).</param>
        /// <param name="builtInStyle">The built-in style for the pivot table to be applied. Default is None.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure with the pivot table name.</returns>
        [Tool(Name = "CreatePivotTable", Description = "Creates a pivot table in the specified worksheet using a data range from a source worksheet. Supports row, column, and data (values) fields with a chosen aggregation function. Only supported in XLSX format. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult CreatePivotTable(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet containing the source data")] string dataWorksheetName,
            [ToolParameter(Description = "The cell range of the source data")] string dataRange,
            [ToolParameter(Description = "The name of the worksheet where the pivot table will be placed")] string pivotWorksheetName,
            [ToolParameter(Description = "The name for the new pivot table (e.g., PivotTable1)")] string pivotTableName,
            [ToolParameter(Description = "The top-left cell address where the pivot table starts (e.g., A1)")] string pivotLocation,
            [ToolParameter(Description = "Comma-separated zero-based field indices for row fields (e.g., 2,6)")] string rowFieldIndices,
            [ToolParameter(Description = "Comma-separated zero-based field indices for column fields (e.g., 3)")] string columnFieldIndices,
            [ToolParameter(Description = "Zero-based field index to use as the data/values field")] int dataFieldIndex,
            [ToolParameter(Description = "Caption label for the data field (e.g., Sum of Units)")] string dataFieldCaption,
            [ToolParameter(Description = "Built-in style name: PivotStyleLight1-28, PivotStyleMedium1-28, PivotStyleDark1-28, or None")] string builtInStyle = "None",
            [ToolParameter(Description = "Aggregation type: Sum, Count, Average, Max, Min, Product, CountNums, StdDev, StdDevP, Var, VarP (default: Sum)")] string subtotalType = "Sum",
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(dataWorksheetName);
                ArgumentNullException.ThrowIfNull(dataRange);
                ArgumentNullException.ThrowIfNull(pivotWorksheetName);
                ArgumentNullException.ThrowIfNull(pivotTableName);
                ArgumentNullException.ThrowIfNull(pivotLocation);
                ArgumentNullException.ThrowIfNull(rowFieldIndices);
                ArgumentNullException.ThrowIfNull(columnFieldIndices);
                ArgumentNullException.ThrowIfNull(dataFieldCaption);

                var workbook = OpenDocument(workbookIdOrFilePath);
    
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");
                workbook.Application.DefaultVersion = ExcelVersion.Xlsx;
                workbook.Version = ExcelVersion.Xlsx;
                var dataWorksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == dataWorksheetName);
                if (dataWorksheet == null)
                    return AgentToolResult.Fail($"Data worksheet not found: {dataWorksheetName}");
                dataWorksheet.Activate();
                var pivotWorksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == pivotWorksheetName);
                if (pivotWorksheet == null)
                    return AgentToolResult.Fail($"Pivot worksheet not found: {pivotWorksheetName}");

                // Create pivot cache from data range
                IPivotCache cache = workbook.PivotCaches.Add(dataWorksheet[dataRange]);
                // Ensure the pivot cache is set to refresh when the workbook is opened in Excel
                if (cache is PivotCacheImpl cacheImpl)
                    cacheImpl.IsRefreshOnLoad = true;

                // Create pivot table at the specified location
                IPivotTable pivotTable = pivotWorksheet.PivotTables.Add(pivotTableName, pivotWorksheet[pivotLocation], cache);
                
                // Add row fields
                var rowIndices = ParseIndices(rowFieldIndices);
                foreach (int idx in rowIndices)
                {
                    if (idx >= 0 && idx < pivotTable.Fields.Count)
                        pivotTable.Fields[idx].Axis = PivotAxisTypes.Row;
                }

                // Add column fields
                var colIndices = ParseIndices(columnFieldIndices);
                foreach (int idx in colIndices)
                {
                    if (idx >= 0 && idx < pivotTable.Fields.Count)
                        pivotTable.Fields[idx].Axis = PivotAxisTypes.Column;
                }

                // Add data field
                if (dataFieldIndex < 0 || dataFieldIndex >= pivotTable.Fields.Count)
                    return AgentToolResult.Fail($"Data field index {dataFieldIndex} is out of range. The pivot table has {pivotTable.Fields.Count} fields.");

                IPivotField dataField = pivotTable.Fields[dataFieldIndex];
                PivotSubtotalTypes subtotal = ParseSubtotalType(subtotalType);
                pivotTable.DataFields.Add(dataField, dataFieldCaption, subtotal);

                if (Enum.TryParse<PivotBuiltInStyles>(builtInStyle, ignoreCase: true, out var style))
                    pivotTable.BuiltInStyle = style;

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_pivot_table.xlsx";
                    
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Pivot table '{pivotTableName}' created successfully in worksheet '{pivotWorksheetName}' into document {outputKey}",
                    new
                    {
                        PivotTableName = pivotTableName,
                        PivotWorksheet = pivotWorksheetName,
                        DataWorksheet = dataWorksheetName,
                        DataRange = dataRange,
                        RowFields = rowIndices,
                        ColumnFields = colIndices,
                        DataFieldIndex = dataFieldIndex,
                        DataFieldCaption = dataFieldCaption,
                        SubtotalType = subtotalType,
                        OutputKey = outputKey
                    });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to create pivot table: {ex.Message}");
            }
        }


        // ─────────────────────────────────────────────────────────────────────
        // PRIVATE HELPERS
        // ─────────────────────────────────────────────────────────────────────

        /// <summary>Parses a comma-separated string of integers into a list.</summary>
        private static List<int> ParseIndices(string input)
        {
            var result = new List<int>();
            if (string.IsNullOrWhiteSpace(input)) return result;

            foreach (var part in input.Split(',', StringSplitOptions.RemoveEmptyEntries))
            {
                if (int.TryParse(part.Trim(), out int idx))
                    result.Add(idx);
            }
            return result;
        }

        /// <summary>Maps a subtotal type string to the <see cref="PivotSubtotalTypes"/> enum.</summary>
        private static PivotSubtotalTypes ParseSubtotalType(string subtotalType)
        {
            return subtotalType.ToUpperInvariant() switch
            {
                "COUNT"     => PivotSubtotalTypes.Count,
                "AVERAGE"   => PivotSubtotalTypes.Average,
                "MAX"       => PivotSubtotalTypes.Max,
                "MIN"       => PivotSubtotalTypes.Min,
                "PRODUCT"   => PivotSubtotalTypes.Product,
                "COUNTA"    => PivotSubtotalTypes.Counta,
                "STDEV"    => PivotSubtotalTypes.Stdev,
                "STDEVP"   => PivotSubtotalTypes.Stdevp,
                "VAR"       => PivotSubtotalTypes.Var,
                "VARP"      => PivotSubtotalTypes.Varp,
                _           => PivotSubtotalTypes.Sum
            };
        }
    }
}
