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
using Syncfusion.Drawing;
using Syncfusion.XlsIO;

namespace Syncfusion.AI.AgentTools.Excel
{
    /// <summary>
    /// Provides AI agent tools for Excel conditional formatting operations.
    /// Handles adding, removing, and managing conditional formatting rules in worksheets.
    /// Supports CellValue, Formula, DataBar, ColorScale, and IconSet format types.
    /// </summary>
    public class ExcelConditionalFormattingAgentTools : AgentToolBase<IWorkbook>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelConditionalFormattingAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The Excel workbook manager.</param>
        public ExcelConditionalFormattingAgentTools(ExcelWorkbookManager manager)
            : base(manager, DocumentType.Excel) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelConditionalFormattingAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public ExcelConditionalFormattingAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Excel) { }

        /// <summary>
        /// Adds conditional formatting to a cell or range in the worksheet.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="rangeAddress">The cell or range address (e.g., "A1", "B5:C10").</param>
        /// <param name="formatType">The conditional format type: CellValue, Formula, DataBar, ColorScale, IconSet, etc.</param>
        /// <param name="operatorType">Comparison operator (only for CellValue type) EXACT enum values only: Equal, NotEqual, Greater, Less, Between, NotBetween, GreaterOrEqual, LessOrEqual. Do NOT use LessThan, GreaterThan, or variants.  Not used for DataBar, ColorScale, or IconSet.</param>
        /// <param name="firstFormula">The first formula or value for comparison (only for CellValue/Formula types). Not used for DataBar, ColorScale, or IconSet.</param>
        /// <param name="secondFormula">Optional second formula (required for Between/NotBetween operators).</param>
        /// <param name="backColor">Optional background color. MUST be a name from the `ExcelKnownColors` enum (e.g., Light_orange, Red, Yellow, Light_green). Do NOT pass hex values (like #FF0000), RGB values, or other formats — they will be rejected.</param>
        /// <param name="isBold">Optional: Apply bold formatting.</param>
        /// <param name="isItalic">Optional: Apply italic formatting.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "AddConditionalFormat", Description = "Adds conditional formatting to a cell or range based on specified criteria. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult AddConditionalFormat(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "The cell or range address (e.g., A1, B5:C10)")] string rangeAddress,
            [ToolParameter(Description = "Format type: CellValue, Formula, DataBar, ColorScale, IconSet")] string formatType,
            [ToolParameter(Description = "Operator (only for CellValue format type) EXACT values only: 'Equal', 'NotEqual', 'Greater', 'Less', 'Between', 'NotBetween', 'GreaterOrEqual', 'LessOrEqual'. Use empty string or 'None' for DataBar/ColorScale/IconSet")] string? operatorType = null,
            [ToolParameter(Description = "The first formula or value (only for CellValue/Formula). Not needed for DataBar/ColorScale/IconSet")] string? firstFormula = null,
            [ToolParameter(Description = "Optional second formula (for Between/NotBetween)")] string? secondFormula = null,
            [ToolParameter(Description = "Optional background color. Must be a valid ExcelKnownColors and only ExcelKnownColors enum names are accepted (e.g., Red, Yellow, Light_orange, Light_green, Blue, Green, White, Black). Do NOT use hex values (#FF0000), RGB values, or other color formats—they will fail.")] string? backColor = null,
            [ToolParameter(Description = "Optional: Apply bold formatting")] bool? isBold = null,
            [ToolParameter(Description = "Optional: Apply italic formatting")] bool? isItalic = null,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(rangeAddress);
                ArgumentNullException.ThrowIfNull(formatType);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                // Get the range and add conditional format
                var range = worksheet.Range[rangeAddress];
                var conditionalFormats = range.ConditionalFormats;
                var condition = conditionalFormats.AddCondition();

                // Parse and set format type
                if (!Enum.TryParse<ExcelCFType>(formatType, true, out var cfType))
                {
                    string formatLower = formatType.ToLower();
                    if (formatLower.Contains("cell") || formatLower.Contains("value"))
                        cfType = ExcelCFType.CellValue;
                    else if (formatLower.Contains("formula"))
                        cfType = ExcelCFType.Formula;
                    else if (formatLower.Contains("data") || formatLower.Contains("bar"))
                        cfType = ExcelCFType.DataBar;
                    else if (formatLower.Contains("color") || formatLower.Contains("scale"))
                        cfType = ExcelCFType.ColorScale;
                    else if (formatLower.Contains("icon") || formatLower.Contains("set"))
                        cfType = ExcelCFType.IconSet;
                    else
                        return AgentToolResult.Fail($"Invalid format type: {formatType}");
                }

                condition.FormatType = cfType;

                // Handle parameters based on format type
                if (cfType == ExcelCFType.CellValue)
                {
                    // CellValue type requires operator and formula
                    if (string.IsNullOrEmpty(operatorType))
                        return AgentToolResult.Fail("Operator is required for CellValue format type");

                    if (string.IsNullOrEmpty(firstFormula))
                        return AgentToolResult.Fail("First formula is required for CellValue format type");

                    if (!Enum.TryParse<ExcelComparisonOperator>(operatorType, true, out var compOperator))
                    {
                        if (operatorType.ToLower().Contains("less"))
                        {
                            if (operatorType.ToLower().Contains("equal"))
                                compOperator = ExcelComparisonOperator.LessOrEqual;
                            else
                                compOperator = ExcelComparisonOperator.Less;
                        }
                        else if (operatorType.ToLower().Contains("greater"))
                        {
                            if (operatorType.ToLower().Contains("equal"))
                                compOperator = ExcelComparisonOperator.GreaterOrEqual;
                            else
                                compOperator = ExcelComparisonOperator.Greater;
                        }
                        else
                            compOperator = ExcelComparisonOperator.Equal;
                    }

                    condition.Operator = compOperator;

                    // Check if second formula is required
                    if ((compOperator == ExcelComparisonOperator.Between || compOperator == ExcelComparisonOperator.NotBetween) 
                        && string.IsNullOrEmpty(secondFormula))
                    {
                        return AgentToolResult.Fail($"Second formula is required for {operatorType} operator");
                    }

                    // Set formulas for CellValue type
                    condition.FirstFormula = firstFormula;
                    if (!string.IsNullOrEmpty(secondFormula))
                    {
                        condition.SecondFormula = secondFormula;
                    }
                }
                else if (cfType == ExcelCFType.Formula)
                {
                    // Formula type requires formula but no operator
                    if (string.IsNullOrEmpty(firstFormula))
                        return AgentToolResult.Fail("First formula is required for Formula format type");

                    condition.FirstFormula = firstFormula;
                }
                else if (cfType == ExcelCFType.DataBar || cfType == ExcelCFType.ColorScale || cfType == ExcelCFType.IconSet)
                {
                    // DataBar, ColorScale, and IconSet don't need operators or formulas
                    // They work automatically based on cell values in the range
                    // No additional setup needed - format is applied automatically
                }
                else
                {
                    // For other types, set formulas if provided
                    if (!string.IsNullOrEmpty(firstFormula))
                    {
                        condition.FirstFormula = firstFormula;
                    }
                    if (!string.IsNullOrEmpty(secondFormula))
                    {
                        condition.SecondFormula = secondFormula;
                    }
                }

                // Apply formatting options
                if (!string.IsNullOrEmpty(backColor))
                {
                    if (Enum.TryParse<ExcelKnownColors>(backColor, true, out var color))
                    {
                        condition.BackColor = color;
                    }
                    else
                    {
                        try
                        {
                            condition.BackColorRGB = ColorTranslator.FromHtml(backColor);
                        }
                        catch
                        {
                            condition.BackColor = ExcelKnownColors.White;
                        }
                    }
                }

                if (isBold.HasValue)
                    condition.IsBold = isBold.Value;

                if (isItalic.HasValue)
                    condition.IsItalic = isItalic.Value;
          // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_conditional_format.xlsx";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Conditional formatting added successfully to range {rangeAddress} in worksheet '{worksheetName}' into document {outputKey}",
                    new 
                    { 
                        RangeAddress = rangeAddress,
                        FormatType = formatType,
                        Operator = operatorType,
                        FirstFormula = firstFormula,
                        SecondFormula = secondFormula,
                        BackColor = backColor,
                        IsBold = isBold,
                        IsItalic = isItalic,
                        OutputKey = outputKey
                    });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to add conditional formatting: {ex.Message}");
            }
        }
    }
}
