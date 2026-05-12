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
    /// Provides AI agent tools for Excel data validation operations.
    /// Handles creating and managing data validation rules for cells and ranges.
    /// </summary>
    public class ExcelDataValidationAgentTools : AgentToolBase<IWorkbook>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelDataValidationAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The Excel workbook manager.</param>
        public ExcelDataValidationAgentTools(ExcelWorkbookManager manager)
            : base(manager, DocumentType.Excel) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelDataValidationAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public ExcelDataValidationAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Excel) { }

        /// <summary>
        /// Adds a dropdown list validation to a cell or range.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="rangeAddress">The cell or range address (e.g., "A1" or "A1:A10").</param>
        /// <param name="listValues">Comma-separated list of values (e.g., "Option1,Option2,Option3"). Limited to 255 characters.</param>
        /// <param name="showErrorBox">Whether to show error message box. Default is true.</param>
        /// <param name="errorTitle">Optional title for the error message box.</param>
        /// <param name="errorMessage">Optional error message to display.</param>
        /// <param name="showPromptBox">Whether to show input prompt box. Default is false.</param>
        /// <param name="promptMessage">Optional prompt message to display.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "AddDropdownValidation", Description = "Adds a dropdown list data validation to a cell or range. List values are limited to 255 characters including separators. Supports both InMemory and DocumentStorage modes.")]
        public AgentToolResult AddDropdownValidation(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "The cell or range address (e.g., A1 or A1:A10)")] string rangeAddress,
            [ToolParameter(Description = "The source range formula for range Validation(e.g., =Sheet1!$A$1:$A$10). Use this ONLY when the dropdown values come from a cell range. Leave null if using a custom list.")] string sourceRange = null,
            [ToolParameter(Description = "Comma-separated list of values (e.g., Option1,Option2,Option3). Use this ONLY when providing a custom list. Leave null if using a source range.")] string listValues = "",
            [ToolParameter(Description = "Whether to show error message box")] bool showErrorBox = true,
            [ToolParameter(Description = "Optional title for error message box")] string? errorTitle = null,
            [ToolParameter(Description = "Optional error message to display")] string? errorMessage = null,
            [ToolParameter(Description = "Whether to show input prompt box")] bool showPromptBox = false,
            [ToolParameter(Description = "Optional prompt message to display")] string? promptMessage = null,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only)")] string outputFilePath = "output_AddDropdownValidation.xlsx")
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(rangeAddress);
                ArgumentNullException.ThrowIfNull(listValues);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                var range = worksheet.Range[rangeAddress];
                var validation = range.DataValidation;

                // Configure error box
                validation.ShowErrorBox = showErrorBox;
                if (!string.IsNullOrEmpty(errorTitle))
                    validation.ErrorBoxTitle = errorTitle;
                if (!string.IsNullOrEmpty(errorMessage))
                    validation.ErrorBoxText = errorMessage;

                // Configure prompt box
                validation.ShowPromptBox = showPromptBox;
                if (!string.IsNullOrEmpty(promptMessage))
                    validation.PromptBoxText = promptMessage;
                
                if (!string.IsNullOrEmpty(listValues))
                {
                    // Split and validate list values
                    var valuesList = listValues.Split(',', StringSplitOptions.RemoveEmptyEntries)
                        .Select(v => v.Trim())
                        .ToArray();

                    if (valuesList.Length == 0)
                        return AgentToolResult.Fail("List values cannot be empty");

                    validation.ListOfValues = valuesList;

                    // ── Save ────────────────────────────────────────────────────────
                    string outputKey = outputFilePath;
                    SaveDocument(outputKey, workbook);
                    if (Mode == DocumentManagerMode.InMemory)
                        outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                    return AgentToolResult.Ok(
                    $"Dropdown list validation added successfully to range {rangeAddress}. Output: {outputKey}",
                    new
                    {
                        Range = rangeAddress,
                        WorksheetName = worksheetName,
                        ListItems = valuesList,
                        ItemCount = valuesList.Length,
                        OutputKey = outputKey
                    });
                }
                else if (!string.IsNullOrEmpty(sourceRange))
                {
                    validation.AllowType = ExcelDataType.User;
                    validation.FirstFormula = worksheet[sourceRange].AddressGlobal;

                    // ── Save ────────────────────────────────────────────────────────
                    string outputKey = outputFilePath;
                    SaveDocument(outputKey, workbook);
                    if (Mode == DocumentManagerMode.InMemory)
                        outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                    return AgentToolResult.Ok(
                    $"Dropdown validation with source range added successfully to {rangeAddress}. Output: {outputKey}",
                    new
                    {
                        Range = rangeAddress,
                        WorksheetName = worksheetName,
                        SourceRange = sourceRange,
                        OutputKey = outputKey
                    });
                }
                else
                {
                    return AgentToolResult.Fail("Either listValues or sourceRange must be provided for dropdown validation.");
                }

                
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to add dropdown list validation: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds number validation to a cell or range.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="rangeAddress">The cell or range address.</param>
        /// <param name="numberType">The number type: Integer, Decimal.</param>
        /// <param name="comparisonOperator">The comparison operator: Between, NotBetween, Equal, NotEqual, Greater, GreaterOrEqual, Less, LessOrEqual.</param>
        /// <param name="firstValue">The first value or minimum value.</param>
        /// <param name="secondValue">Optional second value or maximum value (required for Between and NotBetween).</param>
        /// <param name="showErrorBox">Whether to show error message box. Default is true.</param>
        /// <param name="errorTitle">Optional title for the error message box.</param>
        /// <param name="errorMessage">Optional error message to display.</param>
        /// <param name="showPromptBox">Whether to show input prompt box. Default is false.</param>
        /// <param name="promptMessage">Optional prompt message to display.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "AddNumberValidation", Description = "Adds number validation to a cell or range with specified comparison operator and values. Supports both InMemory and DocumentStorage modes.")]
        public AgentToolResult AddNumberValidation(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "The cell or range address")] string rangeAddress,
            [ToolParameter(Description = "Number type ONLY these exact values: Integer or Decimal")] string numberType = "decimal",
            [ToolParameter(Description = "Comparison operator must ONLY be these exact values: Between, NotBetween, Equal, NotEqual, Greater, GreaterOrEqual, Less, LessOrEqual. Do NOT use LessThan, GreaterThan, <, >, <=, >=, or other variants.")] string comparisonOperator = "",
            [ToolParameter(Description = "The first value or minimum value")] string firstValue = "",
            [ToolParameter(Description = "Optional second value or maximum value (required for Between/NotBetween)")] string? secondValue = null,
            [ToolParameter(Description = "Whether to show error message box")] bool showErrorBox = true,
            [ToolParameter(Description = "Optional title for error message box")] string? errorTitle = null,
            [ToolParameter(Description = "Optional error message to display")] string? errorMessage = null,
            [ToolParameter(Description = "Whether to show input prompt box")] bool showPromptBox = false,
            [ToolParameter(Description = "Optional prompt message to display")] string? promptMessage = null,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only)")] string outputFilePath = "output_AddNumberValidation.xlsx")
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(rangeAddress);
                ArgumentNullException.ThrowIfNull(numberType);
                ArgumentNullException.ThrowIfNull(comparisonOperator);
                ArgumentNullException.ThrowIfNull(firstValue);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                var range = worksheet.Range[rangeAddress];
                var validation = range.DataValidation;

                // Set number type
                validation.AllowType = numberType.ToLower() switch
                {
                    "integer" => ExcelDataType.Integer,
                    "decimal" => ExcelDataType.Decimal,
                    _ => ExcelDataType.Decimal
                };

                // Set comparison operator
                validation.CompareOperator = GetDataValidationOperator(comparisonOperator);

                // Set formulas
                validation.FirstFormula = firstValue;
                if (!string.IsNullOrEmpty(secondValue))
                    validation.SecondFormula = secondValue;

                // Configure error box
                validation.ShowErrorBox = showErrorBox;
                if (!string.IsNullOrEmpty(errorTitle))
                    validation.ErrorBoxTitle = errorTitle;
                if (!string.IsNullOrEmpty(errorMessage))
                    validation.ErrorBoxText = errorMessage;

                // Configure prompt box
                validation.ShowPromptBox = showPromptBox;
                if (!string.IsNullOrEmpty(promptMessage))
                    validation.PromptBoxText = promptMessage;

                // ── Save ────────────────────────────────────────────────────────
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Number validation ({numberType}, {comparisonOperator}) added successfully to range {rangeAddress}. Output: {outputKey}",
                    new
                    {
                        Range = rangeAddress,
                        WorksheetName = worksheetName,
                        NumberType = numberType,
                        ComparisonOperator = comparisonOperator,
                        FirstValue = firstValue,
                        SecondValue = secondValue,
                        OutputKey = outputKey
                    });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to add number validation: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds date validation to a cell or range.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="rangeAddress">The cell or range address.</param>
        /// <param name="comparisonOperator">Comparison operator must ONLY be these exact values: Between, NotBetween, Equal, NotEqual, Greater, GreaterOrEqual, Less, LessOrEqual. Do NOT use LessThan, GreaterThan, <, >, <=, >=, or other variants."ONLY exact values: Between, NotBetween, Equal, NotEqual, Greater, GreaterOrEqual, Less, LessOrEqual. Do NOT use LessThan, GreaterThan, <, >, <=, >=.</param>
        /// <param name="firstDate">The first date in yyyy-MM-dd format.</param>
        /// <param name="secondDate">Optional second date in yyyy-MM-dd format (required for Between and NotBetween).</param>
        /// <param name="showErrorBox">Whether to show error message box. Default is true.</param>
        /// <param name="errorTitle">Optional title for the error message box.</param>
        /// <param name="errorMessage">Optional error message to display.</param>
        /// <param name="showPromptBox">Whether to show input prompt box. Default is false.</param>
        /// <param name="promptMessage">Optional prompt message to display.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "AddDateValidation", Description = "Adds date validation to a cell or range with specified comparison operator and dates. Supports both InMemory and DocumentStorage modes.")]
        public AgentToolResult AddDateValidation(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "The cell or range address")] string rangeAddress,
            [ToolParameter(Description = "Comparison operator must ONLY be these exact values: Between, NotBetween, Equal, NotEqual, Greater, GreaterOrEqual, Less, LessOrEqual. Do NOT use LessThan, GreaterThan, <, >, <=, >=, or other variants.")] string comparisonOperator,
            [ToolParameter(Description = "First date in yyyy-MM-dd format")] string firstDate,
            [ToolParameter(Description = "Optional second date in yyyy-MM-dd format (required for Between/NotBetween)")] string? secondDate = null,
            [ToolParameter(Description = "Whether to show error message box")] bool showErrorBox = true,
            [ToolParameter(Description = "Optional title for error message box")] string? errorTitle = null,
            [ToolParameter(Description = "Optional error message to display")] string? errorMessage = null,
            [ToolParameter(Description = "Whether to show input prompt box")] bool showPromptBox = false,
            [ToolParameter(Description = "Optional prompt message to display")] string? promptMessage = null,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only)")] string outputFilePath = "output_AddDateValidation.xlsx")
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(rangeAddress);
                ArgumentNullException.ThrowIfNull(comparisonOperator);
                ArgumentNullException.ThrowIfNull(firstDate);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                var range = worksheet.Range[rangeAddress];
                var validation = range.DataValidation;

                validation.AllowType = ExcelDataType.Date;

                validation.CompareOperator = GetDataValidationOperator(comparisonOperator);

                // Parse dates
                if (!DateTime.TryParse(firstDate, out DateTime firstDateTime))
                    return AgentToolResult.Fail($"Invalid first date format: {firstDate}. Use yyyy-MM-dd format.");

                validation.FirstDateTime = firstDateTime;

                if (!string.IsNullOrEmpty(secondDate))
                {
                    if (!DateTime.TryParse(secondDate, out DateTime secondDateTime))
                        return AgentToolResult.Fail($"Invalid second date format: {secondDate}. Use yyyy-MM-dd format.");
                    validation.SecondDateTime = secondDateTime;
                }

                // Configure error box
                validation.ShowErrorBox = showErrorBox;
                if (!string.IsNullOrEmpty(errorTitle))
                    validation.ErrorBoxTitle = errorTitle;
                if (!string.IsNullOrEmpty(errorMessage))
                    validation.ErrorBoxText = errorMessage;

                // Configure prompt box
                validation.ShowPromptBox = showPromptBox;
                if (!string.IsNullOrEmpty(promptMessage))
                    validation.PromptBoxText = promptMessage;

                // ── Save ────────────────────────────────────────────────────────
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Date validation ({comparisonOperator}) added successfully to range {rangeAddress}. Output: {outputKey}",
                    new
                    {
                        Range = rangeAddress,
                        WorksheetName = worksheetName,
                        ComparisonOperator = comparisonOperator,
                        FirstDate = firstDate,
                        SecondDate = secondDate,
                        OutputKey = outputKey
                    });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to add date validation: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds time validation to a cell or range.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="rangeAddress">The cell or range address.</param>
        /// <param name="comparisonOperator">ONLY exact values: Between, NotBetween, Equal, NotEqual, Greater, GreaterOrEqual, Less, LessOrEqual. Do NOT use LessThan, GreaterThan, <, >, <=, >=.</param>
        /// <param name="firstTime">The first time value in 24-hour format (e.g., "10:00" or "18:30").</param>
        /// <param name="secondTime">Optional second time value (required for Between and NotBetween).</param>
        /// <param name="showErrorBox">Whether to show error message box. Default is true.</param>
        /// <param name="errorTitle">Optional title for the error message box.</param>
        /// <param name="errorMessage">Optional error message to display.</param>
        /// <param name="showPromptBox">Whether to show input prompt box. Default is false.</param>
        /// <param name="promptMessage">Optional prompt message to display.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "AddTimeValidation", Description = "Adds time validation to a cell or range with specified comparison operator and time values. Use 24-hour format like 10:00 or 18:30. Supports both InMemory and DocumentStorage modes.")]
        public AgentToolResult AddTimeValidation(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "The cell or range address")] string rangeAddress,
            [ToolParameter(Description = "Comparison operator must ONLY be these exact values: Between, NotBetween, Equal, NotEqual, Greater, GreaterOrEqual, Less, LessOrEqual. Do NOT use LessThan, GreaterThan, <, >, <=, >=, or other variants.")] string comparisonOperator,
            [ToolParameter(Description = "First time value in HH:mm format (e.g., 10:00 or 18:30)")] string firstTime,
            [ToolParameter(Description = "Optional second time value in HH:mm format (required for Between/NotBetween)")] string? secondTime = null,
            [ToolParameter(Description = "Whether to show error message box")] bool showErrorBox = true,
            [ToolParameter(Description = "Optional title for error message box")] string? errorTitle = null,
            [ToolParameter(Description = "Optional error message to display")] string? errorMessage = null,
            [ToolParameter(Description = "Whether to show input prompt box")] bool showPromptBox = false,
            [ToolParameter(Description = "Optional prompt message to display")] string? promptMessage = null,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only)")] string outputFilePath = "output_AddTimeValidation.xlsx")
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(rangeAddress);
                ArgumentNullException.ThrowIfNull(comparisonOperator);
                ArgumentNullException.ThrowIfNull(firstTime);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                var range = worksheet.Range[rangeAddress];
                var validation = range.DataValidation;

                validation.AllowType = ExcelDataType.Time;

                validation.CompareOperator = GetDataValidationOperator(comparisonOperator);

                // Convert time string to decimal fraction of day
                var firstTimeDecimal = ConvertTimeToDecimal(firstTime);
                if (firstTimeDecimal == null)
                    return AgentToolResult.Fail($"Invalid time format: {firstTime}. Use HH:mm format (e.g., 10:00 or 18:30)");

                validation.FirstFormula = firstTimeDecimal.Value.ToString("0.################");
                
                if (!string.IsNullOrEmpty(secondTime))
                {
                    var secondTimeDecimal = ConvertTimeToDecimal(secondTime);
                    if (secondTimeDecimal == null)
                        return AgentToolResult.Fail($"Invalid time format: {secondTime}. Use HH:mm format (e.g., 10:00 or 18:30)");
                    
                    validation.SecondFormula = secondTimeDecimal.Value.ToString("0.################");
                }

                // Configure error box
                validation.ShowErrorBox = showErrorBox;
                if (!string.IsNullOrEmpty(errorTitle))
                    validation.ErrorBoxTitle = errorTitle;
                if (!string.IsNullOrEmpty(errorMessage))
                    validation.ErrorBoxText = errorMessage;

                // Configure prompt box
                validation.ShowPromptBox = showPromptBox;
                if (!string.IsNullOrEmpty(promptMessage))
                    validation.PromptBoxText = promptMessage;

                // ── Save ────────────────────────────────────────────────────────
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Time validation ({comparisonOperator}) added successfully to range {rangeAddress}. Output: {outputKey}",
                    new
                    {
                        Range = rangeAddress,
                        WorksheetName = worksheetName,
                        ComparisonOperator = comparisonOperator,
                        FirstTime = firstTime,
                        SecondTime = secondTime,
                        OutputKey = outputKey
                    });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to add time validation: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds text length validation to a cell or range.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="rangeAddress">The cell or range address.</param>
        /// <param name="comparisonOperator">"Comparison operator must ONLY be these exact values: Between, NotBetween, Equal, NotEqual, Greater, GreaterOrEqual, Less, LessOrEqual. Do NOT use LessThan, GreaterThan, <, >, <=, >=, or other variants."</param>
        /// <param name="firstLength">The first length value or minimum length.</param>
        /// <param name="secondLength">Optional second length value or maximum length (required for Between and NotBetween).</param>
        /// <param name="showErrorBox">Whether to show error message box. Default is true.</param>
        /// <param name="errorTitle">Optional title for the error message box.</param>
        /// <param name="errorMessage">Optional error message to display.</param>
        /// <param name="showPromptBox">Whether to show input prompt box. Default is false.</param>
        /// <param name="promptMessage">Optional prompt message to display.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "AddTextLengthValidation", Description = "Adds text length validation to a cell or range with specified comparison operator and length values. Supports both InMemory and DocumentStorage modes.")]
        public AgentToolResult AddTextLengthValidation(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "The cell or range address")] string rangeAddress,
            [ToolParameter(Description = "Comparison operator must ONLY be these exact values: Between, NotBetween, Equal, NotEqual, Greater, GreaterOrEqual, Less, LessOrEqual. Do NOT use LessThan, GreaterThan, <, >, <=, >=, or other variants.")] string comparisonOperator,
            [ToolParameter(Description = "First length value or minimum length")] string firstLength,
            [ToolParameter(Description = "Optional second length value or maximum length (required for Between/NotBetween)")] string? secondLength = null,
            [ToolParameter(Description = "Whether to show error message box")] bool showErrorBox = true,
            [ToolParameter(Description = "Optional title for error message box")] string? errorTitle = null,
            [ToolParameter(Description = "Optional error message to display")] string? errorMessage = null,
            [ToolParameter(Description = "Whether to show input prompt box")] bool showPromptBox = false,
            [ToolParameter(Description = "Optional prompt message to display")] string? promptMessage = null,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only)")] string outputFilePath = "output_AddTextLengthValidation.xlsx")
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(rangeAddress);
                ArgumentNullException.ThrowIfNull(comparisonOperator);
                ArgumentNullException.ThrowIfNull(firstLength);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                var range = worksheet.Range[rangeAddress];
                var validation = range.DataValidation;

                validation.AllowType = ExcelDataType.TextLength;
                
                validation.CompareOperator = GetDataValidationOperator(comparisonOperator);

                validation.FirstFormula = firstLength;
                if (!string.IsNullOrEmpty(secondLength))
                    validation.SecondFormula = secondLength;

                // Configure error box
                validation.ShowErrorBox = showErrorBox;
                if (!string.IsNullOrEmpty(errorTitle))
                    validation.ErrorBoxTitle = errorTitle;
                if (!string.IsNullOrEmpty(errorMessage))
                    validation.ErrorBoxText = errorMessage;

                // Configure prompt box
                validation.ShowPromptBox = showPromptBox;
                if (!string.IsNullOrEmpty(promptMessage))
                    validation.PromptBoxText = promptMessage;

                // ── Save ────────────────────────────────────────────────────────
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Text length validation ({comparisonOperator}) added successfully to range {rangeAddress}. Output: {outputKey}",
                    new
                    {
                        Range = rangeAddress,
                        WorksheetName = worksheetName,
                        ComparisonOperator = comparisonOperator,
                        FirstLength = firstLength,
                        SecondLength = secondLength,
                        OutputKey = outputKey
                    });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to add text length validation: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds custom formula validation to a cell or range.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="rangeAddress">The cell or range address.</param>
        /// <param name="formula">The validation formula (e.g., "=A1>10").</param>
        /// <param name="showErrorBox">Whether to show error message box. Default is true.</param>
        /// <param name="errorTitle">Optional title for the error message box.</param>
        /// <param name="errorMessage">Optional error message to display.</param>
        /// <param name="showPromptBox">Whether to show input prompt box. Default is false.</param>
        /// <param name="promptMessage">Optional prompt message to display.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "AddCustomValidation", Description = "Adds custom formula-based validation to a cell or range. Supports both InMemory and DocumentStorage modes.")]
        public AgentToolResult AddCustomValidation(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "The cell or range address")] string rangeAddress,
            [ToolParameter(Description = "The validation formula (e.g., =A1>10)")] string formula,
            [ToolParameter(Description = "Whether to show error message box")] bool showErrorBox = true,
            [ToolParameter(Description = "Optional title for error message box")] string? errorTitle = null,
            [ToolParameter(Description = "Optional error message to display")] string? errorMessage = null,
            [ToolParameter(Description = "Whether to show input prompt box")] bool showPromptBox = false,
            [ToolParameter(Description = "Optional prompt message to display")] string? promptMessage = null,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only)")] string outputFilePath = "output_AddCustomValidation.xlsx")
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(rangeAddress);
                ArgumentNullException.ThrowIfNull(formula);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                var range = worksheet.Range[rangeAddress];
                var validation = range.DataValidation;

                validation.AllowType = ExcelDataType.Formula;
                validation.FirstFormula = formula;

                // Configure error box
                validation.ShowErrorBox = showErrorBox;
                if (!string.IsNullOrEmpty(errorTitle))
                    validation.ErrorBoxTitle = errorTitle;
                if (!string.IsNullOrEmpty(errorMessage))
                    validation.ErrorBoxText = errorMessage;

                // Configure prompt box
                validation.ShowPromptBox = showPromptBox;
                if (!string.IsNullOrEmpty(promptMessage))
                    validation.PromptBoxText = promptMessage;

                // ── Save ────────────────────────────────────────────────────────
                string outputKey = outputFilePath;
                SaveDocument(outputKey, workbook);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = workbookIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Custom validation added successfully to range {rangeAddress}. Output: {outputKey}",
                    new
                    {
                        Range = rangeAddress,
                        WorksheetName = worksheetName,
                        Formula = formula,
                        OutputKey = outputKey
                    });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to add custom validation: {ex.Message}");
            }
        }

        /// <summary>
        /// Converts a time string to a decimal fraction of a day for Excel time validation.
        /// </summary>
        /// <param name="timeStr">Time string in HH:mm format (e.g., "10:00" or "18:30")</param>
        /// <returns>Decimal fraction of day, or null if invalid format</returns>
        private double? ConvertTimeToDecimal(string timeStr)
        {
            // Try to parse different time formats
            // Format 1: HH:mm (e.g., "10:00", "18:30")
            if (TimeSpan.TryParse(timeStr, out TimeSpan timeSpan))
            {
                // Convert to fraction of a day (Excel time format)
                return timeSpan.TotalHours / 24.0;
            }

            // Format 2: Decimal hours (e.g., "10.5" for 10:30)
            if (double.TryParse(timeStr, out double decimalHours))
            {
                if (decimalHours >= 0 && decimalHours < 24)
                {
                    return decimalHours / 24.0;
                }
            }

            return null;
        }
        private ExcelDataValidationComparisonOperator GetDataValidationOperator(string comparisonOperator)
        {
            if (!Enum.TryParse<ExcelDataValidationComparisonOperator>(comparisonOperator, true, out var compOperator))
            {
                if (comparisonOperator.ToLower().Contains("less"))
                {
                    if (comparisonOperator.ToLower().Contains("equal"))
                        compOperator = ExcelDataValidationComparisonOperator.LessOrEqual;
                    else
                        compOperator = ExcelDataValidationComparisonOperator.Less;
                }
                else if (comparisonOperator.ToLower().Contains("greater"))
                {
                    if (comparisonOperator.ToLower().Contains("equal"))
                        compOperator = ExcelDataValidationComparisonOperator.GreaterOrEqual;
                    else
                        compOperator = ExcelDataValidationComparisonOperator.Greater;
                }
                else
                    compOperator = ExcelDataValidationComparisonOperator.Equal;
            }
            return compOperator;
        }
    }
}
