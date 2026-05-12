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
using System.Linq;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace Syncfusion.AI.AgentTools.Excel
{
    /// <summary>
    /// Provides AI agent tools for Excel conversion operations.
    /// Handles Excel to Image, Excel to HTML, Excel to ODS, and Excel to JSON conversions including worksheet, range, chart to image, workbook to HTML, workbook/worksheet to ODS, and workbook/worksheet/range to JSON with/without schema.
    /// </summary>
    public class ExcelConversionAgentTools : AgentToolBase<IWorkbook>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelConversionAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The Excel workbook manager.</param>
        public ExcelConversionAgentTools(ExcelWorkbookManager manager)
            : base(manager, DocumentType.Excel) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelConversionAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public ExcelConversionAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Excel) { }

        /// <summary>
        /// Converts an entire worksheet to an image file.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet to convert.</param>
        /// <param name="outputPath">The output file path for the image (e.g., "output.png").</param>
        /// <param name="imageFormat">Optional image format: "PNG", "JPEG", "BMP", "GIF", "TIFF". Default is "PNG".</param>
        /// <param name="scalingMode">Optional scaling mode: "Best" (default), "NoScaling". Best provides better quality.</param>
        /// <returns>Result containing the output file path.</returns>
        [Tool(Name = "ConvertWorksheetToImage", Description = "Converts an entire worksheet to an image file. Supports PNG, JPEG, BMP, GIF, and TIFF formats. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ConvertWorksheetToImage(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet to convert")] string worksheetName,
            [ToolParameter(Description = "The cell range address (e.g., A1:D10)")] string rangeAddress,
            [ToolParameter(Description = "The output file path for the image (e.g., output.png)")] string outputPath,
            [ToolParameter(Description = "Image format: PNG (default), JPEG, BMP, GIF, TIFF")] string imageFormat = "PNG",
            [ToolParameter(Description = "Scaling mode: Best (default), NoScaling")] string scalingMode = "Best")
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(outputPath);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Find worksheet
                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                // Initialize XlsIORenderer
                var application = workbook.Application;
                if (application.XlsIORenderer == null)
                {
                    application.XlsIORenderer = new Syncfusion.XlsIORenderer.XlsIORenderer();
                }

                // Set export options
                var exportOptions = new ExportImageOptions
                {
                    ImageFormat = ParseImageFormat(imageFormat),
                    ScalingMode = ParseScalingMode(scalingMode)
                };

                // Create output directory if it doesn't exist
                if (Mode == DocumentManagerMode.InMemory)
                {
                    var outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }
                    // Convert worksheet to image
                    using (FileStream outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        worksheet.ConvertToImage(worksheet.Range[rangeAddress], exportOptions, outputStream);
                    }
                }
                else
                {
                    MemoryStream outputStream = new MemoryStream();
                    worksheet.ConvertToImage(worksheet.Range[rangeAddress], exportOptions, outputStream);
                    outputStream.Position = 0;
                    SaveFile(outputPath, outputStream);
                   
                }

                return AgentToolResult.Ok(
                    $"Worksheet '{worksheetName}' converted to image successfully",
                    new { OutputPath = outputPath, ImageFormat = imageFormat, ScalingMode = scalingMode });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to convert worksheet to image: {ex.Message}");
            }
        }

        /// <summary>
        /// Converts an Excel chart to an image file.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet containing the chart.</param>
        /// <param name="chartIndex">The index of the chart in the worksheet (0-based).</param>
        /// <param name="outputPath">The output file path for the image (e.g., "chart.png").</param>
        /// <param name="imageFormat">Optional image format: "PNG" (default), "JPEG". Default is "PNG".</param>
        /// <param name="scalingMode">Optional scaling mode: "Best" (default), "Normal".</param>
        /// <returns>Result containing the output file path.</returns>
        [Tool(Name = "ConvertChartToImage", Description = "Converts an Excel chart to an image file. Supports PNG and JPEG formats. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ConvertChartToImage(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet containing the chart")] string worksheetName,
            [ToolParameter(Description = "The index of the chart in the worksheet (0-based)")] int chartIndex,
            [ToolParameter(Description = "The output file path for the image (e.g., chart.png)")] string outputPath,
            [ToolParameter(Description = "Image format: PNG (default), JPEG")] string imageFormat = "PNG",
            [ToolParameter(Description = "Scaling mode: Best (default), Normal")] string scalingMode = "Best")
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(outputPath);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Find worksheet
                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                // Check if chart index is valid
                if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
                    return AgentToolResult.Fail($"Chart index {chartIndex} is out of range. Worksheet has {worksheet.Charts.Count} charts.");

                var chart = worksheet.Charts[chartIndex];

                // Initialize XlsIORenderer
                var application = workbook.Application;
                if (application.XlsIORenderer == null)
                {
                    application.XlsIORenderer = new Syncfusion.XlsIORenderer.XlsIORenderer();
                }

                // Set chart rendering options
                application.XlsIORenderer.ChartRenderingOptions.ImageFormat = ParseImageFormat(imageFormat);
                application.XlsIORenderer.ChartRenderingOptions.ScalingMode = ParseScalingMode(scalingMode);

                

                if (Mode == DocumentManagerMode.InMemory)
                {
                    // Create output directory if it doesn't exist
                    var outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }
                    // Convert chart to image
                    using (FileStream outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        chart.SaveAsImage(outputStream);
                    }
                }
                else
                {
                    MemoryStream outputStream = new MemoryStream();
                    chart.SaveAsImage(outputStream);
                    outputStream.Position = 0;
                    SaveFile(outputPath, outputStream);
                }

                return AgentToolResult.Ok(
                    $"Chart at index {chartIndex} from worksheet '{worksheetName}' converted to image successfully",
                    new 
                    { 
                        OutputPath = outputPath, 
                        ChartIndex = chartIndex,
                        WorksheetName = worksheetName,
                        ImageFormat = imageFormat, 
                        ScalingMode = scalingMode 
                    });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to convert chart to image: {ex.Message}");
            }
        }

        /// <summary>
        /// Converts an entire workbook to an HTML file.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="outputPath">The output file path for the HTML file (e.g., "output.html").</param>
        /// <param name="textMode">Optional text mode: "DisplayText" (shows formatted text), "Value" (shows cell values). Default is "DisplayText".</param>
        /// <returns>Result containing the output file path.</returns>
        [Tool(Name = "ConvertWorkbookToHtml", Description = "Converts an entire Excel workbook to an HTML file with styles, hyperlinks, images, and charts preserved. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ConvertWorkbookToHtml(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The output file path for the HTML file (e.g., output.html)")] string outputPath,
            [ToolParameter(Description = "Text mode: DisplayText (default - shows formatted text), Value (shows cell values)")] string textMode = "DisplayText")
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(outputPath);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                
                // Create HtmlSaveOptions
                var saveOptions = new Syncfusion.XlsIO.Implementation.HtmlSaveOptions();
                saveOptions.TextMode = ParseTextMode(textMode);

                if (Mode == DocumentManagerMode.InMemory)
                {
                    // Create output directory if it doesn't exist
                    var outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Convert workbook to HTML
                    workbook.SaveAsHtml(outputPath, saveOptions);
                }
                else
                {
                    MemoryStream outputStream = new MemoryStream();
                    // Convert workbook to HTML
                    workbook.SaveAsHtml(outputStream,saveOptions);
                    outputStream.Position = 0;
                    SaveFile(outputPath, outputStream);
                }

                return AgentToolResult.Ok(
                    $"Workbook converted to HTML successfully",
                    new { OutputPath = outputPath, TextMode = textMode, WorksheetCount = workbook.Worksheets.Count });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to convert workbook to HTML: {ex.Message}");
            }
        }

        /// <summary>
        /// Converts a specific worksheet to an HTML file.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet to convert.</param>
        /// <param name="outputPath">The output file path for the HTML file (e.g., "worksheet.html").</param>
        /// <param name="textMode">Optional text mode: "DisplayText" (shows formatted text), "Value" (shows cell values). Default is "DisplayText".</param>
        /// <returns>Result containing the output file path.</returns>
        [Tool(Name = "ConvertWorksheetToHtml", Description = "Converts a specific Excel worksheet to an HTML file with styles, hyperlinks, images, and charts preserved. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ConvertWorksheetToHtml(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet to convert")] string worksheetName,
            [ToolParameter(Description = "The output file path for the HTML file (e.g., worksheet.html)")] string outputPath,
            [ToolParameter(Description = "Text mode: DisplayText (default - shows formatted text), Value (shows cell values)")] string textMode = "DisplayText")
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(outputPath);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Find worksheet
                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                // Create output directory if it doesn't exist
                

                // Create HtmlSaveOptions
                var saveOptions = new Syncfusion.XlsIO.Implementation.HtmlSaveOptions();
                saveOptions.TextMode = ParseTextMode(textMode);

                
                if (Mode == DocumentManagerMode.InMemory)
                {
                    var outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }
                    // Convert worksheet to HTML using stream
                    using (FileStream outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        worksheet.SaveAsHtml(outputStream, saveOptions);
                    }
                }
                else
                {
                    MemoryStream outputStream = new MemoryStream();
                    // Convert workbook to HTML
                    worksheet.SaveAsHtml(outputStream, saveOptions);
                    outputStream.Position = 0;
                    SaveFile(outputPath, outputStream);
                }

                return AgentToolResult.Ok(
                    $"Worksheet '{worksheetName}' converted to HTML successfully",
                    new { OutputPath = outputPath, WorksheetName = worksheetName, TextMode = textMode });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to convert worksheet to HTML: {ex.Message}");
            }
        }

        /// <summary>
        /// Converts an entire workbook to JSON format.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="outputPath">The output file path for the JSON file (e.g., "output.json").</param>
        /// <param name="includeSchema">Optional flag to include schema in the JSON output. Default is true.</param>
        /// <returns>Result containing the output file path.</returns>
        [Tool(Name = "ConvertWorkbookToJson", Description = "Converts an entire workbook to JSON format with optional schema. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ConvertWorkbookToJson(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The output file path for the JSON file (e.g., output.json)")] string outputPath,
            [ToolParameter(Description = "Include schema in JSON output (default: true)")] bool includeSchema = true)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(outputPath);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                if (Mode == DocumentManagerMode.InMemory)
                {
                    // Create output directory if it doesn't exist
                    var outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }
                    // Convert workbook to JSON
                    workbook.SaveAsJson(outputPath, includeSchema);
                }
                else
                {
                    MemoryStream outputStream = new MemoryStream();
                    // Convert workbook to HTML
                    workbook.SaveAsJson(outputStream, includeSchema);
                    outputStream.Position = 0;
                    SaveFile(outputPath, outputStream);
                }

                return AgentToolResult.Ok(
                    $"Workbook converted to JSON format successfully",
                    new { OutputPath = outputPath, Format = "JSON", IncludeSchema = includeSchema, WorksheetCount = workbook.Worksheets.Count });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to convert workbook to JSON: {ex.Message}");
            }
        }

        /// <summary>
        /// Converts a specific worksheet to JSON format.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet to convert.</param>
        /// <param name="outputPath">The output file path for the JSON file (e.g., "worksheet.json").</param>
        /// <param name="includeSchema">Optional flag to include schema in the JSON output. Default is true.</param>
        /// <returns>Result containing the output file path.</returns>
        [Tool(Name = "ConvertWorksheetToJson", Description = "Converts a specific worksheet to JSON format with optional schema. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ConvertWorksheetToJson(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet to convert")] string worksheetName,
            [ToolParameter(Description = "The output file path for the JSON file (e.g., worksheet.json)")] string outputPath,
            [ToolParameter(Description = "Include schema in JSON output (default: true)")] bool includeSchema = true)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(outputPath);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Find worksheet
                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                if (Mode == DocumentManagerMode.InMemory)
                {
                    // Create output directory if it doesn't exist
                    var outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Convert worksheet to JSON
                    workbook.SaveAsJson(outputPath, worksheet, includeSchema);
                }
                else
                {
                    MemoryStream outputStream = new MemoryStream();
                    // Convert worksheet to JSON
                    workbook.SaveAsJson(outputStream, worksheet, includeSchema);
                    outputStream.Position = 0;
                    SaveFile(outputPath, outputStream);
                }

                return AgentToolResult.Ok(
                    $"Worksheet '{worksheetName}' converted to JSON format successfully",
                    new { OutputPath = outputPath, WorksheetName = worksheetName, Format = "JSON", IncludeSchema = includeSchema });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to convert worksheet to JSON: {ex.Message}");
            }
        }

        /// <summary>
        /// Converts a specific cell range to JSON format.
        /// </summary>
        /// <param name="workbookIdOrFilePath">The workbook ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="worksheetName">The name of the worksheet.</param>
        /// <param name="rangeAddress">The cell range address (e.g., "A1:D10").</param>
        /// <param name="outputPath">The output file path for the JSON file (e.g., "range.json").</param>
        /// <param name="includeSchema">Optional flag to include schema in the JSON output. Default is true.</param>
        /// <returns>Result containing the output file path.</returns>
        [Tool(Name = "ConvertRangeToJson", Description = "Converts a specific cell range to JSON format with optional schema. workbookIdOrFilePath: The workbook ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ConvertRangeToJson(
            [ToolParameter(Description = "The workbook ID (InMemory mode) or input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The name of the worksheet")] string worksheetName,
            [ToolParameter(Description = "The cell range address (e.g., A1:D10)")] string rangeAddress,
            [ToolParameter(Description = "The output file path for the JSON file (e.g., range.json)")] string outputPath,
            [ToolParameter(Description = "Include schema in JSON output (default: true)")] bool includeSchema = true)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(workbookIdOrFilePath);
                ArgumentNullException.ThrowIfNull(worksheetName);
                ArgumentNullException.ThrowIfNull(rangeAddress);
                ArgumentNullException.ThrowIfNull(outputPath);

                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Find worksheet
                var worksheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name == worksheetName);
                if (worksheet == null)
                    return AgentToolResult.Fail($"Worksheet not found: {worksheetName}");

                // Get the range
                var range = worksheet.Range[rangeAddress];
                if (range == null)
                    return AgentToolResult.Fail($"Invalid range address: {rangeAddress}");

                

                if (Mode == DocumentManagerMode.InMemory)
                {
                    // Create output directory if it doesn't exist
                    var outputDir = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Convert range to JSON
                    workbook.SaveAsJson(outputPath, range, includeSchema);
                }
                else
                {
                    MemoryStream outputStream = new MemoryStream();
                    // Convert range to JSON
                    workbook.SaveAsJson(outputStream, range, includeSchema);
                    outputStream.Position = 0;
                    SaveFile(outputPath, outputStream);
                }


                return AgentToolResult.Ok(
                    $"Range '{rangeAddress}' from worksheet '{worksheetName}' converted to JSON format successfully",
                    new { OutputPath = outputPath, Range = rangeAddress, WorksheetName = worksheetName, Format = "JSON", IncludeSchema = includeSchema });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to convert range to JSON: {ex.Message}");
            }
        }

        /// <summary>
        /// Converts the workbook to the file system in the specified format (DocumentStorage mode only).
        /// </summary>
        [Tool(Name = "ConvertWorkbook", Description = "Converts the workbook to the file system in the specified format. Works only in DocumentStorage mode. workbookIdOrFilePath: The input file path from storage. Supported formats: xls, xlsx, xlsm")]
        public AgentToolResult ConvertWorkbook(
            [ToolParameter(Description = "The input file path (DocumentStorage mode)")] string workbookIdOrFilePath,
            [ToolParameter(Description = "The file path to export to")] string outputPath,
            [ToolParameter(Description = "The format: xls, xlsx, xlsm, csv, tsv. Defaults to xlsx")] string? formatType = "xlsx")
        {
            try
            {

                // Open the workbook from storage
                var workbook = OpenDocument(workbookIdOrFilePath);
                if (workbook == null)
                    return AgentToolResult.Fail($"Workbook not found: {workbookIdOrFilePath}");

                // Ensure correct file extension based on version
                string extension = formatType.ToUpperInvariant() switch
                {
                    "XLS" => ".xls",
                    "XLSX" => ".xlsx",
                    "XLSM" => ".xlsm",
                    "CSV" => ".csv",
                    "TSV" => ".tsv",
                    _ => ".xlsx"
                };

                if (!outputPath.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
                {
                    outputPath = Path.ChangeExtension(outputPath, extension);
                }

                // Save the workbook to storage 
                SaveDocument(outputPath, workbook);

                return AgentToolResult.Ok($"Workbook exported successfully to {outputPath}", new { FilePath = outputPath });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to export workbook: {ex.Message}");
            }
        }

        /// <summary>
        /// Parses the text mode string to HtmlSaveOptions.GetText enum.
        /// </summary>
        private Syncfusion.XlsIO.Implementation.HtmlSaveOptions.GetText ParseTextMode(string mode)
        {
            return mode.ToUpperInvariant() switch
            {
                "DISPLAYTEXT" => Syncfusion.XlsIO.Implementation.HtmlSaveOptions.GetText.DisplayText,
                "VALUE" => Syncfusion.XlsIO.Implementation.HtmlSaveOptions.GetText.Value,
                _ => Syncfusion.XlsIO.Implementation.HtmlSaveOptions.GetText.DisplayText // Default to DisplayText
            };
        }

        /// <summary>
        /// Parses the image format string to ExportImageFormat enum.
        /// </summary>
        private ExportImageFormat ParseImageFormat(string format)
        {
            return format.ToUpperInvariant() switch
            {
                "PNG" => ExportImageFormat.Png,
                "JPEG" or "JPG" => ExportImageFormat.Jpeg,
                _ => ExportImageFormat.Png // Default to PNG
            };
        }

        /// <summary>
        /// Parses the scaling mode string to ScalingMode enum.
        /// </summary>
        private ScalingMode ParseScalingMode(string mode)
        {
            return mode.ToUpperInvariant() switch
            {
                "BEST" => ScalingMode.Best,
                "NORMAL" => ScalingMode.Normal,
                _ => ScalingMode.Best // Default to Best
            };
        }
    }
}
