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
using System.Text;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.SmartDataExtractor;
using Syncfusion.SmartTableExtractor;
using Syncfusion.SmartFormRecognizer;

namespace Syncfusion.AI.AgentTools.DataExtraction
{
    /// <summary>
    /// Provides AI agent tools for extracting structured data from PDF documents and images.
    /// Supports extraction of text, forms, tables, and images as JSON output.
    /// Combines Smart Data Extractor and Smart Table Extractor capabilities.
    /// </summary>
    public class DataExtractionAgentTools : AgentToolBase
    {
        private readonly string? _outputDirectory;
        private readonly DocumentStorageManager? _storageManager;

        /// <summary>
        /// Initializes a new instance of the <see cref="DataExtractionAgentTools"/> class.
        /// </summary>
        /// <param name="outputDirectory">Optional output directory for saving JSON files.</param>
        public DataExtractionAgentTools(string? outputDirectory = null)
        {
            _outputDirectory = outputDirectory;
        }

        /// <summary>
        /// Initializes a new instance for DocumentStorage mode.
        /// </summary>
        /// <param name="storageManager">The document storage manager to read documents from.</param>
        /// <param name="outputDirectory">Optional output directory for saving JSON files.</param>
        public DataExtractionAgentTools(DocumentStorageManager storageManager)
        {
            ArgumentNullException.ThrowIfNull(storageManager);
            _storageManager = storageManager;

        }

        /// <summary>
        /// Extracts structured data including text, forms, tables, and images from a PDF document or image file.
        /// </summary>
        /// <param name="inputFilePath">Path to input PDF or image file (.pdf, .png, .jpg, .jpeg).</param>
        /// <param name="enableFormDetection">Enable form field detection in the document.</param>
        /// <param name="enableTableDetection">Enable table detection in the document.</param>
        /// <param name="confidenceThreshold">Confidence threshold for extraction (0.0-1.0). Higher values return only high-confidence results.</param>
        /// <param name="startPage">Start page number for extraction (1-based). Use -1 for all pages.</param>
        /// <param name="endPage">End page number for extraction (1-based). Use -1 for all pages.</param>
        /// <param name="detectSignatures">Detect signature fields in forms.</param>
        /// <param name="detectTextboxes">Detect textbox fields in forms.</param>
        /// <param name="detectCheckboxes">Detect checkbox fields in forms.</param>
        /// <param name="detectRadioButtons">Detect radio button fields in forms.</param>
        /// <param name="detectBorderlessTables">Enable detection of border-less tables.</param>
        /// <param name="outputFilePath">Optional output JSON file path to save the extracted data.</param>
        /// <returns>Result containing the extracted JSON data or error message.</returns>
        [Tool(Name = "ExtractDataAsJson",
              Description = "Extracts structured data including text, forms, tables, and images from a PDF document or image file and returns as JSON string. Supports form field detection with configurable options (signatures, textboxes, checkboxes, radio buttons) and table detection with border-less table support. Page range and confidence threshold can be specified for fine-tuned extraction.")]
        public AgentToolResult ExtractDataAsJson(
            [ToolParameter(Description = "Path to input PDF or image file (.pdf, .png, .jpg, .jpeg)")]
            string inputFilePath,

            [ToolParameter(Description = "Enable form field detection in the document")]
            bool enableFormDetection = true,

            [ToolParameter(Description = "Enable table detection in the document")]
            bool enableTableDetection = true,

            [ToolParameter(Description = "Confidence threshold for extraction (0.0-1.0). Higher values return only high-confidence results")]
            double confidenceThreshold = 0.6,

            [ToolParameter(Description = "Start page number for extraction (1-based). Use -1 for all pages")]
            int startPage = -1,

            [ToolParameter(Description = "End page number for extraction (1-based). Use -1 for all pages")]
            int endPage = -1,

            [ToolParameter(Description = "Detect signature fields in forms")]
            bool detectSignatures = true,

            [ToolParameter(Description = "Detect textbox fields in forms")]
            bool detectTextboxes = true,

            [ToolParameter(Description = "Detect checkbox fields in forms")]
            bool detectCheckboxes = true,

            [ToolParameter(Description = "Detect radio button fields in forms")]
            bool detectRadioButtons = true,

            [ToolParameter(Description = "Enable detection of border-less tables")]
            bool detectBorderlessTables = true,

            [ToolParameter(Description = "Optional output JSON file path to save the extracted data")]
            string? outputFilePath = null)
        {
            try
            {
                // Validate input parameters
                if (string.IsNullOrWhiteSpace(inputFilePath))
                    return AgentToolResult.Fail("File path cannot be null or empty.");

                if (confidenceThreshold < 0.0 || confidenceThreshold > 1.0)
                    return AgentToolResult.Fail("Confidence threshold must be between 0.0 and 1.0.");

                string jsonData;

                // Open the input as a stream (supports: storage manager or local file path)
                string? resolvedName = null;
                using (var stream = GetInputStream(inputFilePath, out resolvedName))
                {
                    // Initialize the Smart Data Extractor
                    DataExtractor extractor = new DataExtractor();

                    // Configure extraction options
                    extractor.EnableFormDetection = enableFormDetection;
                    extractor.EnableTableDetection = enableTableDetection;
                    extractor.ConfidenceThreshold = confidenceThreshold;

                    // Configure page range if specified
                    var pageRange = BuildPageRange(startPage, endPage);
                    if (pageRange != null)
                    {
                        extractor.PageRange = pageRange;
                    }

                    // Configure form recognition options if form detection is enabled
                    if (enableFormDetection)
                    {
                        FormRecognizeOptions formOptions = new FormRecognizeOptions();
                        
                        if (pageRange != null)
                        {
                            formOptions.PageRange = pageRange;
                        }
                        
                        formOptions.ConfidenceThreshold = confidenceThreshold;
                        formOptions.DetectSignatures = detectSignatures;
                        formOptions.DetectTextboxes = detectTextboxes;
                        formOptions.DetectCheckboxes = detectCheckboxes;
                        formOptions.DetectRadioButtons = detectRadioButtons;

                        extractor.FormRecognizeOptions = formOptions;
                    }

                    // Configure table extraction options if table detection is enabled
                    if (enableTableDetection)
                    {
                        TableExtractionOptions tableOptions = new TableExtractionOptions();
                        
                        if (pageRange != null)
                        {
                            tableOptions.PageRange = pageRange;
                        }
                        
                        tableOptions.ConfidenceThreshold = confidenceThreshold;
                        tableOptions.DetectBorderlessTables = detectBorderlessTables;

                        extractor.TableExtractionOptions = tableOptions;
                    }

                    // Extract data as JSON
                    jsonData = extractor.ExtractDataAsJson(stream);
                }

                // Save to file if output path is provided
                if (!string.IsNullOrEmpty(outputFilePath))
                {
                    string fullPath = outputFilePath;
                    if (_storageManager == null)
                    {
                        fullPath = ResolveOutputPath(outputFilePath);
                    }
                   
                    SaveJsonToFile(jsonData, fullPath);
                }

                string message = $"Data extracted successfully from {Path.GetFileName(resolvedName ?? inputFilePath)}";
                if (!string.IsNullOrEmpty(outputFilePath))
                    message += $". Saved to {outputFilePath}";

                return AgentToolResult.Ok(message, jsonData);
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to extract data: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts only table data from a PDF document and returns as JSON string.
        /// </summary>
        /// <param name="inputFilePath">Path to input PDF file.</param>
        /// <param name="detectBorderlessTables">Enable detection of border-less tables.</param>
        /// <param name="confidenceThreshold">Confidence threshold for table extraction (0.0-1.0). Higher values return only high-confidence tables.</param>
        /// <param name="startPage">Start page number for extraction (1-based). Use -1 for all pages.</param>
        /// <param name="endPage">End page number for extraction (1-based). Use -1 for all pages.</param>
        /// <param name="outputFilePath">Optional output JSON file path to save the extracted table data.</param>
        /// <returns>Result containing the extracted table JSON data or error message.</returns>
        [Tool(Name = "ExtractTableAsJson",
              Description = "Extracts only table data from a PDF document and returns as JSON string. Optimized for table-focused extraction with support for border-less table detection, page range specification, and confidence thresholding. Use this method when you only need table data without form fields or other content.")]
        public AgentToolResult ExtractTableAsJson(
            [ToolParameter(Description = "Path to input PDF file")]
            string inputFilePath,

            [ToolParameter(Description = "Enable detection of border-less tables")]
            bool detectBorderlessTables = true,

            [ToolParameter(Description = "Confidence threshold for table extraction (0.0-1.0). Higher values return only high-confidence tables")]
            double confidenceThreshold = 0.6,

            [ToolParameter(Description = "Start page number for extraction (1-based). Use -1 for all pages")]
            int startPage = -1,

            [ToolParameter(Description = "End page number for extraction (1-based). Use -1 for all pages")]
            int endPage = -1,

            [ToolParameter(Description = "Optional output JSON file path to save the extracted table data")]
            string? outputFilePath = null)
        {
            try
            {
                // Validate input parameters
                if (string.IsNullOrWhiteSpace(inputFilePath))
                    return AgentToolResult.Fail("File path cannot be null or empty.");

                if (confidenceThreshold < 0.0 || confidenceThreshold > 1.0)
                    return AgentToolResult.Fail("Confidence threshold must be between 0.0 and 1.0.");

                string jsonData;

                // Open the input as a stream (supports: DocumentStorageManager or local file path)
                string? resolvedName = null;
                using (var stream = GetInputStream(inputFilePath, out resolvedName))
                {
                    // Initialize the Smart Table Extractor
                    TableExtractor extractor = new TableExtractor();

                    // Configure table extraction options
                    TableExtractionOptions options = new TableExtractionOptions();
                    options.DetectBorderlessTables = detectBorderlessTables;
                    options.ConfidenceThreshold = confidenceThreshold;

                    // Configure page range if specified
                    var pageRange = BuildPageRange(startPage, endPage);
                    if (pageRange != null)
                    {
                        options.PageRange = pageRange;
                    }

                    // Assign the configured options to the extractor
                    extractor.TableExtractionOptions = options;

                    // Extract table data as JSON
                    jsonData = extractor.ExtractTableAsJson(stream);
                }

                // Save to file if output path is provided
                if (!string.IsNullOrEmpty(outputFilePath))
                {
                    string fullPath = ResolveOutputPath(outputFilePath);
                    SaveJsonToFile(jsonData, fullPath);
                }

                string message = $"Tables extracted successfully from {Path.GetFileName(resolvedName ?? inputFilePath)}";
                if (!string.IsNullOrEmpty(outputFilePath))
                    message += $". Saved to {outputFilePath}";

                return AgentToolResult.Ok(message, jsonData);
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to extract tables: {ex.Message}");
            }
        }
        /// <summary>
        /// Extracts only form field data from a PDF document and returns as JSON string.
        /// </summary>
        /// <param name="inputFilePath">Path to input PDF file.</param>
        /// <param name="detectSignatures">Detect signature fields in forms.</param>
        /// <param name="detectTextboxes">Detect textbox fields in forms.</param>
        /// <param name="detectCheckboxes">Detect checkbox fields in forms.</param>
        /// <param name="detectRadioButtons">Detect radio button fields in forms.</param>
		/// <param name="confidenceThreshold">Confidence threshold for form recognition (0.0-1.0). Higher values return only high-confidence tables.</param>
        /// <param name="startPage">Start page number for recognition (1-based). Use -1 for all pages.</param>
        /// <param name="endPage">End page number for recognition (1-based). Use -1 for all pages.</param>
        /// <param name="outputFilePath">Optional output JSON file path to save the recognized form fields data.</param>
        /// <returns>Result containing the recognized form fields JSON data or error message.</returns>
        [Tool(Name = "ExtractFormFieldsAsJson",
              Description = "Extracts only form data from a PDF document and returns as JSON string. Optimized for form-focused extraction with support  page range specification, and confidence thresholding. Use this method when you only need form data without table or other content.")]
        public AgentToolResult ExtractFormFieldsAsJson(
            [ToolParameter(Description = "Path to input PDF file")]
            string inputFilePath,

            [ToolParameter(Description = "Detect signature fields in forms")]
            bool detectSignatures = true,

            [ToolParameter(Description = "Detect textbox fields in forms")]
            bool detectTextboxes = true,

            [ToolParameter(Description = "Detect checkbox fields in forms")]
            bool detectCheckboxes = true,

            [ToolParameter(Description = "Detect radio button fields in forms")]
            bool detectRadioButtons = true,

            [ToolParameter(Description = "Confidence threshold for form recognition (0.0-1.0). Higher values return only high-confidence form fields")]
            double confidenceThreshold = 0.6,

            [ToolParameter(Description = "Start page number for recognition (1-based). Use -1 for all pages")]
            int startPage = -1,

            [ToolParameter(Description = "End page number for recognition (1-based). Use -1 for all pages")]
            int endPage = -1,

            [ToolParameter(Description = "Optional output JSON file path to save the recognized form fields")]
            string? outputFilePath = null)
        {
            try
            {
                // Validate input parameters
                if (string.IsNullOrWhiteSpace(inputFilePath))
                    return AgentToolResult.Fail("File path cannot be null or empty.");

                if (confidenceThreshold < 0.0 || confidenceThreshold > 1.0)
                    return AgentToolResult.Fail("Confidence threshold must be between 0.0 and 1.0.");

                string jsonData;

                // Open the input as a stream (supports: DocumentStorageManager or local file path)
                string? resolvedName = null;
                using (var stream = GetInputStream(inputFilePath, out resolvedName))
                {
                    // Initialize the Smart Form Recognizer
                    FormRecognizer recognizer = new FormRecognizer();

                    // Configure form recognize options
                    FormRecognizeOptions options = new FormRecognizeOptions();
                    options.DetectTextboxes = detectTextboxes;
                    options.DetectRadioButtons = detectRadioButtons;
                    options.DetectCheckboxes = detectCheckboxes;
                    options.DetectSignatures = detectSignatures;
                    options.ConfidenceThreshold = confidenceThreshold;

                    // Configure page range if specified
                    var pageRange = BuildPageRange(startPage, endPage);
                    if (pageRange != null)
                    {
                        options.PageRange = pageRange;
                    }

                    // Assign the configured options to the recognizer
                    recognizer.FormRecognizeOptions = options;

                    // Recognize form fiels as JSON
                    jsonData = recognizer.RecognizeFormAsJson(stream);
                }

                // Save to file if output path is provided
                if (!string.IsNullOrEmpty(outputFilePath))
                {
                    string fullPath = ResolveOutputPath(outputFilePath);
                    SaveJsonToFile(jsonData, fullPath);
                }

                string message = $"Forms extracted successfully from {Path.GetFileName(resolvedName ?? inputFilePath)}";
                if (!string.IsNullOrEmpty(outputFilePath))
                    message += $". Saved to {outputFilePath}";

                return AgentToolResult.Ok(message, jsonData);
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to extract form field: {ex.Message}");
            }
        }
        /// <summary>
        /// Converts a PDF document or image file into Markdown by extracting structured data.
        /// </summary>
        /// <param name="inputFilePath">Path to input PDF or image file.</param>
        /// <param name="enableTableDetection">Enable table detection in the document.</param>
        /// <param name="confidenceThreshold">Confidence threshold for extraction (0.0-1.0). Higher values return only high-confidence results.</param>
        /// <param name="startPage">Start page number for extraction (1-based). Use -1 for all pages.</param>
        /// <param name="endPage">End page number for extraction (1-based). Use -1 for all pages.</param>
        /// <param name="detectBorderlessTables">Enable detection of border-less tables.</param>
        /// <param name="outputFilePath">Optional output Markdown (.md) file path.</param>
        /// <returns>Result containing the extracted Markdown data or error message.</returns>
        [Tool(Name = "ConvertPdfToMarkdown",
            Description = "Converts structured information from PDF documents and scanned images into Markdown (MD) format. It analyzes text blocks, tables and headers to preserve the original layout and formatting in the generated Markdown output.")]
        public AgentToolResult ConvertPdfToMarkdown(
            [ToolParameter(Description = "Path to input PDF file")]
            string inputFilePath,

            [ToolParameter(Description = "Enable table detection in the document")]
            bool enableTableDetection = true,

            [ToolParameter(Description = "Confidence threshold for extraction (0.0-1.0). Higher values return only high-confidence results")]
            double confidenceThreshold = 0.6,

            [ToolParameter(Description = "Start page number for extraction (1-based). Use -1 for all pages")]
            int startPage = -1,

            [ToolParameter(Description = "End page number for extraction (1-based). Use -1 for all pages")]
            int endPage = -1,

            [ToolParameter(Description = "Enable detection of border-less tables")]
            bool detectBorderlessTables = true,

            [ToolParameter(Description = "Optional output Markdown (.md) file path")]
            string? outputFilePath = null)
        {
            try
            {
                // Validate input parameters
                if (string.IsNullOrWhiteSpace(inputFilePath))
                    return AgentToolResult.Fail("File path cannot be null or empty.");

                if (confidenceThreshold < 0.0 || confidenceThreshold > 1.0)
                    return AgentToolResult.Fail("Confidence threshold must be between 0.0 and 1.0.");

                string markdownData;

                // Open the input as a stream (supports: storage manager or local file path)
                string? resolvedName = null;
                using (var stream = GetInputStream(inputFilePath, out resolvedName))
                {
                    // Initialize the Smart Data Extractor
                    DataExtractor extractor = new DataExtractor();

                    // Configure extraction options
                    extractor.EnableTableDetection = enableTableDetection;
                    extractor.ConfidenceThreshold = confidenceThreshold;

                    // Configure page range if specified
                    var pageRange = BuildPageRange(startPage, endPage);
                    if (pageRange != null)
                    {
                        extractor.PageRange = pageRange;
                    }

                    // Configure table extraction options
                    if (enableTableDetection)
                    {
                        TableExtractionOptions tableOptions = new TableExtractionOptions
                        {
                            PageRange = pageRange,
                            ConfidenceThreshold = confidenceThreshold,
                            DetectBorderlessTables = detectBorderlessTables
                        };

                        extractor.TableExtractionOptions = tableOptions;
                    }

                    markdownData = extractor.ExtractDataAsMarkdown(stream);
                }

                // Save to file if output path is provided
                if (!string.IsNullOrEmpty(outputFilePath))
                {
                    string fullPath = outputFilePath;
                    if (_storageManager == null)
                    {
                        fullPath = ResolveOutputPath(outputFilePath);
                    }

                    SaveMarkdownToFile(markdownData, fullPath);
                }

                string message = $"Markdown converted successfully from {Path.GetFileName(resolvedName ?? inputFilePath)}";
                if (!string.IsNullOrEmpty(outputFilePath))
                    message += $". Saved to {outputFilePath}";

                return AgentToolResult.Ok(message, markdownData);
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to convert PDF to Markdown: {ex.Message}");
            }
        }
        /// <summary>
        /// Converts tables from a PDF document into Markdown format.
        /// </summary>
        /// <param name="inputFilePath">Path to input PDF or image file.</param>
        /// <param name="confidenceThreshold">Confidence threshold for table extraction (0.0-1.0).</param>
        /// <param name="startPage">Start page number for extraction (1-based). Use -1 for all pages.</param>
        /// <param name="endPage">End page number for extraction (1-based). Use -1 for all pages.</param>
        /// <param name="detectBorderlessTables">Enable detection of border-less tables.</param>
        /// <param name="outputFilePath">Optional output Markdown (.md) file path.</param>
        /// <returns>Result containing the extracted table Markdown data or error message.</returns>
        [Tool(Name = "ConvertPdfTableToMarkdown",
            Description = "Converts tables from PDF documents and scanned images into Markdown (MD) format. It analyzes visual table structures, including bordered and border-less tables, enabling developers to programmatically convert tabular content into clean and well-structured Markdown representations.")]
        public AgentToolResult ConvertPdfTableToMarkdown(
            [ToolParameter(Description = "Path to input PDF file")]
            string inputFilePath,

            [ToolParameter(Description = "Confidence threshold for table extraction (0.0-1.0)")]
            double confidenceThreshold = 0.6,

            [ToolParameter(Description = "Start page number for extraction (1-based). Use -1 for all pages")]
            int startPage = -1,

            [ToolParameter(Description = "End page number for extraction (1-based). Use -1 for all pages")]
            int endPage = -1,

            [ToolParameter(Description = "Enable detection of border-less tables")]
            bool detectBorderlessTables = true,

            [ToolParameter(Description = "Optional output Markdown (.md) file path")]
            string? outputFilePath = null)
        {
            try
            {
                // Validate input parameters
                if (string.IsNullOrWhiteSpace(inputFilePath))
                    return AgentToolResult.Fail("File path cannot be null or empty.");

                if (confidenceThreshold < 0.0 || confidenceThreshold > 1.0)
                    return AgentToolResult.Fail("Confidence threshold must be between 0.0 and 1.0.");

                string markdownData;

                // Open input stream
                string? resolvedName = null;
                using (var stream = GetInputStream(inputFilePath, out resolvedName))
                {
                    // Initialize Table Extractor
                    TableExtractor extractor = new TableExtractor();

                    // Configure table extraction options
                    TableExtractionOptions options = new TableExtractionOptions
                    {
                        DetectBorderlessTables = detectBorderlessTables,
                        ConfidenceThreshold = confidenceThreshold
                    };

                    // Configure page range if specified
                    var pageRange = BuildPageRange(startPage, endPage);
                    if (pageRange != null)
                    {
                        options.PageRange = pageRange;
                    }

                    extractor.TableExtractionOptions = options;
                    markdownData = extractor.ExtractTableAsMarkdown(stream);
                }

                // Save to file if output path is provided
                if (!string.IsNullOrEmpty(outputFilePath))
                {
                    string fullPath = outputFilePath;
                    if (_storageManager == null)
                    {
                        fullPath = ResolveOutputPath(outputFilePath);
                    }

                    SaveMarkdownToFile(markdownData, fullPath);
                }

                string message = $"Table Markdown converted successfully from {Path.GetFileName(resolvedName ?? inputFilePath)}";
                if (!string.IsNullOrEmpty(outputFilePath))
                    message += $". Saved to {outputFilePath}";

                return AgentToolResult.Ok(message, markdownData);
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to extract tables as Markdown: {ex.Message}");
            }
        }
        #region Private Helper Methods

        /// <summary>
        /// Resolves the output file path using the configured output directory.
        /// </summary>
        /// <param name="outputFilePath">The output file path (can be relative or absolute).</param>
        /// <returns>The resolved full path.</returns>
        private string ResolveOutputPath(string? outputFilePath)
        {
            if (string.IsNullOrEmpty(outputFilePath))
                return string.Empty;

            if (!Path.IsPathRooted(outputFilePath) && !string.IsNullOrEmpty(_outputDirectory))
                return Path.Combine(_outputDirectory, outputFilePath);

            return outputFilePath;
        }

        /// <summary>
        /// <summary>
        /// Opens the input source as a readable stream. Supports:
        /// - DocumentStorage (when constructed with a <see cref="DocumentStorageManager"/>)
        /// - Local file system path
        /// Returns a stream that the caller must dispose. Also returns a resolved name for messaging.
        /// </summary>
        private Stream GetInputStream(string inputPath, out string? resolvedName)
        {
            resolvedName = null;

            // If storage manager was provided and the path exists there, use it
            if (_storageManager != null)
            {
                if (_storageManager.HasDocument(inputPath))
                {
                    var fileStream = _storageManager.GetDocumentStream(inputPath);
                    if (fileStream != null)
                    {
                        resolvedName = inputPath;
                        return fileStream;
                    }
                }
            }

            // Fallback to local file system path
            if (File.Exists(inputPath))
            {
                resolvedName = inputPath;
                return new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            }

            throw new FileNotFoundException("Input document not found in storage or local file system.", inputPath);
        }

        /// <summary>
        /// Saves JSON data to a file.
        /// </summary>
        /// <param name="jsonData">The JSON data to save.</param>
        /// <param name="filePath">The file path to save to.</param>
        private void SaveJsonToFile(string jsonData, string filePath)
        {
            if (_storageManager != null)
            {
                
                var byteArray = Encoding.UTF8.GetBytes(jsonData);
                var memoryStream = new MemoryStream(byteArray);
                _storageManager.WriteRawStream(filePath, memoryStream);
            } else
            {
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    Directory.CreateDirectory(directory);

                File.WriteAllText(filePath, jsonData, Encoding.UTF8);
            }
               
        }
        /// <summary>
        /// Saves text (Markdown) data to a file.
        /// </summary>
        /// <param name="textData">The text / Markdown data to save.</param>
        /// <param name="filePath">The file path to save to.</param>
        private void SaveMarkdownToFile(string textData, string filePath)
        {
            if (_storageManager != null)
            {
                var byteArray = Encoding.UTF8.GetBytes(textData);
                var memoryStream = new MemoryStream(byteArray);
                _storageManager.WriteRawStream(filePath, memoryStream);
            }
            else
            {
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    Directory.CreateDirectory(directory);

                File.WriteAllText(filePath, textData, Encoding.UTF8);
            }
        }
        /// <summary>
        /// Builds a page range array for extraction.
        /// </summary>
        /// <param name="startPage">Start page number (1-based, -1 for all).</param>
        /// <param name="endPage">End page number (1-based, -1 for all).</param>
        /// <returns>Page range array or null for all pages.</returns>
        /// <exception cref="ArgumentException">Thrown when page range is invalid.</exception>
        private int[,]? BuildPageRange(int startPage, int endPage)
        {
            // If both are -1, extract all pages
            if (startPage == -1 && endPage == -1)
                return null;

            // Default start page to 1 if not specified
            if (startPage == -1)
                startPage = 1;

            // If end page is -1, it means extract to the end
            // We'll use a large number and let the extractor handle it
            if (endPage == -1)
                endPage = int.MaxValue;

            // Validate page range
            if (startPage < 1)
                throw new ArgumentOutOfRangeException(nameof(startPage),
                    "Page numbers must be 1-based (minimum value is 1).");

            if (startPage > endPage)
                throw new ArgumentException("Start page must be less than or equal to end page.");

            return new int[,] { { startPage, endPage } };
        }

        #endregion
    }
}
