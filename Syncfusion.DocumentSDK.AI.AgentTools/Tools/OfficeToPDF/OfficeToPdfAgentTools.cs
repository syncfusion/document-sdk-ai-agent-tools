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
using Syncfusion.AI.AgentTools.Word;
using Syncfusion.AI.AgentTools.Excel;
using Syncfusion.AI.AgentTools.PowerPoint;
using Syncfusion.AI.AgentTools.PDF;
using Syncfusion.DocIO.DLS;
using Syncfusion.Pdf;
using Syncfusion.Presentation;
using Syncfusion.XlsIO;

namespace Syncfusion.AI.AgentTools.OfficeToPDF
{
    /// <summary>
    /// Provides AI agent tools for converting Office documents (Word, Excel, PowerPoint) to PDF format.
    /// Handles the conversion workflow by retrieving documents from their respective managers,
    /// performing the conversion, and storing the result in the PDF manager.
    /// </summary>
    public class OfficeToPdfAgentTools : AgentToolBase<PdfDocumentBase>
    {
        private readonly DocumentManagerCollection? _managerCollection;
        private readonly string _outputDirectory;

        /// <summary>
        /// Initializes a new instance of the <see cref="OfficeToPdfAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="managerCollection">The collection of document managers.</param>
        /// <param name="outputDirectory">Optional output directory for temporary files. If null or empty, uses current directory.</param>
        public OfficeToPdfAgentTools(DocumentManagerCollection managerCollection, string? outputDirectory = null)
            : base(managerCollection.GetManager<PdfDocumentBase>(DocumentType.PDF) as PdfDocumentManager ?? throw new ArgumentException("PDF manager not found in collection"), DocumentType.PDF)
        {
            ArgumentNullException.ThrowIfNull(managerCollection);
            
            _managerCollection = managerCollection;
            _outputDirectory = string.IsNullOrEmpty(outputDirectory) 
                ? Environment.CurrentDirectory 
                : outputDirectory;

            // Ensure output directory exists
            if (!Directory.Exists(_outputDirectory))
                Directory.CreateDirectory(_outputDirectory);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="OfficeToPdfAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public OfficeToPdfAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.PDF)
        {
            _managerCollection = null;
        }

        /// <summary>
        /// Converts an Office document (Word, Excel, or PowerPoint) to PDF format.
        /// The source document must already exist in its respective manager.
        /// The resulting PDF is stored in the PDF manager for further operations.
        /// </summary>
        /// <param name="sourceDocumentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the source document.</param>
        /// <param name="sourceType">The type of source document: "Word", "Excel", or "PowerPoint".</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the generated PDF document ID.</returns>
        [Tool(
            Name = "ConvertToPDF",
            Description = "Converts an Office document (Word, Excel, or PowerPoint) to PDF format. sourceDocumentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Returns the PDF document ID.")]
        public AgentToolResult ConvertToPDF(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the source document")]
            string sourceDocumentIdOrFilePath,
            [ToolParameter(Description = "The type of source document: Word, Excel, or PowerPoint")]
            string sourceType,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(sourceDocumentIdOrFilePath);
                ArgumentNullException.ThrowIfNull(sourceType);

                // Normalize the source type to handle file extensions
                var normalizedSourceType = NormalizeSourceType(sourceType);

                // Parse the source document type
                if (!Enum.TryParse<DocumentType>(normalizedSourceType, true, out var docType))
                {
                    return AgentToolResult.Fail($"Invalid source type: {sourceType}. Supported types are: Word, Excel, PowerPoint or their file extensions (docx, xlsx, pptx, etc.)");
                }

                // Validate that the source type is supported for conversion
                if (docType == DocumentType.PDF)
                {
                    return AgentToolResult.Fail("Cannot convert PDF to PDF. Source type must be Word, Excel, or PowerPoint.");
                }

                // Perform conversion based on document type
                PdfDocument pdfDocument;
                switch (docType)
                {
                    case DocumentType.Word:
                        pdfDocument = ConvertWordToPdf(sourceDocumentIdOrFilePath);
                        break;

                    case DocumentType.Excel:
                        pdfDocument = ConvertExcelToPdf(sourceDocumentIdOrFilePath);
                        break;

                    case DocumentType.PowerPoint:
                        pdfDocument = ConvertPowerPointToPdf(sourceDocumentIdOrFilePath);
                        break;

                    default:
                        return AgentToolResult.Fail($"Unsupported document type: {docType}");
                }

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_converted.pdf";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, pdfDocument);
                if (Mode == DocumentManagerMode.InMemory)
                {
                    // Get the PDF manager
                    var pdfManager = _managerCollection.GetManager<PdfDocumentBase>(DocumentType.PDF) as PdfDocumentManager;
                    // Import the PdfDocument instance directly into the manager
                    var pdfDocumentId = pdfManager.ImportDocumentInstance(pdfDocument);
                    outputKey = pdfDocumentId;
                }

                return AgentToolResult.Ok(
                    $"Successfully converted {sourceType} document '{sourceDocumentIdOrFilePath}' to PDF: {outputKey}",
                    new { PdfDocumentId = outputKey, SourceDocumentId = sourceDocumentIdOrFilePath, SourceType = sourceType });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to convert document to PDF: {ex.Message}");
            }
        }

        /// <summary>
        /// Normalizes the source type by converting file extensions to document type names.
        /// Handles both document type names (Word, Excel, PowerPoint) and file extensions (.docx, .xlsx, .pptx, etc.).
        /// </summary>
        /// <param name="sourceType">The source type string which can be a document type name or file extension.</param>
        /// <returns>A normalized document type name (Word, Excel, or PowerPoint), or the original value if unrecognized.</returns>
        private string NormalizeSourceType(string sourceType)
        {
            // Remove leading dot if present and convert to lowercase for comparison
            var normalizedType = sourceType.TrimStart('.').ToLowerInvariant();
            
            // Check if it's already a valid document type name
            if (normalizedType == "word" || normalizedType == "excel" || normalizedType == "powerpoint")
            {
                // Return with proper casing (first letter uppercase)
                return char.ToUpper(normalizedType[0]) + normalizedType.Substring(1);
            }
            
            // Map file extensions to document types
            return normalizedType switch
            {
                // Word document extensions
                "doc" or "docx" or "docm" or "dot" or "dotx" or "dotm" or "rtf" => "Word",
                
                // Excel workbook extensions
                "xls" or "xlsx" or "xlsm" or "xlt" or "xltx" or "xltm" or "xlsb" => "Excel",
                
                // PowerPoint presentation extensions
                "pptx" or "pptm" or "potx" or "potm" => "PowerPoint",
                
                _ => sourceType
            };
        }

        /// <summary>
        /// Converts a Word document to PDF using DocIORenderer.
        /// </summary>
        private PdfDocument ConvertWordToPdf(string documentIdOrFilePath)
        {
            WordDocument? wordDocument = null;
            bool isTemporary = false;

            if (Mode == DocumentManagerMode.InMemory)
            {
                // Mode 1: InMemory - use manager collection
                var wordManager = _managerCollection?.GetManager<WordDocument>(DocumentType.Word);
                if (wordManager == null)
                {
                    throw new InvalidOperationException("Word manager is not registered.");
                }

                wordDocument = wordManager.GetDocument(documentIdOrFilePath);
                if (wordDocument == null)
                {
                    throw new InvalidOperationException($"Word document with ID '{documentIdOrFilePath}' not found in manager.");
                }
            }
            else
            {
                // Mode 2: DocumentStorage - use storage manager
                if (StorageManager!.HasDocument(documentIdOrFilePath))
                {
                    wordDocument = StorageManager.GetDocumentInstance(documentIdOrFilePath, DocumentType.Word) as WordDocument;
                    isTemporary = true;
                }
                
                if (wordDocument == null)
                {
                    throw new InvalidOperationException($"Word document '{documentIdOrFilePath}' not found in storage.");
                }
            }

            // Convert Word to PDF using DocIORenderer
            using var renderer = new Syncfusion.DocIORenderer.DocIORenderer();
            var pdfDocument = renderer.ConvertToPDF(wordDocument);

            if (isTemporary)
                wordDocument.Close();
            
            return pdfDocument;
        }

        /// <summary>
        /// Converts an Excel workbook to PDF using XlsIORenderer.
        /// </summary>
        private PdfDocument ConvertExcelToPdf(string documentIdOrFilePath)
        {
            IWorkbook? workbook = null;
            bool isTemporary = false;

            if (Mode == DocumentManagerMode.InMemory)
            {
                // Mode 1: InMemory - use manager collection
                var excelManager = _managerCollection?.GetManager<IWorkbook>(DocumentType.Excel);
                if (excelManager == null)
                {
                    throw new InvalidOperationException("Excel manager is not registered.");
                }

                workbook = excelManager.GetDocument(documentIdOrFilePath);
                if (workbook == null)
                {
                    throw new InvalidOperationException($"Excel workbook with ID '{documentIdOrFilePath}' not found in manager.");
                }
            }
            else
            {
                // Mode 2: DocumentStorage - use storage manager
                if (StorageManager!.HasDocument(documentIdOrFilePath))
                {
                    workbook = StorageManager.GetDocumentInstance(documentIdOrFilePath, DocumentType.Excel) as IWorkbook;
                    isTemporary = true;
                }
                
                if (workbook == null)
                {
                    throw new InvalidOperationException($"Excel workbook '{documentIdOrFilePath}' not found in storage.");
                }
            }

            // Convert Excel to PDF using XlsIORenderer
            var renderer = new Syncfusion.XlsIORenderer.XlsIORenderer();
            var pdfDocument = renderer.ConvertToPDF(workbook);

            if (isTemporary)
                workbook.Close();
            
            return pdfDocument;
        }

        /// <summary>
        /// Converts a PowerPoint presentation to PDF using PresentationRenderer.
        /// </summary>
        private PdfDocument ConvertPowerPointToPdf(string documentIdOrFilePath)
        {
            IPresentation? presentation = null;
            bool isTemporary = false;

            if (Mode == DocumentManagerMode.InMemory)
            {
                // Mode 1: InMemory - use manager collection
                var presentationManager = _managerCollection?.GetManager<IPresentation>(DocumentType.PowerPoint);
                if (presentationManager == null)
                {
                    throw new InvalidOperationException("PowerPoint manager is not registered.");
                }

                presentation = presentationManager.GetDocument(documentIdOrFilePath);
                if (presentation == null)
                {
                    throw new InvalidOperationException($"PowerPoint presentation with ID '{documentIdOrFilePath}' not found in manager.");
                }
            }
            else
            {
                // Mode 2: DocumentStorage - use storage manager
                if (StorageManager!.HasDocument(documentIdOrFilePath))
                {
                    presentation = StorageManager.GetDocumentInstance(documentIdOrFilePath, DocumentType.PowerPoint) as IPresentation;
                    isTemporary = true;
                }
                
                if (presentation == null)
                {
                    throw new InvalidOperationException($"PowerPoint presentation '{documentIdOrFilePath}' not found in storage.");
                }
            }

            // Convert PowerPoint to PDF using PresentationRenderer
            var pdfDocument = Syncfusion.PresentationRenderer.PresentationToPdfConverter.Convert(presentation);

            if (isTemporary)
                presentation.Close();
            
            return pdfDocument;
        }
    }
}
