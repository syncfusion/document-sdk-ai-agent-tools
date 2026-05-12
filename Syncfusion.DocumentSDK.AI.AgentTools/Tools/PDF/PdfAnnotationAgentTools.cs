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

using SkiaSharp;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.Drawing;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Interactive;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Pdf.Security;
using System;
using System.IO;
using System.Text;

namespace Syncfusion.AI.AgentTools.PDF
{
    /// <summary>
    /// Provides AI agent tools for PDF annotation and modification operations.
    /// Handles watermarking, digital signatures, and annotation management.
    /// </summary>
    public class PdfAnnotationAgentTools : AgentToolBase<PdfDocumentBase>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PdfAnnotationAgentTools"/> class (Mode 1 � InMemory).
        /// </summary>
        /// <param name="manager">The PDF document manager.</param>
        public PdfAnnotationAgentTools(PdfDocumentManager manager)
            : base(manager, DocumentType.PDF) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="PdfAnnotationAgentTools"/> class (Mode 2 � DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public PdfAnnotationAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.PDF) { }

        /// <summary>
        /// Applies a configurable text watermark to all pages of a PDF document.
        /// Supports opacity, rotation, color, and custom positioning.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="watermarkText">The watermark text to apply.</param>
        /// <param name="rotation">Optional rotation angle in degrees. Defaults to 45.</param>
        /// <param name="locationX">X coordinate. Use -1 to center horizontally.</param>
        /// <param name="locationY">Y coordinate. Use -1 to center vertically.</param>
        /// <param name="watermarkColor">Optional watermark color (RGB). Defaults to gray.</param>
        /// <param name="opacity">Opacity percentage (0–100). Defaults to 50.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the document ID and watermark details.</returns>
        [Tool(
        Name = "WatermarkPdf",
        Description = "Applies a configurable text watermark to all pages of a PDF document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Supports opacity, rotation, color, and positioning."
        )]
        public AgentToolResult WatermarkPdf(
        [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
        string documentIdOrFilePath,

        [ToolParameter(Description = "The watermark text to apply.")]
        string watermarkText,

        [ToolParameter(Description = "Optional rotation angle in degrees. Defaults to 45.")]
        int? rotation = null,

        [ToolParameter(Description = "X coordinate. Use -1 to center horizontally.")]
        float locationX = -1,

        [ToolParameter(Description = "Y coordinate. Use -1 to center vertically.")]
        float locationY = -1,

        [ToolParameter(Description = "Optional watermark color (RGB). Defaults to gray. The color ranges must be 0 to 255. Example: Blue [0,0,255]")]
        byte[]? watermarkColor = null,

        [ToolParameter(Description = "Opacity percentage (0�100). Defaults to 50.")]
        float opacity = 50f,

        [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
        string? outputFilePath = null
)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);

                if (string.IsNullOrWhiteSpace(watermarkText))
                    return AgentToolResult.Fail("Watermark text cannot be empty.");

                if (opacity < 0 || opacity > 100)
                    return AgentToolResult.Fail("Opacity must be between 0 and 100.");

                if ((locationX < 0 && locationX != -1) || (locationY < 0 && locationY != -1))
                    return AgentToolResult.Fail("Location coordinates must be -1 or non-negative values.");

                // -- Open --------------------------------------------------------
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                bool isNewDocument = false;

                if (document is PdfDocument)
                {
                    MemoryStream reloadStream = new MemoryStream();
                    document.Save(reloadStream);
                    if (Mode == DocumentManagerMode.InMemory)
                        InMemoryManager!.RemoveDocument(documentIdOrFilePath);
                    document = new PdfLoadedDocument(reloadStream);
                    isNewDocument = true;
                }

                PdfLoadedDocument loadedDocument = document as PdfLoadedDocument;

                // ---- Normalized Parameters ----------------------------------------
                int appliedRotation = -rotation ?? (-45);
                float appliedOpacity = opacity / 100f;
                PdfColor appliedColor = watermarkColor == null || watermarkColor.Length < 3 
                    ? new PdfColor(128, 128, 128) // Default to gray
                    : new PdfColor(watermarkColor[0], watermarkColor[1], watermarkColor[2]);

                PdfFont font = new PdfStandardFont(PdfFontFamily.Helvetica, 60);
                int pageCount = 0;

                // ---- Apply Watermark ----------------------------------------------
                foreach (PdfPageBase page in loadedDocument.Pages)
                {
                    PdfGraphics graphics = page.Graphics;
                    SizeF pageSize = page.Size;

                    PdfGraphicsState state = graphics.Save();
                    graphics.SetTransparency(appliedOpacity);

                    float x = locationX == -1 ? pageSize.Width / 2f : locationX;
                    float y = locationY == -1 ? pageSize.Height / 2f : locationY;

                    graphics.TranslateTransform(x, y);
                    graphics.RotateTransform(appliedRotation);

                    SizeF textSize = font.MeasureString(watermarkText);
                    PdfBrush brush = new PdfSolidBrush(appliedColor);

                    graphics.DrawString(
                        watermarkText,
                        font,
                        brush,
                        -textSize.Width / 2f,
                        -textSize.Height / 2f
                    );

                    graphics.Restore(state);
                    pageCount++;
                }

                // -- Save --------------------------------------------------------
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_watermarked.pdf";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, loadedDocument);
                if (Mode == DocumentManagerMode.InMemory)
                {
                    if (isNewDocument)
                        outputKey = ((PdfDocumentManager)InMemoryManager!).ImportDocumentInstance(loadedDocument);
                    else
                        outputKey = documentIdOrFilePath;
                }

                // ---- Result --------------------------------------------------------
                return AgentToolResult.Ok(
                    $"Watermark applied successfully to {pageCount} page(s) into document {outputKey}.",
                    new
                    {
                        DocumentId = outputKey,
                        PagesProcessed = pageCount,
                        WatermarkText = watermarkText,
                        RotationDegrees = appliedRotation,
                        OpacityPercent = opacity,
                        Color = appliedColor
                    }
                );
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to apply watermark: {ex.Message}");
            }
        }

        /// <summary>
        /// Exports annotations from a PDF document in the specified format.
        /// </summary>
        /// <remarks>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// Supported formats:
        /// <list type="bullet">
        /// <item>XFDF</item>
        /// <item>FDF</item>
        /// <item>JSON</item>
        /// </list>
        /// 
        /// Export behavior:
        /// <list type="bullet">
        /// <item>
        /// If <paramref name="exportFilePath"/> is a folder, a file is auto-created inside it.
        /// </item>
        /// <item>
        /// If <paramref name="exportFilePath"/> is a file path, it is overwritten.
        /// </item>
        /// <item>
        /// If no path is provided, annotation data is returned as a UTF-8 string.
        /// </item>
        /// </list>
        /// </remarks>
        [Tool(
            Name = "ExportAnnotations",
            Description = "Exports annotations from a PDF document into XFDF, FDF, or JSON format. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode)."
        )]
        public AgentToolResult ExportAnnotations(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
        string documentIdOrFilePath,

            [ToolParameter(Description = "The export format (XFDF, FDF, JSON)")]
        AnnotationDataFormat format,

            [ToolParameter(Description = "Optional: Export file or folder path")]
        string? exportFilePath = null
        )
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);

                

                var document = OpenDocument(documentIdOrFilePath);

                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                bool isReloaded = false;
                PdfLoadedDocument loadedDocument;

                // ✅ Ensure PdfLoadedDocument
                if (document is PdfLoadedDocument pdfLoaded)
                {
                    loadedDocument = pdfLoaded;
                }
                else
                {
                    var reloadStream = new MemoryStream();
                    document.Save(reloadStream);
                    reloadStream.Position = 0;

                    if (Mode == DocumentManagerMode.InMemory)
                        InMemoryManager!.RemoveDocument(documentIdOrFilePath);
                    loadedDocument = new PdfLoadedDocument(reloadStream);

                    isReloaded = true;
                }

                using var memoryStream = new MemoryStream();
                loadedDocument.ExportAnnotations(memoryStream, format);
                memoryStream.Position = 0;

                string resolvedDocId = documentIdOrFilePath;
                if (isReloaded && Mode == DocumentManagerMode.InMemory)
                    resolvedDocId = ((PdfDocumentManager)InMemoryManager!).ImportDocumentInstance(loadedDocument);

                if (!string.IsNullOrWhiteSpace(exportFilePath))
                {
                    string finalFilePath = ResolveExportPath(exportFilePath, documentIdOrFilePath, format);

                    if (Mode == DocumentManagerMode.InMemory)
                    {
                        using var fileStream = new FileStream(finalFilePath, FileMode.Create, FileAccess.Write, FileShare.None);
                        memoryStream.CopyTo(fileStream);
                    }
                    else
                    {
                        using (var annotationStream = new MemoryStream())
                        {
                            memoryStream.Position = 0;
                            memoryStream.CopyTo(annotationStream);
                            memoryStream.Position = 0;
                            annotationStream.Position = 0;
                            SaveFile(finalFilePath, annotationStream);
                        }
                    }

                    return AgentToolResult.Ok(
                        "Annotations exported successfully.",
                        new { DocumentId = resolvedDocId, Format = format.ToString(), ExportPath = finalFilePath });
                }

                string annotationData = System.Text.Encoding.UTF8.GetString(memoryStream.ToArray());
               

                return AgentToolResult.Ok(
                    "Annotations exported successfully.",
                    new { DocumentId = resolvedDocId, Format = format.ToString(), AnnotationData = annotationData });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to export annotations: {ex.Message}");
            }
        }

        /// <summary>
        /// Resolves export file path from folder or file input.
        /// </summary>
        private string ResolveExportPath(string exportPath, string documentId, AnnotationDataFormat format)
        {
            string extension = format switch
            {
                AnnotationDataFormat.XFdf => ".xfdf",
                AnnotationDataFormat.Fdf => ".fdf",
                AnnotationDataFormat.Json => ".json",
                _ => ".txt"
            };
            if(Mode == DocumentManagerMode.InMemory)
            {
                // Folder path or path without extension
                if (Directory.Exists(exportPath) || string.IsNullOrEmpty(Path.GetExtension(exportPath)))
                {
                    Directory.CreateDirectory(exportPath);
                    return Path.Combine(exportPath, $"Annotations_{documentId}{extension}");
                }

                string? directory = Path.GetDirectoryName(exportPath);
                if (!string.IsNullOrWhiteSpace(directory))
                    Directory.CreateDirectory(directory);
            }
            else
            {
                // For DocumentStorage mode, ensure the path has proper extension
                if (string.IsNullOrEmpty(Path.GetExtension(exportPath)))
                {
                    return Path.Combine(exportPath,$"Annotations_Export{extension}");
                }
            }

            return exportPath;
        }

        /// <summary>
        /// Imports annotations into a PDF document from an XFDF, FDF, or JSON file.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="format">The import format (XFDF, FDF, JSON).</param>
        /// <param name="importFilePath">Annotation file path (XFDF, FDF, or JSON file).</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the document ID and import details.</returns>
        [Tool(
            Name = "ImportAnnotations",
            Description = "Imports annotations into a PDF document from XFDF, FDF, or JSON. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode)."
        )]
        public AgentToolResult ImportAnnotations(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
        string documentIdOrFilePath,

            [ToolParameter(Description = "The import format (XFDF, FDF, JSON)")]
        AnnotationDataFormat format,

            [ToolParameter(Description = "Annotation file path (XFDF, FDF, or JSON file)")]
        string importFilePath,

            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
        string? outputFilePath = null
        )
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);
                ArgumentNullException.ThrowIfNull(importFilePath);

                // Check if import file exists based on mode
                bool importExists = false;
                if (Mode == DocumentManagerMode.InMemory)
                {
                    importExists = File.Exists(importFilePath);
                }
                else
                {
                    importExists = StorageManager!.HasDocument(importFilePath);
                }

                if (!importExists)
                    return AgentToolResult.Fail($"Annotation import file not found: {importFilePath}");

                // -- Open --------------------------------------------------------
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                bool isReloaded = false;
                PdfLoadedDocument loadedDocument;

                // ✅ Ensure PdfLoadedDocument
                if (document is PdfLoadedDocument pdfLoaded)
                {
                    loadedDocument = pdfLoaded;
                }
                else
                {
                    var reloadStream = new MemoryStream();
                    document.Save(reloadStream);
                    reloadStream.Position = 0;

                    if (Mode == DocumentManagerMode.InMemory)
                        InMemoryManager!.RemoveDocument(documentIdOrFilePath);
                    loadedDocument = new PdfLoadedDocument(reloadStream);

                    isReloaded = true;
                }

                // Load import data based on mode
                Stream? importStream = null;
                try
                {
                    if (Mode == DocumentManagerMode.InMemory)
                    {
                        importStream = new FileStream(importFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    }
                    else
                    {
                        importStream = StorageManager!.GetDocumentStream(importFilePath);
                    }

                    if (importStream == null)
                        return AgentToolResult.Fail($"Failed to read annotation import file: {importFilePath}");

                    loadedDocument.ImportAnnotations(importStream, format);
                }
                finally
                {
                    importStream?.Dispose();
                }

                // -- Save --------------------------------------------------------
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_annotations_imported.pdf";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, loadedDocument);
                if (Mode == DocumentManagerMode.InMemory)
                {
                    if (isReloaded)
                        outputKey = ((PdfDocumentManager)InMemoryManager!).ImportDocumentInstance(loadedDocument);
                    else
                        outputKey = documentIdOrFilePath;
                }

                return AgentToolResult.Ok(
                    $"Annotations imported successfully into document {outputKey}.",
                    new { DocumentId = outputKey, Format = format.ToString(), SourceFile = importFilePath });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to import annotations: {ex.Message}");
            }
        }

        /// <summary>
        /// Exports form field data from a PDF document into FDF, XFDF, or XML format.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="format">Export format (FDF, XFDF, XML).</param>
        /// <param name="exportPath">Optional export file or folder path.</param>
        /// <returns>Result containing the exported form field data or the export file path.</returns>
        [Tool(
            Name = "ExportFormFields",
            Description = "Exports form field data from a PDF document into FDF, XFDF, or XML format. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode)."
        )]
        public AgentToolResult ExportFormFields(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
        string documentIdOrFilePath,

            [ToolParameter(Description = "Export format (FDF, XFDF, XML)")]
        DataFormat format,

            [ToolParameter(Description = "Optional: Export file or folder path")]
        string? exportPath = null
        )
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);

                
                

                // -- Open --------------------------------------------------------
                var document = OpenDocument(documentIdOrFilePath);

                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                bool isReloaded = false;
                PdfLoadedDocument loadedDocument;

                // ✅ Ensure PdfLoadedDocument
                if (document is PdfLoadedDocument pdfLoaded)
                {
                    loadedDocument = pdfLoaded;
                }
                else
                {
                    var reloadStream = new MemoryStream();
                    document.Save(reloadStream);
                    reloadStream.Position = 0;

                    if (Mode == DocumentManagerMode.InMemory)
                        InMemoryManager!.RemoveDocument(documentIdOrFilePath);
                    loadedDocument = new PdfLoadedDocument(reloadStream);

                    isReloaded = true;
                }

                PdfLoadedForm loadedForm = loadedDocument.Form;
                if (loadedForm == null)
                    return AgentToolResult.Fail("No form fields found in the PDF document.");

                string sourceFileName = !string.IsNullOrWhiteSpace(exportPath)
                    ? Path.GetFileName(exportPath)
                    : "Source.pdf";

                using var memoryStream = new MemoryStream();
                loadedForm.ExportData(memoryStream, format, sourceFileName);
                memoryStream.Position = 0;

                string resolvedDocId = documentIdOrFilePath;
                if (isReloaded && Mode == DocumentManagerMode.InMemory)
                    resolvedDocId = ((PdfDocumentManager)InMemoryManager!).ImportDocumentInstance(loadedDocument);

                if (!string.IsNullOrWhiteSpace(exportPath))
                {
                    string finalFilePath = ResolveFormExportPath(exportPath, documentIdOrFilePath, format);
                    if(Mode == DocumentManagerMode.InMemory)
                    {
                        using var fileStream = new FileStream(finalFilePath, FileMode.Create, FileAccess.Write, FileShare.None);
                        memoryStream.CopyTo(fileStream);
                    }
                    else
                    {
                        using (var formStream = new MemoryStream())
                        {
                            memoryStream.Position = 0;
                            memoryStream.CopyTo(formStream);
                            memoryStream.Position = 0;
                            formStream.Position = 0;
                            SaveFile(finalFilePath, formStream);
                        }
                    }
                    return AgentToolResult.Ok(
                        "Form fields exported successfully.",
                        new { DocumentId = resolvedDocId, Format = format.ToString(), ExportPath = finalFilePath });
                }

                string formFieldData = Encoding.UTF8.GetString(memoryStream.ToArray());

                

                return AgentToolResult.Ok(
                    "Form fields exported successfully.",
                    new { DocumentId = resolvedDocId, Format = format.ToString(), FormFieldData = formFieldData });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to export form fields: {ex.Message}");
            }
        }

        /// <summary>
        /// Resolves export file path from folder or file input.
        /// </summary>
        private string ResolveFormExportPath(string exportPath, string documentId, DataFormat format)
        {
            string extension = format switch
            {
                DataFormat.Fdf => ".fdf",
                DataFormat.XFdf => ".xfdf",
                DataFormat.Xml => ".xml",
                _ => ".txt"
            };

            if (Mode == DocumentManagerMode.InMemory)
            {
                // Folder path or path without extension
                if (Directory.Exists(exportPath) || string.IsNullOrEmpty(Path.GetExtension(exportPath)))
                {
                    Directory.CreateDirectory(exportPath);
                    return Path.Combine(exportPath, $"FormFields_{documentId}{extension}");
                }

                string? directory = Path.GetDirectoryName(exportPath);
                if (!string.IsNullOrWhiteSpace(directory))
                    Directory.CreateDirectory(directory);
            }
            else
            {
                // For DocumentStorage mode, ensure the path has proper extension
                if (string.IsNullOrEmpty(Path.GetExtension(exportPath)))
                {
                    return Path.Combine(exportPath,$"FormFields_Export{extension}");
                }
            }

            return exportPath;
        }

        /// <summary>
        /// Imports form field data into a PDF document from an FDF, XFDF, or XML file.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="format">Import format (FDF, XFDF, XML).</param>
        /// <param name="sourcePdfPath">Form field data file path (FDF/XFDF/XML file).</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the document ID and import details.</returns>
        [Tool(
            Name = "ImportFormFields",
            Description = "Imports form field data into a PDF document from FDF, XFDF, or XML. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode)."
        )]
        public AgentToolResult ImportFormFields(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
        string documentIdOrFilePath,

            [ToolParameter(Description = "Import format (FDF, XFDF, XML)")]
        DataFormat format,

            [ToolParameter(Description = "Form field data file path (FDF/XFDF/XML file)")]
        string? sourcePdfPath = null,

            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
        string? outputFilePath = null
        )
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);

                if (string.IsNullOrWhiteSpace(sourcePdfPath))
                    return AgentToolResult.Fail("A valid form field data file path must be provided.");

                // Check if source file exists based on mode
                bool sourceExists = false;
                if (Mode == DocumentManagerMode.InMemory)
                {
                    sourceExists = File.Exists(sourcePdfPath);
                }
                else
                {
                    sourceExists = StorageManager!.HasDocument(sourcePdfPath);
                }

                if (!sourceExists)
                    return AgentToolResult.Fail($"Form field data file not found: {sourcePdfPath}");

                // -- Open --------------------------------------------------------
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                bool isReloaded = false;
                PdfLoadedDocument loadedDocument;

                // ✅ Ensure PdfLoadedDocument
                if (document is PdfLoadedDocument pdfLoaded)
                {
                    loadedDocument = pdfLoaded;
                }
                else
                {
                    var reloadStream = new MemoryStream();
                    document.Save(reloadStream);
                    reloadStream.Position = 0;

                    if (Mode == DocumentManagerMode.InMemory)
                        InMemoryManager!.RemoveDocument(documentIdOrFilePath);
                    loadedDocument = new PdfLoadedDocument(reloadStream);

                    isReloaded = true;
                }

                PdfLoadedForm loadedForm = loadedDocument.Form;
                if (loadedForm == null)
                    return AgentToolResult.Fail("No form fields found in the PDF document.");

                if (format != DataFormat.Fdf && format != DataFormat.XFdf && format != DataFormat.Xml && format != DataFormat.Json)
                    return AgentToolResult.Fail($"Unsupported import format: {format}");

                // Load import data based on mode
                Stream? importStream = null;
                try
                {
                    if (Mode == DocumentManagerMode.InMemory)
                    {
                        if (File.Exists(sourcePdfPath))
                        {
                            importStream = new FileStream(sourcePdfPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                            
                        }
                        
                    }
                    else
                    {
                        if (StorageManager!.HasDocument(sourcePdfPath))
                        {
                            importStream = StorageManager!.GetDocumentStream(sourcePdfPath);
                        }
                        
                    }

                    if (importStream == null)
                        return AgentToolResult.Fail($"Failed to read form field data file: {sourcePdfPath}");
                    ImportFormSettings settings = new ImportFormSettings();
                    settings.DataFormat = format;
                    loadedForm.ImportData(importStream, settings);

                    
                }
                finally
                {
                    importStream?.Dispose();
                }

                // -- Save --------------------------------------------------------
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_formfields_imported.pdf";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, loadedDocument);
                if (Mode == DocumentManagerMode.InMemory)
                {
                    if (isReloaded)
                        outputKey = ((PdfDocumentManager)InMemoryManager!).ImportDocumentInstance(loadedDocument);
                    else
                        outputKey = documentIdOrFilePath;
                }

                return AgentToolResult.Ok(
                    $"Form fields imported successfully into document {outputKey}.",
                    new { DocumentId = outputKey, Format = format.ToString(), SourceFile = sourcePdfPath });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to import form fields: {ex.Message}");
            }
        }

    }
}
