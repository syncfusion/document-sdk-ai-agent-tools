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
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using System;
using System.IO;
using System.Linq;
using Syncfusion.Drawing;


namespace Syncfusion.AI.AgentTools.PDF
{
    /// <summary>
    /// Provides AI agent tools for PDF document convertion operations.
    /// </summary>
    public class PdfConverterAgentTools : AgentToolBase<PdfDocumentBase>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PdfConverterAgentTools"/> class (Mode 1 � InMemory).
        /// </summary>
        /// <param name="manager">The PDF document manager.</param>
        public PdfConverterAgentTools(PdfDocumentManager manager)
            : base(manager, DocumentType.PDF) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="PdfConverterAgentTools"/> class (Mode 2 � DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public PdfConverterAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.PDF) { }

        /// <summary>
        /// Helper method to handle missing fonts during PDF conversion.
        /// </summary>
        private void LoadedDocument_SubstituteFont(object sender, PdfFontEventArgs args)
        {
            // Extract the font name (ignoring style suffixes)
            string fontName = args.FontName.Split(',')[0];
            // Determine the font style
            PdfFontStyle fontStyle = args.FontStyle;
            SKFontStyle skFontStyle = SKFontStyle.Normal;
            // Map PDF font styles to SkiaSharp font styles
            if (fontStyle == PdfFontStyle.Bold)
                skFontStyle = SKFontStyle.Bold;
            else if (fontStyle == PdfFontStyle.Italic)
                skFontStyle = SKFontStyle.Italic;
            else if (fontStyle == (PdfFontStyle.Bold | PdfFontStyle.Italic))
                skFontStyle = SKFontStyle.BoldItalic;
            // Load the typeface using SkiaSharp
            SKTypeface typeface = SKTypeface.FromFamilyName(fontName, skFontStyle);
            SKStreamAsset typeFaceStream = typeface.OpenStream();

            // Create a memory stream from the font data
            MemoryStream memoryStream = null;
            if (typeFaceStream != null && typeFaceStream.Length > 0)
            {
                byte[] fontData = new byte[typeFaceStream.Length];
                typeFaceStream.Read(fontData, fontData.Length);
                typeFaceStream.Dispose();
                memoryStream = new MemoryStream(fontData);
            }
            // Assign the font stream to the event arguments
            args.FontStream = memoryStream;
        }

        /// <summary>
        /// Converts an existing PDF document into a PDF/A-compliant document
        /// based on the specified conformance level.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="conformanceLevel">Target PDF/A conformance level (PdfA1B, PdfA2B, PdfA3B, Pdf_A4, Pdf_A4F, Pdf_A4E).</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the document ID and the applied conformance level.</returns>
        [Tool(
            Name = "ConvertPdfToPdfA",
            Description = "Converts an existing PDF document to a PDF/A compliant format. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). " +
                          "conformanceLevel defines the supported PDF/A standard to apply ( PdfA1B, PdfA2B, PdfA3B, Pdf_A4, Pdf_A4F, Pdf_A4E)."
        )]
        public AgentToolResult ConvertPdfToPdfA(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath,
            [ToolParameter(Description = "Target PDF/A conformance level")] PdfConformanceLevel conformanceLevel,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
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

                // Register font substitution handler to handle missing fonts during conformance conversion
                loadedDocument.SubstituteFont += LoadedDocument_SubstituteFont;
                try
                {
                    // Convert the existing PDF to PDF/A (modifies in-place)
                    loadedDocument.ConvertToPDFA(conformanceLevel);
                }
                finally
                {
                    // Unregister the font substitution handler
                    loadedDocument.SubstituteFont -= LoadedDocument_SubstituteFont;
                }

                // -- Save --------------------------------------------------------
                 if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_pdfa.pdf";
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
                    $"PDF document successfully converted to {conformanceLevel} into document {outputKey}.",
                    new { DocumentId = outputKey, ConformanceLevel = conformanceLevel.ToString() });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to convert PDF to PDF/A: {ex.Message}");
            }
        }
        /// <summary>
        /// Creates a PDF document from one or more image files using ImageToPdfConverter
        /// with control over image placement and page size.
        /// </summary>
        /// <param name="imageFiles">Array of image file paths (InMemory mode) or storage keys (DocumentStorage mode).</param>
        /// <param name="imagePosition">Image placement style on the PDF page.</param>
        /// <param name="pageWidth">The width of the PDF page in pixels.</param>
        /// <param name="pageHeight">The height of the PDF page in pixels.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the document ID of the created PDF.</returns>
        [Tool(
            Name = "ImageToPdf",
            Description = "Creates a PDF document from one or more image files using ImageToPdfConverter with control over image placement and page size. " +
                          "imageFiles: Array of image file paths (InMemory mode) or storage keys (DocumentStorage mode). " +
                          "outputFilePath: Output file path for saving the result (DocumentStorage mode only)."
        )]
        public AgentToolResult ImageToPdf(
            [ToolParameter(Description = "Array of image file paths (InMemory mode) or storage keys (DocumentStorage mode)")]
            string[] imageFiles,

            [ToolParameter(Description = "Image placement style on the PDF page")]
            PdfImagePosition imagePosition = PdfImagePosition.FitToPage,

            [ToolParameter(Description = "The width of the PDF page in pixels")]
            int pageWidth = 612,

            [ToolParameter(Description = "The height of the PDF page in pixels")]
            int pageHeight = 792,

            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null
        )
        {
            if (imageFiles == null || imageFiles.Length == 0)
                return AgentToolResult.Fail("No image files provided.");

            List<Stream> imageStreams = new List<Stream>();

            try
            {
                // ── Resolve image streams based on mode ─────────────────────────
                foreach (string imageFile in imageFiles)
                {
                    if (Mode == DocumentManagerMode.InMemory)
                    {
                        if (!File.Exists(imageFile))
                            return AgentToolResult.Fail($"Image file not found: {imageFile}");

                        imageStreams.Add(new FileStream(imageFile, FileMode.Open, FileAccess.Read));
                    }
                    else
                    {
                        // Mode 2: try storage first, then fall back to local file system
                        if (StorageManager!.HasDocument(imageFile))
                        {
                            Stream? storageStream = StorageManager.GetDocumentStream(imageFile);
                            if (storageStream == null)
                                return AgentToolResult.Fail($"Image not found in storage: {imageFile}");

                            imageStreams.Add(storageStream);
                        }
                        else if (File.Exists(imageFile))
                        {
                            imageStreams.Add(new FileStream(imageFile, FileMode.Open, FileAccess.Read));
                        }
                        else
                        {
                            return AgentToolResult.Fail($"Image file not found: {imageFile}");
                        }
                    }
                }

                // ── Convert images to PDF ────────────────────────────────────────
                ImageToPdfConverter converter = new ImageToPdfConverter
                {
                    ImagePosition = imagePosition,
                    PageSize = new Syncfusion.Drawing.SizeF(pageWidth, pageHeight)
                };

                PdfDocument pdfDocument = converter.Convert(imageStreams.ToArray());

                // ── Save ─────────────────────────────────────────────────────────
                 if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                outputFilePath = "output_image_to_pdf.pdf";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, pdfDocument);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = ((PdfDocumentManager)InMemoryManager!).ImportDocumentInstance(pdfDocument);

                return AgentToolResult.Ok(
                    $"Images successfully converted to PDF into document {outputKey}.",
                    new
                    {
                        DocumentId = outputKey,
                        ImageCount = imageFiles.Length,
                        PageWidth = pageWidth,
                        PageHeight = pageHeight,
                        ImagePosition = imagePosition.ToString()
                    }
                );
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to convert images to PDF: {ex.Message}");
            }
            finally
            {
                // ── Ensure all streams are disposed ──────────────────────────────
                foreach (Stream stream in imageStreams)
                {
                    stream.Dispose();
                }
            }
        }
    }
}
