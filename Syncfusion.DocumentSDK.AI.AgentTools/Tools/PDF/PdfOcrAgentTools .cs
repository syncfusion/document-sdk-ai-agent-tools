#region Copyright Syncfusion
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
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;
using Syncfusion.OCRProcessor;

namespace Syncfusion.AI.AgentTools.PDF
{
    /// <summary>
    /// Provides AI agent tools for performing OCR operations on PDF documents.
    /// </summary>
    public class PdfOcrAgentTools : AgentToolBase<PdfDocumentBase>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PdfOcrAgentTools"/> class (Mode 1 � InMemory).
        /// </summary>
        /// <param name="manager">The PDF document manager.</param>
        public PdfOcrAgentTools(PdfDocumentManager manager)
            : base(manager, DocumentType.PDF) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="PdfOcrAgentTools"/> class (Mode 2 � DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public PdfOcrAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.PDF) { }

        /// <summary>
        /// Performs Optical Character Recognition (OCR) on a PDF document,
        /// making scanned or image-based text selectable and searchable.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="dataPath">Path to the Tesseract tessdata folder (required when language is not 'eng').</param>
        /// <param name="language">OCR language code (e.g., eng, fra, deu, tam). Default is eng.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the document ID, language used, and data path.</returns>
        [Tool(
        Name = "OcrPdf",
        Description = "Performs Optical Character Recognition (OCR) on a PDF document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode)."
        )]
        public AgentToolResult OcrPdf(
        [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
        string documentIdOrFilePath,

        [ToolParameter(Description = "Path to the Tesseract tessdata folder (required when language is not 'eng')")]
        string? dataPath = null,

        [ToolParameter(Description = "OCR language code (e.g., eng, fra, deu, tam). Default is eng")]
        string language = "eng",

        [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
        string? outputFilePath = null
)
        {
            if (string.IsNullOrWhiteSpace(documentIdOrFilePath))
                return AgentToolResult.Fail("Document ID or file path cannot be null or empty.");

            language = string.IsNullOrWhiteSpace(language)
                ? "eng"
                : language.Trim().ToLowerInvariant();

            if (language != "eng")
            {
                if (string.IsNullOrWhiteSpace(dataPath))
                    return AgentToolResult.Fail(
                        $"dataPath is required when language is '{language}'."
                    );

                if (!Directory.Exists(dataPath))
                    return AgentToolResult.Fail(
                        $"Invalid dataPath. Directory not found: {dataPath}"
                    );
            }

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

            using (var processor = new Syncfusion.OCRProcessor.OCRProcessor())
            {
                processor.Settings.Language = language;

                if (language == "eng")
                    processor.PerformOCR(loadedDocument);
                else
                    processor.PerformOCR(loadedDocument, dataPath!);
            }

            // -- Save --------------------------------------------------------
            if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                outputFilePath = "output_ocr.pdf";

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
                $"OCR successfully performed on document {outputKey}",
                new
                {
                    DocumentId = outputKey,
                    Language = language,
                    DataPathUsed = language == "eng" ? "Default" : dataPath
                }
            );
        }
    }
}
