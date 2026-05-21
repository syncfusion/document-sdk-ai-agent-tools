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
using System.Text;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Drawing;
using Syncfusion.Pdf.Exporting;

namespace Syncfusion.AI.AgentTools.PDF
{
    /// <summary>
    /// Provides AI agent tools for extracting content from PDF documents.
    /// Handles text, image, and table extraction operations.
    /// </summary>
    public class PdfContentExtractionAgentTools : AgentToolBase<PdfDocumentBase>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PdfContentExtractionAgentTools"/> class (Mode 1 � InMemory).
        /// </summary>
        /// <param name="manager">The PDF document manager.</param>
        public PdfContentExtractionAgentTools(PdfDocumentManager manager)
            : base(manager, DocumentType.PDF) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="PdfContentExtractionAgentTools"/> class (Mode 2 � DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public PdfContentExtractionAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.PDF) { }

        /// <summary>
        /// Extracts text content from a PDF document, either from a specific range of pages or from the entire file.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="startPageIndex">Starting page index (0-based), -1 for all pages.</param>
        /// <param name="endPageIndex">Ending page index (0-based), -1 for all pages.</param>
        /// <returns>Result containing the extracted text and page count.</returns>
        [Tool(Name = "ExtractText", Description = "Extracts text content from a PDF document, either from a specific range of pages or from the entire file. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ExtractText(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath,
            [ToolParameter(Description = "Starting page index (0-based), -1 for all pages")] int startPageIndex = -1,
            [ToolParameter(Description = "Ending page index (0-based), -1 for all pages")] int endPageIndex = -1)
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

                StringBuilder extractedText = new StringBuilder();

                // Determine page range
                int startPage = startPageIndex == -1 ? 0 : startPageIndex;
                int endPage = endPageIndex == -1 ? loadedDocument.PageCount - 1 : endPageIndex;

                // Validate page range
                if (startPage < 0 || startPage >= loadedDocument.PageCount)
                    return AgentToolResult.Fail($"Invalid start page index: {startPageIndex}. Document has {loadedDocument.PageCount} pages.");

                if (endPage < startPage || endPage >= loadedDocument.PageCount)
                    return AgentToolResult.Fail($"Invalid end page index: {endPageIndex}");

                // Extract text from each page
                for (int i = startPage; i <= endPage; i++)
                {
                    PdfPageBase page = loadedDocument.Pages[i];
                    string pageText = page.ExtractText();
                    extractedText.AppendLine($"--- Page {i + 1} ---");
                    extractedText.AppendLine(pageText);
                    extractedText.AppendLine();
                }

                string resolvedDocId = documentIdOrFilePath;
                if (isReloaded && Mode == DocumentManagerMode.InMemory)
                    resolvedDocId = ((PdfDocumentManager)InMemoryManager!).ImportDocumentInstance(loadedDocument);

                return AgentToolResult.Ok(
                    $"Successfully extracted text from {endPage - startPage + 1} pages",
                    new { DocumentId = resolvedDocId, Text = extractedText.ToString(), PageCount = endPage - startPage + 1 });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to extract text from PDF: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts images from a PDF document using a page range or the entire document.
        /// Saves the extracted images to the specified output folder.
        /// </summary>
        /// <remarks>
        /// Supports:
        /// <list type="bullet">
        /// <item>
        /// Full document extraction (default behavior).
        /// </item>
        /// <item>
        /// Page range extraction using <paramref name="startPageIndex"/> and <paramref name="endPageIndex"/>.
        /// </item>
        /// <item>
        /// Single-page extraction by passing the same value for both 
        /// <paramref name="startPageIndex"/> and <paramref name="endPageIndex"/>.
        /// </item>
        /// </list>
        /// Page indexes are zero-based.
        /// </remarks>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="startPageIndex">Starting page index (0-based). Use -1 to extract all pages.</param>
        /// <param name="endPageIndex">Ending page index (0-based). Use -1 to extract all pages.</param>
        /// <param name="outputPath">Optional output folder path.</param>
        /// <returns>Result containing the image count and output folder path.</returns>
        [Tool(
            Name = "ExtractImages",
            Description = "Extracts images from a PDF document using a page range or the entire document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode)."
        )]
        public AgentToolResult ExtractImages(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
        string documentIdOrFilePath,

            [ToolParameter(Description = "Starting page index (0-based). Use -1 to extract all pages.")]
        int startPageIndex = -1,

            [ToolParameter(Description = "Ending page index (0-based). Use -1 to extract all pages.")]
        int endPageIndex = -1,

            [ToolParameter(Description = "Optional output folder path.")]
        string? outputPath = null
        )
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);

                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                PdfLoadedDocument loadedDocument = document as PdfLoadedDocument;
                bool isReloaded = false;

                if (document is PdfDocument)
                {
                    MemoryStream stream = new MemoryStream();
                    document.Save(stream);
                    if (Mode == DocumentManagerMode.InMemory)
                        InMemoryManager!.RemoveDocument(documentIdOrFilePath);
                    loadedDocument = new PdfLoadedDocument(stream);
                    isReloaded = true;
                }

                // Validate page range (if provided)
                bool isPageRangeRequested = startPageIndex >= 0 && endPageIndex >= 0;

                if (isPageRangeRequested)
                {
                    if (startPageIndex > endPageIndex)
                        return AgentToolResult.Fail("Start page index cannot be greater than end page index.");

                    if (endPageIndex >= loadedDocument.PageCount)
                        return AgentToolResult.Fail("Page range exceeds document page count.");
                }

                using var extractor = new PdfDocumentExtractor();
                using var pdfStream = new MemoryStream();

                loadedDocument.Save(pdfStream);
                pdfStream.Position = 0;
                extractor.Load(pdfStream);

                string resolvedDocId = documentIdOrFilePath;
                if (isReloaded && Mode == DocumentManagerMode.InMemory)
                    resolvedDocId = ((PdfDocumentManager)InMemoryManager!).ImportDocumentInstance(loadedDocument);

                // Extract images
                Stream[] images = isPageRangeRequested

                    ? extractor.ExtractImages(startPageIndex, endPageIndex)
                    : extractor.ExtractImages();

                // Resolve output folder
                string finalOutputFolder = ResolveOutputFolder(outputPath);

                if (Mode == DocumentManagerMode.InMemory)
                {
                    Directory.CreateDirectory(finalOutputFolder);
                }

                // Save images
                for (int i = 0; i < images.Length; i++)
                {
                    string imagePath = Path.Combine(finalOutputFolder, $"Image_{i + 1}.png");

                    Stream imageStream = images[i];
                    if (imageStream.Length > 0 && imageStream.CanRead)
                    {
                        if (Mode == DocumentManagerMode.InMemory)
                        {
                            using var fileStream = new FileStream(imagePath, FileMode.Create, FileAccess.Write, FileShare.None);
                            imageStream.Position = 0;
                            imageStream.CopyTo(fileStream);
                            imageStream.Dispose();
                        }
                        else
                        {
                            using var memoryStream = new MemoryStream();
                            imageStream.Position = 0;
                            imageStream.CopyTo(memoryStream);
                            memoryStream.Position = 0;
                            SaveFile(imagePath, memoryStream);
                            imageStream.Dispose();
                        }
                    }
                }

                return AgentToolResult.Ok(
                    "Images extracted successfully.",
                    new
                    {
                        DocumentId = resolvedDocId,
                        ImageCount = images.Length,
                        OutputFolder = finalOutputFolder,
                        Mode = isPageRangeRequested ? "PageRange" : "AllPages"
                    });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to extract images from PDF: {ex.Message}");
            }
        }

        /// <summary>
        /// Searches a PDF document for the specified text and returns all occurrences
        /// grouped by page with their bounding rectangle positions.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// The unique identifier of the PDF document to search.
        /// </param>
        /// <param name="texts">
        /// The text to be searched.
        /// </param>
        /// <returns>
        /// A mapping of page index to a list of rectangle positions
        /// where the text was found.
        /// </returns>
        [Tool(
            Name = "FindTextInPdf",
            Description = "Searches a PDF document for matching array of text and returns all occurrences grouped by page. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode)."
        )]
        public AgentToolResult FindTextInPdf(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath,
            [ToolParameter(Description = "The array of text to be searched")] string[] texts)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);
                ArgumentNullException.ThrowIfNull(texts);

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

                // Find text using Syncfusion UG-supported API
                loadedDocument.FindText(texts.ToList(), out var matchRects);

                string resolvedDocId = documentIdOrFilePath;
                if (isReloaded && Mode == DocumentManagerMode.InMemory)
                    resolvedDocId = ((PdfDocumentManager)InMemoryManager!).ImportDocumentInstance(loadedDocument);

                return AgentToolResult.Ok(
                    $"Text search completed. Found matches on {matchRects.Count} pages.",
                    new { DocumentId = resolvedDocId, Matches = matchRects });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to find text in PDF: {ex.Message}");
            }
        }

        /// <summary>
        /// Retrieves the number of pages in a PDF document identified by the specified document ID or file path.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <returns>An AgentToolResult containing the page count of the specified PDF document if successful; otherwise, a
        /// failure result with an error message.</returns>
        [Tool(Name = "GetPdfDocumentPageCount", Description = "Returns the number of pages in the specified PDF document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult GetPdfDocumentPageCount(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);

                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                int pageCount = document.PageCount;

                return AgentToolResult.Ok(
                    $"PDF document {documentIdOrFilePath} has {pageCount} page(s)",
                    new { DocumentId = documentIdOrFilePath, PageCount = pageCount });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to get PDF document page count: {ex.Message}");
            }
        }
       
        [Tool(Name = "GetPdfDocumentPageSize", Description = "Gets the width and height of a specific page in a PDF document. Works for both newly created and loaded documents without reloading.This is required when placing elements like signatures, stamps, or annotations at specific positions such as bottom-right, bottom-left, top-right, top-left.")]
        public AgentToolResult GetPdfDocumentPageSize(
        [ToolParameter(Description = "The document ID (InMemory mode) or input file path.")]
        string documentIdOrFilePath,

        [ToolParameter(Description = "Page number (1-based index).")]
        int pageNumber)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);

                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                {
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");
                }

                if (pageNumber < 1)
                {
                    return AgentToolResult.Fail($"Invalid page number: {pageNumber}");
                }

                SizeF pageSize;

                switch (document)
                {
                    case PdfLoadedDocument loadedDoc:
                    {
                        int pageCount = loadedDoc.Pages.Count;

                        if (pageNumber > pageCount)
                        {
                            return AgentToolResult.Fail(
                                $"Invalid page number: {pageNumber}. Total pages: {pageCount}");
                        }

                        pageSize = loadedDoc.Pages[pageNumber - 1].Size;
                        break;
                    }

                    case PdfDocument pdfDoc:
                    {
                        int pageCount = pdfDoc.Pages.Count;

                        if (pageNumber > pageCount)
                        {
                            return AgentToolResult.Fail(
                                $"Invalid page number: {pageNumber}. Total pages: {pageCount}");
                        }

                        pageSize = pdfDoc.Pages[pageNumber - 1].GetClientSize();
                        break;
                    }

                    default:
                        return AgentToolResult.Fail("Unsupported PDF document type.");
                }

                return AgentToolResult.Ok(
                    $"Page {pageNumber} size retrieved successfully.",
                    new
                    {
                        PageNumber = pageNumber,
                        Width = pageSize.Width,
                        Height = pageSize.Height
                    });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to get page size: {ex.Message}");
            }
        }

        /// <summary>
        /// Resolves the output folder path based on the current mode.
        /// </summary>
        private string ResolveOutputFolder(string? outputPath)
        {
            if (Mode == DocumentManagerMode.InMemory)
            {
                return string.IsNullOrWhiteSpace(outputPath)
                    ? Path.Combine(Directory.GetCurrentDirectory(), $"ExtractedImages_{DateTime.Now:yyyyMMddHHmmss}")
                    : outputPath;
            }
            else
            {
                return string.IsNullOrWhiteSpace(outputPath)
                    ? $"ExtractedImages_{DateTime.Now:yyyyMMddHHmmss}"
                    : outputPath;
            }
        }

    }
}
