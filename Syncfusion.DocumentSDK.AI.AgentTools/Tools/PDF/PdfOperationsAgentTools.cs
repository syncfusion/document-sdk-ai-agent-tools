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
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;

namespace Syncfusion.AI.AgentTools.PDF
{
    /// <summary>
    /// Provides AI agent tools for PDF document manipulation operations.
    /// Handles merging, splitting, and compression of PDF documents.
    /// </summary>
    public class PdfOperationsAgentTools : AgentToolBase<PdfDocumentBase>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PdfOperationsAgentTools"/> class (Mode 1 � InMemory).
        /// </summary>
        /// <param name="manager">The PDF document manager.</param>
        public PdfOperationsAgentTools(PdfDocumentManager manager)
            : base(manager, DocumentType.PDF) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="PdfOperationsAgentTools"/> class (Mode 2 � DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public PdfOperationsAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.PDF) { }

        /// <summary>
        /// Merges multiple PDF files into a single PDF document and returns the new document ID.
        /// Each input may include an optional password for encrypted PDFs.
        /// </summary>
        /// <param name="pdfFiles">List of PDF files to merge. Each item must contain a FilePath and an optional Password.</param>
        /// <param name="mergeAccessibilityTags">Whether to merge accessibility tags from source PDFs.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the merged document ID and file count.</returns>
        [Tool(
            Name = "MergePdfs",
            Description =
                "Merges multiple PDF files into a single PDF document and returns the document ID. " +
                "FilePath accepts a local file path (InMemory mode) or a storage key (DocumentStorage mode). " +
                "outputFilePath: Output file path for saving the result (DocumentStorage mode only). " +
                "Input format example: " +
                "{\"pdfFiles\":[{\"FilePath\":\"invoice.pdf\",\"Password\":null},{\"FilePath\":\"image-pdf.pdf\",\"Password\":null}],\"mergeAccessibilityTags\":true}"
        )]
        public AgentToolResult MergePdfs(
            [ToolParameter(
                Description =
                    "List of PDF files to merge. Each item must contain a FilePath (local file path in InMemory mode, or storage key in DocumentStorage mode) and an optional Password. " +
                    "Example: [{ \"FilePath\": \"invoice.pdf\", \"Password\": null }]"
            )]
            List<PdfFileInput> pdfFiles,

            [ToolParameter(Description = "Whether to merge accessibility tags from source PDFs")]
            bool mergeAccessibilityTags = false,

            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null
        )
        {
            if (pdfFiles == null || pdfFiles.Count == 0)
                return AgentToolResult.Fail("At least one PDF file must be provided.");

            PdfDocument mergedDocument = new PdfDocument
            {
                EnableMemoryOptimization = true
            };

            PdfMergeOptions mergeOptions = new PdfMergeOptions
            {
                MergeAccessibilityTags = mergeAccessibilityTags
            };

            List<PdfLoadedDocument> loadedDocuments = new();
            List<Stream> openedStreams = new();

            try
            {
                foreach (var input in pdfFiles)
                {
                    if (input == null)
                        return AgentToolResult.Fail("PDF file input cannot be null.");

                    if (string.IsNullOrWhiteSpace(input.FilePath))
                        return AgentToolResult.Fail("PDF file path cannot be null or empty.");

                    // ── Resolve stream based on mode ─────────────────────────────
                    Stream pdfStream;

                    if (Mode == DocumentManagerMode.InMemory)
                    {
                        if (!Path.IsPathRooted(input.FilePath))
                            return AgentToolResult.Fail(
                                $"PDF file path must be absolute: {input.FilePath}");

                        if (!File.Exists(input.FilePath))
                            return AgentToolResult.Fail(
                                $"PDF file not found: {input.FilePath}");

                        pdfStream = new FileStream(
                            input.FilePath,
                            FileMode.Open,
                            FileAccess.Read,
                            FileShare.Read);
                    }
                    else
                    {
                        // Mode 2: try storage first, then fall back to local file system
                        if (StorageManager!.HasDocument(input.FilePath))
                        {
                            Stream? storageStream = StorageManager.GetDocumentStream(input.FilePath);
                            if (storageStream == null)
                                return AgentToolResult.Fail(
                                    $"PDF not found in storage: {input.FilePath}");
                            pdfStream = storageStream;
                        }
                        else if (File.Exists(input.FilePath))
                        {
                            pdfStream = new FileStream(
                                input.FilePath,
                                FileMode.Open,
                                FileAccess.Read,
                                FileShare.Read);
                        }
                        else
                        {
                            return AgentToolResult.Fail(
                                $"PDF file not found: {input.FilePath}");
                        }
                    }

                    openedStreams.Add(pdfStream);

                    PdfLoadedDocument loadedDocument;

                    try
                    {
                        loadedDocument = string.IsNullOrWhiteSpace(input.Password)
                            ? new PdfLoadedDocument(pdfStream, true)
                            : new PdfLoadedDocument(pdfStream, input.Password, true);
                    }
                    catch (Exception ex)
                    {
                        return AgentToolResult.Fail(
                            $"Failed to open PDF '{input.FilePath}'. Possible invalid password or corrupt file. Error: {ex.Message}");
                    }

                    loadedDocument.EnableMemoryOptimization = true;
                    loadedDocuments.Add(loadedDocument);
                }

                PdfDocument.Merge(mergedDocument, mergeOptions, loadedDocuments.ToArray());

                MemoryStream memory = new MemoryStream();
                mergedDocument.Save(memory);
                mergedDocument.Close(true);

                PdfLoadedDocument mergedLoaded = new PdfLoadedDocument(memory);

                // ── Save ─────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_merged.pdf";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, mergedLoaded);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = ((PdfDocumentManager)InMemoryManager!).ImportDocumentInstance(mergedLoaded);

                return AgentToolResult.Ok(
                    $"Successfully merged {pdfFiles.Count} PDF documents into document {outputKey}.",
                    new
                    {
                        DocumentId = outputKey,
                        MergedFileCount = pdfFiles.Count,
                        AccessibilityTagsMerged = mergeAccessibilityTags
                    });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail(
                    $"Failed to merge PDF documents: {ex.Message}");
            }
            finally
            {
                // ✅ Close all loaded documents
                foreach (PdfLoadedDocument doc in loadedDocuments)
                {
                    doc.Close(true);
                }

                foreach (Stream stream in openedStreams)
                {
                    stream.Dispose();
                }
            }
        }

   
        /// <summary>
        /// Splits a PDF document into multiple files by individual pages or by specified page ranges.
        /// If <paramref name="pageRanges"/> is null, the document is split into individual pages.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="pageRanges">Optional page ranges [[start, end]] (zero-based). Null splits all pages individually.</param>
        /// <param name="outputFilePattern">Output file name pattern (default: Output{0}.pdf).</param>
        /// <param name="outputFolderPath">Output folder path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the document ID, split mode, and output folder path.</returns>
        [Tool(
            Name = "SplitPdf",
            Description =
                "Splits a loaded PDF document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). " +
                "If pageRanges is null, the document is split into individual pages. " +
                "If pageRanges is provided, each range must be [startPage, endPage] (zero-based)."
        )]
        public AgentToolResult SplitPdf(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
    string documentIdOrFilePath,

            [ToolParameter(Description = "Optional page ranges [[start,end]]. Null = split all pages")]
    int[][]? pageRanges = null,

            [ToolParameter(Description = "Output file name pattern (default: Output{0}.pdf)")]
    string outputFilePattern = "Output{0}.pdf",

            [ToolParameter(Description = "Output folder path for saving the result (DocumentStorage mode only).")]
    string? outputFolderPath = null
        )
        {
            try
            {
                if (string.IsNullOrWhiteSpace(documentIdOrFilePath))
                    return AgentToolResult.Fail("documentIdOrFilePath is required.");

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

                string outputFolder = $"Split_{DateTime.Now:yyyyMMddHHmmss}";
                if (!string.IsNullOrEmpty(outputFolderPath) && Mode == DocumentManagerMode.DocumentStorage)
                {
                    outputFolder = outputFolderPath;
                }
                if (Mode == DocumentManagerMode.InMemory)
                {
                    outputFolder = Path.Combine(Environment.CurrentDirectory, $"Split_{DateTime.Now:yyyyMMddHHmmss}");
                    Directory.CreateDirectory(outputFolder);
                }

                string outputPattern = Path.Combine(outputFolder, outputFilePattern);

                // Note: SplitPdf does not add an outputFilePath param as it writes multiple files
                // to disk � the outputFolder acts as the destination for split pages.
                // The document ID is updated for InMemory mode if reloaded.
                string resolvedDocId = documentIdOrFilePath;
                if (isReloaded && Mode == DocumentManagerMode.InMemory)
                    resolvedDocId = ((PdfDocumentManager)InMemoryManager!).ImportDocumentInstance(loadedDocument);

                if (pageRanges == null)
                {

                    if (Mode == DocumentManagerMode.DocumentStorage)
                    {
                        int counter = 1;
                        loadedDocument.DocumentSplitEvent += (sender, args) =>
                        {
                            string splitFileName = string.Format(outputPattern, counter++);
                            SaveFile(splitFileName, args.PdfDocumentData);
                        };
                        loadedDocument.SplitByFixedNumber(1);
                    }
                    else
                    {
                        loadedDocument.Split(outputPattern);
                    }

                    return AgentToolResult.Ok(
                        "PDF split into individual pages successfully.",
                        new { DocumentId = resolvedDocId, Mode = "IndividualPages", OutputFolder = outputFolder }
                    );
                }

                // ✅ Validate page ranges
                int[,] ranges = new int[pageRanges.Length, 2];

                for (int i = 0; i < pageRanges.Length; i++)
                {
                    if (pageRanges[i] == null || pageRanges[i].Length != 2)
                        return AgentToolResult.Fail(
                            "Each page range must contain exactly two values [startPage, endPage].");

                    ranges[i, 0] = pageRanges[i][0];
                    ranges[i, 1] = pageRanges[i][1];
                }
                if (Mode == DocumentManagerMode.DocumentStorage)
                {
                    int counter = 1;
                    loadedDocument.DocumentSplitEvent += (sender, args) =>
                    {
                        string splitFileName = string.Format(outputPattern, counter++);
                        SaveFile(splitFileName, args.PdfDocumentData);
                    };
                    loadedDocument.SplitByRanges(ranges);
                }
                else
                {
                    // 2️⃣ Split by ranges
                    loadedDocument.SplitByRanges(outputPattern, ranges);
                }

                return AgentToolResult.Ok(
                    "PDF split by page ranges successfully.",
                    new { DocumentId = resolvedDocId, Mode = "Ranges", OutputFolder = outputFolder, PageRanges = pageRanges }
                );
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to split PDF document: {ex.Message}");
            }
        }
        /// <summary>
        /// Compresses a PDF document by optimizing images, fonts, page contents,
        /// and removing metadata using either documentId or filePath.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="compressImages">Whether to compress images.</param>
        /// <param name="imageQuality">Image quality (10–100).</param>
        /// <param name="optimizeFont">Whether to optimize embedded fonts.</param>
        /// <param name="optimizePageContents">Whether to optimize page contents.</param>
        /// <param name="removeMetadata">Whether to remove document metadata.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the document ID and compression options applied.</returns>
        [Tool(
        Name = "CompressPdf",
        Description = "Compresses an existing PDF document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Uses image compression, font optimization, page content optimization, and metadata removal."
        )]
        public AgentToolResult CompressPdf(
        [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
        string documentIdOrFilePath,
        [ToolParameter(Description = "Whether to compress images")]
        bool compressImages = true,
        [ToolParameter(Description = "Image quality (10�100)")]
        int imageQuality = 50,
        [ToolParameter(Description = "Whether to optimize embedded fonts")]
        bool optimizeFont = true,
        [ToolParameter(Description = "Whether to optimize page contents")]
        bool optimizePageContents = true,
        [ToolParameter(Description = "Whether to remove document metadata")]
        bool removeMetadata = true,
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

                bool isNewDocument = false;
                if (document is PdfDocument)
                {
                    MemoryStream stream = new MemoryStream();
                    document.Save(stream);
                    if (Mode == DocumentManagerMode.InMemory)
                        InMemoryManager!.RemoveDocument(documentIdOrFilePath);
                    isNewDocument = true;
                    document = new PdfLoadedDocument(stream);
                }
                PdfLoadedDocument loadedDocument = document as PdfLoadedDocument;
                //return AgentToolResult.Fail("Compression is supported only for loaded PDF documents");
 
                PdfCompressionOptions options = new PdfCompressionOptions
                {
                    CompressImages = compressImages,
                    ImageQuality = imageQuality,
                    OptimizeFont = optimizeFont,
                    OptimizePageContents = optimizePageContents,
                    RemoveMetadata = removeMetadata
                };

                loadedDocument.Compress(options);

                // -- Save --------------------------------------------------------
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_compressed.pdf";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, loadedDocument);
                if (Mode == DocumentManagerMode.InMemory)
                {
                    if (isNewDocument)
                        outputKey = ((PdfDocumentManager)InMemoryManager!).ImportDocumentInstance(loadedDocument);
                    else
                        outputKey = documentIdOrFilePath;
                }

                return AgentToolResult.Ok(
                    $"PDF compressed successfully into document {outputKey}",
                    new
                    {
                        DocumentId = outputKey,
                        CompressImages = compressImages,
                        ImageQuality = imageQuality,
                        OptimizeFont = optimizeFont,
                        OptimizePageContents = optimizePageContents,
                        RemoveMetadata = removeMetadata
                    });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to compress PDF document: {ex.Message}");
            }
        }

        /// <summary>
        /// Reorders the pages of an existing PDF document using a specified page index sequence.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="orderIndexes">Zero-based page indexes defining the new order. Length must equal the page count.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the document ID, page count, and new page order.</returns>
        [Tool(
            Name = "ReorderPdfPages",
            Description = "Rearranges PDF pages using a zero-based page index sequence. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Get the PDF page count first and ensure the index array length matches it."
        )]
        public AgentToolResult ReorderPdfPages(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
        string documentIdOrFilePath,

            [ToolParameter(
        Description = "Zero-based page indexes defining the new order. Get the PDF page count first; the array length must equal the page count. Values must be in range and not repeated."
        )]
        int[] orderIndexes,

            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
        string? outputFilePath = null
        )
        {
            try
            {
                if (string.IsNullOrWhiteSpace(documentIdOrFilePath))
                    return AgentToolResult.Fail("documentIdOrFilePath cannot be null or empty.");

                if (orderIndexes == null || orderIndexes.Length == 0)
                    return AgentToolResult.Fail("orderIndexes cannot be null or empty.");

                // -- Open --------------------------------------------------------
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                bool isNewDocument = false;
                int originalPageCount = document.PageCount;

                // ✅ Validation
                if (orderIndexes.Length != originalPageCount)
                    return AgentToolResult.Fail(
                        $"orderIndexes length must match page count ({originalPageCount}).");

                if (orderIndexes.Any(i => i < 0 || i >= originalPageCount))
                    return AgentToolResult.Fail(
                        "orderIndexes contains invalid page index values.");

                if (document is PdfDocument)
                {
                    MemoryStream stream = new MemoryStream();
                    document.Save(stream);
                    if (Mode == DocumentManagerMode.InMemory)
                        InMemoryManager!.RemoveDocument(documentIdOrFilePath);
                    stream.Position = 0;
                    document = new PdfLoadedDocument(stream);
                    isNewDocument = true;
                }

                PdfLoadedDocument loadedDocument = document as PdfLoadedDocument;
                loadedDocument.Pages.ReArrange(orderIndexes);

                // -- Save --------------------------------------------------------
                
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_reordered.pdf";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, loadedDocument);
                if (Mode == DocumentManagerMode.InMemory)
                {
                    if (isNewDocument)
                        outputKey = ((PdfDocumentManager)InMemoryManager!).ImportDocumentInstance(loadedDocument);
                    else
                        outputKey = documentIdOrFilePath;
                }

                return AgentToolResult.Ok(
                    $"PDF document pages reordered successfully into document {outputKey}.",
                    new { DocumentId = outputKey, PageCount = originalPageCount, NewOrder = orderIndexes }
                );
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail(
                    $"Failed to reorder PDF pages: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Represents a PDF file input with an optional password.
    /// </summary>
    public class PdfFileInput
    {
        /// <summary>
        /// Absolute file path of the PDF.
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// Password for encrypted PDFs (null or empty if not encrypted).
        /// </summary>
        public string? Password { get; set; }
    }
}
