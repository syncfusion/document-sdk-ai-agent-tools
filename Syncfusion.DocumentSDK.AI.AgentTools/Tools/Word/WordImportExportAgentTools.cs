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

using Syncfusion.AI.AgentTools.Core;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.IO;

namespace Syncfusion.AI.AgentTools.Word
{
    /// <summary>
    /// Provides agent tools for content format conversion and import/export operations.
    /// Handles in-memory operations for HTML, Markdown, and text conversions.
    /// </summary>
    public class WordImportExportAgentTools : AgentToolBase<WordDocument>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WordImportExportAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The document manager for managing Word documents.</param>
        public WordImportExportAgentTools(WordDocumentManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordImportExportAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public WordImportExportAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Imports HTML content into a Word document.
        /// </summary>
        /// <param name="htmlContentOrFilePath">The HTML content as a string or the file path to an HTML file.</param>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the destination document.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(
            Name = "ImportHtml",
            Description = "Imports/Merge HTML content into a Word document. htmlContent / filePath: The HTML content as a string or the file path to an HTML file. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ImportHtml(
            [ToolParameter(Description = "The HTML content as a string or the file path to an HTML file")]
            string htmlContentOrFilePath,
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the destination document")]
            string? documentIdOrFilePath = null,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                bool isTemporary = false;
                // ── Open ────────────────────────────────────────────────────────
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                {
                    document = new WordDocument();
                    isTemporary = true;
                }

                string htmlContent;

                if (Mode == DocumentManagerMode.InMemory && File.Exists(htmlContentOrFilePath))
                {
                    // Mode 1: file path fallback
                    htmlContent = File.ReadAllText(htmlContentOrFilePath);
                }
                else if (Mode == DocumentManagerMode.DocumentStorage && StorageManager!.HasDocument(htmlContentOrFilePath))
                {
                    // Mode 2: get document stream from storage and read as HTML
                    Stream? htmlDocStream = StorageManager!.GetDocumentStream(htmlContentOrFilePath);
                    if (htmlDocStream == null)
                        return AgentToolResult.Fail($"HTML Document not found: {htmlContentOrFilePath}");
                    using (var reader = new StreamReader(htmlDocStream, System.Text.Encoding.UTF8))
                    {
                        htmlContent = reader.ReadToEnd();
                    }
                    htmlDocStream.Dispose();
                }
                else
                {
                    htmlContent = htmlContentOrFilePath;
                }

                // Import HTML content
                if (document.LastParagraph == null)
                    document.EnsureMinimal();
                document.LastParagraph.AppendHTML(htmlContent);

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_html_imported.docx";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath ?? InMemoryManager!.ActiveDocumentId!; // InMemory mode always updates the same document ID

                if(isTemporary)
                    document.Close();

                return AgentToolResult.Ok($"HTML content imported successfully into document {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to import HTML: {ex.Message}");
            }
        }

        /// <summary>
        /// Imports Markdown content into a Word document.
        /// </summary>
        /// <param name="markdownContentOrFilePath">The markdown content as a string or the file path to a markdown file.</param>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the destination document.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(
            Name = "ImportMarkdown",
            Description = "Imports markdown content into a Word document. markdownContent / filePath: The markdown content as a string or the file path to a markdown file. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ImportMarkdown(
            [ToolParameter(Description = "The markdown content as a string or the file path to a markdown file")]
            string markdownContentOrFilePath,
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the destination document")]
            string? documentIdOrFilePath = null,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                bool isTemporary = false;
                // ── Open ────────────────────────────────────────────────────────
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                {
                    document = new WordDocument();
                    isTemporary = true;
                }

                string markdownContent;

                if (Mode == DocumentManagerMode.InMemory && File.Exists(markdownContentOrFilePath))
                {
                    // Mode 1: file path fallback
                    markdownContent = File.ReadAllText(markdownContentOrFilePath);
                }
                else if (Mode == DocumentManagerMode.DocumentStorage && StorageManager!.HasDocument(markdownContentOrFilePath))
                {
                    // Mode 2: get document stream from storage and read as Markdown
                    Stream? mdDocStream = StorageManager!.GetDocumentStream(markdownContentOrFilePath);
                    if (mdDocStream == null)
                        return AgentToolResult.Fail($"Markdown Document not found: {markdownContentOrFilePath}");
                    using (var reader = new StreamReader(mdDocStream, System.Text.Encoding.UTF8))
                    {
                        markdownContent = reader.ReadToEnd();
                    }
                    mdDocStream.Dispose();
                }
                else
                {
                    markdownContent = markdownContentOrFilePath;
                }

                // Import Markdown content
                using (MemoryStream stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(markdownContent)))
                {
                    WordDocument tempDocument = new WordDocument(stream, FormatType.Markdown);
                    if (document.LastSection == null)
                        document.AddSection();
                    EntityCollection destinationCollection = document.LastSection.Body.ChildEntities;
                    foreach (Entity entity in tempDocument.LastSection.Body.ChildEntities)
                    {
                        destinationCollection.Add(entity.Clone());
                    }
                    tempDocument.Close();
                }

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_md_imported.docx";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath ?? InMemoryManager!.ActiveDocumentId!; // InMemory mode always updates the same document ID

                if(isTemporary)
                    document.Close();

                return AgentToolResult.Ok($"Markdown content imported successfully into document {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to import Markdown: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets the Word document content as HTML.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <returns>Result containing the HTML content string or an error message.</returns>
        [Tool(
            Name = "GetHtml",
            Description = "Gets the Word document content as HTML using the given documentId or filePath. Returns the HTML string of a Word document.")]
        public AgentToolResult GetHtml(
            [ToolParameter(Description = "The ID of the document or file path")]
            string documentIdOrFilePath)
        {
            try
            {
                WordDocument? document = null;
                bool isTemporary = false;
                // ── Open ────────────────────────────────────────────────────────
                if (Mode == DocumentManagerMode.InMemory)
                {
                    // Mode 1: try manager first, then file path
                    if (InMemoryManager!.HasDocument(documentIdOrFilePath))
                    {
                        document = InMemoryManager.GetDocument(documentIdOrFilePath);
                    }
                    else if (File.Exists(documentIdOrFilePath))
                    {
                        document = new WordDocument(documentIdOrFilePath);
                        isTemporary = true;
                    }
                }
                else
                {
                    // Mode 2: use storage existence check, no File.Exists fallback
                    if (StorageManager!.HasDocument(documentIdOrFilePath))
                    {
                        document = OpenDocument(documentIdOrFilePath);
                        isTemporary = true; // transient copy from storage — must be closed
                    }
                }

                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                using (MemoryStream stream = new MemoryStream())
                {
                    document.SaveOptions.HtmlExportOmitXmlDeclaration = true;
                    document.Save(stream, FormatType.Html);
                    stream.Position = 0;
                    string htmlContent = System.Text.Encoding.UTF8.GetString(stream.ToArray());

                    if (isTemporary)
                        document.Close();

                    return AgentToolResult.Ok($"Generated HTML content from {documentIdOrFilePath} " + htmlContent, new { HtmlContent = htmlContent });
                }
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to get HTML: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets the Word document content as Markdown.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <returns>Result containing the Markdown content string or an error message.</returns>
        [Tool(
            Name = "GetMarkdown",
            Description = "Gets the Word document content as Markdown using the given documentId or filePath. Returns the Markdown content string of a Word document.")]
        public AgentToolResult GetMarkdown(
            [ToolParameter(Description = "The ID of the document or file path")]
            string documentIdOrFilePath)
        {
            try
            {
                WordDocument? document = null;
                bool isTemporary = false;
                // ── Open ────────────────────────────────────────────────────────
                if (Mode == DocumentManagerMode.InMemory)
                {
                    // Mode 1: try manager first, then file path
                    if (InMemoryManager!.HasDocument(documentIdOrFilePath))
                    {
                        document = InMemoryManager.GetDocument(documentIdOrFilePath);
                    }
                    else if (File.Exists(documentIdOrFilePath))
                    {
                        document = new WordDocument(documentIdOrFilePath);
                        isTemporary = true;
                    }
                }
                else
                {
                    // Mode 2: use storage existence check, no File.Exists fallback
                    if (StorageManager!.HasDocument(documentIdOrFilePath))
                    {
                        document = OpenDocument(documentIdOrFilePath);
                        isTemporary = true; // transient copy from storage — must be closed
                    }
                }

                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                // Export to Markdown format
                using (MemoryStream stream = new MemoryStream())
                {
                    document.Save(stream, FormatType.Markdown);
                    stream.Position = 0;
                    string markdownContent = System.Text.Encoding.UTF8.GetString(stream.ToArray());

                    if (isTemporary)
                        document.Close();

                    return AgentToolResult.Ok($"Generated Markdown content from {documentIdOrFilePath} " + markdownContent, new { MarkdownContent = markdownContent });
                }
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to get Markdown: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets the Word document content as plain text.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <returns>Result containing the text content string or an error message.</returns>
        [Tool(
            Name = "GetText",
            Description = "Gets the Word document content as text using the given documentId or filePath. Returns the text of a Word document.")]
        public AgentToolResult GetText(
            [ToolParameter(Description = "The ID of the document or file path")]
            string documentIdOrFilePath)
        {
            try
            {
                WordDocument? document = null;
                bool isTemporary = false;
                // ── Open ────────────────────────────────────────────────────────
                if (Mode == DocumentManagerMode.InMemory)
                {
                    // Mode 1: try manager first, then file path
                    if (InMemoryManager!.HasDocument(documentIdOrFilePath))
                    {
                        document = InMemoryManager.GetDocument(documentIdOrFilePath);
                    }
                    else if (File.Exists(documentIdOrFilePath))
                    {
                        document = new WordDocument(documentIdOrFilePath);
                        isTemporary = true;
                    }
                }
                else
                {
                    // Mode 2: use storage existence check, no File.Exists fallback
                    if (StorageManager!.HasDocument(documentIdOrFilePath))
                    {
                        document = OpenDocument(documentIdOrFilePath);
                        isTemporary = true; // transient copy from storage — must be closed
                    }
                }

                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                string text = document.GetText();

                if (isTemporary)
                    document.Close();

                return AgentToolResult.Ok($"Generated text content from {documentIdOrFilePath} " + text, new { Text = text });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to get text: {ex.Message}");
            }
        }

        /// <summary>
        /// Converts the document to the file system in the specified format (DocumentStorage mode only).
        /// </summary>
        /// <param name="documentIdOrFilePath">The input file path (DocumentStorage mode).</param>
        /// <param name="filePath">The destination file path where the converted document will be saved.</param>
        /// <param name="formatType">The output format: Docx, Doc, Rtf, Html, or Txt. Default is Docx.</param>
        /// <returns>Result indicating whether the document was converted successfully.</returns>
        [Tool(
            Name = "ConvertDocument",
            Description = "Converts the document to the file system in the specified format. Works only in DocumentStorage mode. documentIdOrFilePath: The input file path from storage. Supported formats: DOCX, DOC, RTF, HTML, TXT.")]
        public AgentToolResult ConvertDocument(
            [ToolParameter(Description = "The input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The file path to export to")]
            string filePath,
            [ToolParameter(Description = "The format: Docx, Doc, Rtf, Html, Txt. Defaults to Docx")]
            string? formatType = "Docx")
        {
            try
            {

                // Open the document from storage
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");


                // Save the document to storage 
                SaveDocument(filePath, document);

                return AgentToolResult.Ok($"Document exported successfully to {filePath}", new { FilePath = filePath });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to export document: {ex.Message}");
            }
        }

        /// <summary>
        /// Updates (rebuilds) the Table of Contents in a Word document.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="imageFormat">The image format: Png or Jpeg. Defaults to Png.</param>
        /// <param name="startPageIndex">The 1-based start page index. If null, starts from the first page.</param>
        /// <param name="endPageIndex">The 1-based end page index. If null, converts up to the last page.</param>
        /// <param name="outputDirectory">Output directory for saving the images. Defaults to current directory.</param>
        /// <returns>Result containing the list of exported image file paths.</returns>
        [Tool(
            Name = "ExportAsImage",
            Description = "Exports Word document pages as images (PNG or JPEG) to the output directory. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Optionally specify a page range using startPageIndex and endPageIndex (1-based). Returns the file paths of the exported images.")]
        public AgentToolResult ExportAsImage(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The image format: Png or Jpeg. Defaults to Png")]
            string? imageFormat = "Png",
            [ToolParameter(Description = "The 1-based start page index. If null, starts from the first page")]
            int? startPageIndex = null,
            [ToolParameter(Description = "The 1-based end page index. If null, converts up to the last page")]
            int? endPageIndex = null,
            [ToolParameter(Description = "Output directory for saving the images. Defaults to current directory.")]
            string outputDirectory = ".")
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                // Determine image type
                ExportImageFormat exportFormat = imageFormat?.Equals("Jpeg", StringComparison.OrdinalIgnoreCase) == true
                    ? ExportImageFormat.Jpeg
                    : ExportImageFormat.Png;

                string fileExtension = exportFormat == ExportImageFormat.Jpeg ? ".jpeg" : ".png";

                // Convert document pages to images
                using var renderer = new Syncfusion.DocIORenderer.DocIORenderer();
                Stream[] imageStreams;
                int startPage = 1;

                if (startPageIndex != null && endPageIndex != null)
                {
                    startPage = startPageIndex.Value;
                    imageStreams = document.RenderAsImages(startPageIndex.Value, endPageIndex.Value);
                }
                else
                {
                    imageStreams = document.RenderAsImages();
                }


                var exportedFilePaths = new List<string>();

                for (int i = 0; i < imageStreams.Length; i++)
                {
                    int pageNumber = startPage + i;
                    string fileName = $"{Path.GetFileNameWithoutExtension(documentIdOrFilePath)}_Page{pageNumber}{fileExtension}";
                    string fullPath = Path.Combine(outputDirectory, fileName);

                    if (Mode == DocumentManagerMode.InMemory)
                    {
                        // Create output directory if it doesn't exist
                        var outputDir = Path.GetDirectoryName(fullPath);
                        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                        {
                            Directory.CreateDirectory(outputDir);
                        }
                        
                        using (var fileStream = new FileStream(fullPath, FileMode.Create, FileAccess.Write))
                        {
                            imageStreams[i].Position = 0;
                            imageStreams[i].CopyTo(fileStream);
                        }
                    }
                    else
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            imageStreams[i].Position = 0;
                            imageStreams[i].CopyTo(memoryStream);
                            memoryStream.Position = 0;
                            SaveFile(fullPath, memoryStream);
                        }
                    }

                    imageStreams[i].Dispose();
                    exportedFilePaths.Add(fullPath);
                }

                return AgentToolResult.Ok(
                    $"Successfully exported {exportedFilePaths.Count} page(s) as {imageFormat} images",
                    new { FilePaths = exportedFilePaths.ToArray(), PageCount = exportedFilePaths.Count });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to export document as images: {ex.Message}");
            }
        }
    }
}
