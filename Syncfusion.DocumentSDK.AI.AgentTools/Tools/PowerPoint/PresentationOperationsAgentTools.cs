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
using System.IO;
using System.Linq;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.Presentation;
using Syncfusion.PresentationRenderer;

namespace Syncfusion.AI.AgentTools.PowerPoint
{
    /// <summary>
    /// Provides AI agent tools for PowerPoint presentation manipulation operations.
    /// Handles merging and splitting of presentations.
    /// </summary>
    public class PresentationOperationsAgentTools : AgentToolBase<IPresentation>
    {
        private readonly PresentationManager _manager;
        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationOperationsAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The presentation manager for managing PowerPoint presentations.</param>
        public PresentationOperationsAgentTools(PresentationManager manager)
            : base(manager, DocumentType.PowerPoint) 
        {
            _manager = manager;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationOperationsAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public PresentationOperationsAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.PowerPoint) { }

        /// <summary>
        /// Merges multiple PowerPoint presentations into a single destination presentation.
        /// </summary>
        /// <param name="destinationDocumentId">The ID of the destination presentation.</param>
        /// <param name="sourceDocumentIds">Array of source document IDs or file paths to merge.</param>
        /// <param name="pasteOption">Paste option for formatting: 'SourceFormatting' or 'UseDestinationTheme'. Default is 'SourceFormatting'.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "MergePresentations", Description = "Merges multiple PowerPoint Presentations into a single destination presentation. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode) of the destination presentation.")]
        public AgentToolResult MergePresentations(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the destination presentation")] string documentIdOrFilePath,
            [ToolParameter(Description = "A collection of document IDs or file paths to merge")] string[] sourceDocumentIds,
            [ToolParameter(Description = "Paste option: 'SourceFormatting' (default) or 'UseDestinationTheme'")] string pasteOption = "SourceFormatting",
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);
                ArgumentNullException.ThrowIfNull(sourceDocumentIds);

                if (sourceDocumentIds.Length == 0)
                    return AgentToolResult.Fail("No source documents provided. Specify comma-separated source document IDs or file paths.");

                if (!TryParsePasteOption(pasteOption, out PasteOptions parsedPasteOption))
                    return AgentToolResult.Fail($"Invalid paste option: '{pasteOption}'. Use 'SourceFormatting' or 'UseDestinationTheme'.");

                var destinationPresentation = OpenDocument(documentIdOrFilePath);
                if (destinationPresentation == null)
                    return AgentToolResult.Fail($"Destination presentation not found: {documentIdOrFilePath}");

                int slidesMerged = 0;
                foreach (var sourceId in sourceDocumentIds)
                {
                    var sourcePresentation = OpenDocument(sourceId);
                    if (sourcePresentation == null)
                        return AgentToolResult.Fail($"Source not found: {sourceId}");

                    // Clone and add all slides from source to destination with the specified paste option
                    foreach (ISlide slide in sourcePresentation.Slides)
                    {
                        destinationPresentation.Slides.Add(slide.Clone(), parsedPasteOption);
                        slidesMerged++;
                    }

                    if (Mode == DocumentManagerMode.DocumentStorage)
                        sourcePresentation.Close();
                }

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_merged.pptx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, destinationPresentation);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Successfully merged {sourceDocumentIds.Length} presentation(s) with {slidesMerged} total slides into {outputKey}",
                    new { DestinationId = outputKey, SourceCount = sourceDocumentIds.Length, SlidesMerged = slidesMerged, PasteOption = pasteOption });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to merge PowerPoint presentations: {ex.Message}");
            }
        }

        /// <summary>
        /// Splits a single PowerPoint presentation into multiple presentations based on split rules.
        /// </summary>
        /// <param name="documentId">The ID of the presentation to split.</param>
        /// <param name="splitRules">Split rules: 'sections', 'layout', or specific slide numbers (e.g., '1,3,5').</param>
        /// <param name="pasteOption">Paste option for formatting: 'SourceFormatting' or 'UseDestinationTheme'. Default is 'SourceFormatting'.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing array of document IDs of the split presentations.</returns>
        /// <remarks>
        /// Split by specific slide numbers (1-based).
        /// The provided numbers define split boundaries � each number marks the last slide of a group.
        /// Remaining slides after the last boundary form an additional group.
        /// </remarks>
        [Tool(Name = "SplitPresentation", Description = "Splits a single PowerPoint presentation into multiple presentations based on the specified splitRules (sections, layout type, or slide numbers (e.g., '1,3,5')). documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult SplitPresentation(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath,
            [ToolParameter(Description = "Split rules: 'sections', 'layout', or slide numbers (e.g., '1,3,5')")] string splitRules,
            [ToolParameter(Description = "Paste option: 'SourceFormatting' (default) or 'UseDestinationTheme'")] string pasteOption = "SourceFormatting",
            [ToolParameter(Description = "Output file path prefix for saving the split results (DocumentStorage mode only). Index will be appended, e.g. 'output_split_1.pptx'.")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);
                ArgumentNullException.ThrowIfNull(splitRules);

                if (!TryParsePasteOption(pasteOption, out PasteOptions parsedPasteOption))
                    return AgentToolResult.Fail($"Invalid paste option: '{pasteOption}'. Use 'SourceFormatting' or 'UseDestinationTheme'.");

                var sourcePresentation = OpenDocument(documentIdOrFilePath);
                if (sourcePresentation == null)
                    return AgentToolResult.Fail($"Presentation not found: {documentIdOrFilePath}");

                bool isSaveNeeded = false;

                if (outputFilePath == null)
                {
                    if (Mode == DocumentManagerMode.DocumentStorage)
                    {
                        outputFilePath = "output_split.pptx";
                    }
                }
                else
                {
                    isSaveNeeded = true;
                }
             
               

                // Helper to build a unique key for each split segment
                int splitIndex = 0;
                string CreateSplitKey()
                {
                    splitIndex++;                   
                    
                    if (Mode == DocumentManagerMode.InMemory)
                    {
                        InMemoryManager!.CreateDocument();
                        string documentId= InMemoryManager.ActiveDocumentId ?? throw new InvalidOperationException("Failed to create split presentation");
                        return documentId;
                    }
                    else
                    {
                        string ext = Path.GetExtension(outputFilePath);
                        if (!string.IsNullOrEmpty(ext))
                        {
                            ext = ext.StartsWith(".") ? ext : "." + ext;

                        }
                        else
                        {
                            ext = ".pptx";
                        }
                        string nameWithoutExt = Path.GetFileNameWithoutExtension(outputFilePath);
                        string dir = Path.GetDirectoryName(outputFilePath) ?? string.Empty;
                        return Path.Combine(dir, $"{nameWithoutExt}_{splitIndex}{ext}");
                    }
                    
                  
                }

                IPresentation GetOrCreatePres(string key)
                {
                    if (Mode == DocumentManagerMode.InMemory)
                        return InMemoryManager!.GetDocument(key)!;
                    return Syncfusion.Presentation.Presentation.Create();
                }

                List<string> splitDocumentIds = new List<string>();

                if (splitRules.ToLowerInvariant() == "sections"|| splitRules.ToLowerInvariant() == "section")
                {
                    // Split by actual sections defined in the presentation
                    if (sourcePresentation.Sections.Count == 0)
                        return AgentToolResult.Fail("The presentation does not contain any sections to split by.");

                    foreach (ISection section in sourcePresentation.Sections)
                    {
                        if (section.Slides.Count == 0)
                            continue;

                        // Clone all slides in this section
                        ISlides clonedSlides = section.Clone();

                        string newKey = CreateSplitKey();
                        var newPres = GetOrCreatePres(newKey);
                        foreach (ISlide clonedSlide in clonedSlides)
                        {
                            newPres.Slides.Add(clonedSlide, parsedPasteOption);
                        }

                        if (Mode == DocumentManagerMode.DocumentStorage || isSaveNeeded)
                        {
                                                     
                            if (isSaveNeeded)
                            {
                                string dir = Path.GetDirectoryName(outputFilePath) ?? string.Empty;
                                newKey = Path.Combine(dir, newKey + ".pptx");
                            }                            
                            SaveDocument(newKey, newPres);
                           
                        }
                        if(Mode == DocumentManagerMode.DocumentStorage)
                        {
                            newPres.Close();
                        }
                        splitDocumentIds.Add(newKey);
                    }
                }
                else if (splitRules.ToLowerInvariant() == "layout")
                {
                    // Split by slide layout type - group slides sharing the same layout
                    var layoutGroups = new Dictionary<string, List<ISlide>>();
                    var layoutOrder = new List<string>();

                    foreach (ISlide slide in sourcePresentation.Slides)
                    {
                        string layoutName = slide.LayoutSlide?.Name ?? "Unknown";
                        if (!layoutGroups.ContainsKey(layoutName))
                        {
                            layoutGroups[layoutName] = new List<ISlide>();
                            layoutOrder.Add(layoutName);
                        }
                        layoutGroups[layoutName].Add(slide);
                    }

                    foreach (var layoutName in layoutOrder)
                    {
                        var slides = layoutGroups[layoutName];

                        string newKey = CreateSplitKey();
                        var newPres = GetOrCreatePres(newKey);
                        foreach (ISlide slide in slides)
                        {
                            newPres.Slides.Add(slide.Clone(), parsedPasteOption);
                        }
                        if (Mode == DocumentManagerMode.DocumentStorage || isSaveNeeded)
                        {

                            if (isSaveNeeded)
                            {
                                string dir = Path.GetDirectoryName(outputFilePath) ?? string.Empty;
                                newKey = Path.Combine(dir, newKey + ".pptx");
                            }
                            SaveDocument(newKey, newPres);

                        }
                        if (Mode == DocumentManagerMode.DocumentStorage)
                            newPres.Close();

                        splitDocumentIds.Add(newKey);
                    }
                }
                else
                {
                    // Split by specific slide numbers (1-based).
                    // The provided numbers define split boundaries � each number marks the last slide of a group.
                    // Remaining slides after the last boundary form an additional group.
                    string[] slideNumbers = splitRules.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

                    if (slideNumbers.Length == 0)
                        return AgentToolResult.Fail("No slide numbers provided. Specify comma-separated slide numbers (e.g., '1,3,5').");

                    List<int> indices = new List<int>();
                    List<string> invalidEntries = new List<string>();
                    int totalSlides = sourcePresentation.Slides.Count;

                    foreach (var numStr in slideNumbers)
                    {
                        if (int.TryParse(numStr, out int slideNum))
                        {
                            if (slideNum < 1 || slideNum > totalSlides)
                            {
                                invalidEntries.Add($"'{numStr}' (out of range 1-{totalSlides})");
                            }
                            else
                            {
                                int zeroBasedIndex = slideNum - 1;
                                if (!indices.Contains(zeroBasedIndex))
                                    indices.Add(zeroBasedIndex);
                            }
                        }
                        else
                        {
                            invalidEntries.Add($"'{numStr}' (not a valid number)");
                        }
                    }

                    if (invalidEntries.Count > 0)
                        return AgentToolResult.Fail($"Invalid slide number(s): {string.Join(", ", invalidEntries)}. The presentation has {totalSlides} slide(s).");

                    if (indices.Count == 0)
                        return AgentToolResult.Fail("No valid slide numbers provided after filtering duplicates.");

                    indices.Sort();

                    // Create presentations for each segment defined by the split boundaries
                    for (int i = 0; i < indices.Count; i++)
                    {
                        int startIndex = i == 0 ? 0 : indices[i - 1] + 1;
                        int endIndex = indices[i];

                        if (startIndex > endIndex)
                            continue;

                        string newKey = CreateSplitKey();
                        var newPres = GetOrCreatePres(newKey);
                        for (int j = startIndex; j <= endIndex; j++)
                        {
                            newPres.Slides.Add(sourcePresentation.Slides[j].Clone(), parsedPasteOption);
                        }
                        if (Mode == DocumentManagerMode.DocumentStorage || isSaveNeeded)
                        {

                            if (isSaveNeeded)
                            {
                                string dir = Path.GetDirectoryName(outputFilePath) ?? string.Empty;
                                newKey = Path.Combine(dir, newKey + ".pptx");
                            }
                            SaveDocument(newKey, newPres);

                        }
                        if (Mode == DocumentManagerMode.DocumentStorage)
                            newPres.Close();

                        splitDocumentIds.Add(newKey);
                    }

                    // Capture remaining slides after the last split boundary
                    int lastBoundary = indices[indices.Count - 1];
                    if (lastBoundary + 1 < totalSlides)
                    {
                        string remainingKey = CreateSplitKey();
                        var remainingPres = GetOrCreatePres(remainingKey);
                        for (int j = lastBoundary + 1; j < totalSlides; j++)
                        {
                            remainingPres.Slides.Add(sourcePresentation.Slides[j].Clone(), parsedPasteOption);
                        }
                        
                        if (Mode == DocumentManagerMode.DocumentStorage || isSaveNeeded)
                        {

                            if (isSaveNeeded)
                            {
                                string dir = Path.GetDirectoryName(outputFilePath) ?? string.Empty;
                                remainingKey = Path.Combine(dir, remainingKey + ".pptx");
                            }
                            SaveDocument(remainingKey, remainingPres);

                        }
                        if (Mode == DocumentManagerMode.DocumentStorage)
                            remainingPres.Close();

                        splitDocumentIds.Add(remainingKey);
                    }
                }

                return AgentToolResult.Ok(
                    $"Successfully split presentation into {splitDocumentIds.Count} presentation(s)",
                    new { SplitDocumentIds = splitDocumentIds.ToArray(), Count = splitDocumentIds.Count, SplitRules = splitRules, PasteOption = pasteOption });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to split PowerPoint presentation: {ex.Message}");
            }
        }

        /// <summary>
        /// Imports Markdown content into a PowerPoint presentation.
        /// </summary>
        /// <param name="markdownContentOrFilePath">The markdown content as a string or the file path to a markdown file.</param>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the destination presentation.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(
            Name = "ImportMarkdown",
            Description = "Imports markdown content into a PowerPoint presentation. markdownContent / filePath: The markdown content as a string or the file path to a markdown file. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ImportMarkdown(
            [ToolParameter(Description = "The markdown content as a string or the file path to a markdown file")]
            string markdownContentOrFilePath,
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the destination presentation")]
            string? documentIdOrFilePath = null,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                bool isTemporary = false;
                // ── Open ────────────────────────────────────────────────────────
                var presentation = OpenDocument(documentIdOrFilePath);
                if (presentation == null && Mode == DocumentManagerMode.DocumentStorage)
                {
                    presentation = Syncfusion.Presentation.Presentation.Create();
                    isTemporary = true;
                }
                else if (presentation == null)
                {
                    presentation = _manager.CreateDocument();
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
                    using (IPresentation tempPresentation = Syncfusion.Presentation.Presentation.Open(stream))
                    {
                        foreach (ISlide slide in tempPresentation.Slides)
                        {
                            presentation.Slides.Add(slide.Clone(), PasteOptions.SourceFormatting);
                        }
                    }
                }

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_md_imported.pptx";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, presentation);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath ?? InMemoryManager!.ActiveDocumentId!;

                if(isTemporary)
                    presentation.Close();

                return AgentToolResult.Ok($"Markdown content imported successfully into presentation {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to import Markdown: {ex.Message}");
            }
        }

        /// <summary>
        /// Gets the PowerPoint presentation content as Markdown.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <returns>Result containing the Markdown content string or an error message.</returns>
        [Tool(
            Name = "GetMarkdown",
            Description = "Gets the PowerPoint presentation content as Markdown using the given documentId or filePath. Returns the Markdown content string of a PowerPoint presentation.")]
        public AgentToolResult GetMarkdown(
            [ToolParameter(Description = "The ID of the presentation or file path")]
            string documentIdOrFilePath)
        {
            try
            {
                IPresentation? presentation = null;
                bool isTemporary = false;
                // ── Open ────────────────────────────────────────────────────────
                if (Mode == DocumentManagerMode.InMemory)
                {
                    // Mode 1: try manager first, then file path
                    if (InMemoryManager!.HasDocument(documentIdOrFilePath))
                    {
                        presentation = InMemoryManager.GetDocument(documentIdOrFilePath);
                    }
                    else if (File.Exists(documentIdOrFilePath))
                    {
                        presentation = Syncfusion.Presentation.Presentation.Open(documentIdOrFilePath);
                        isTemporary = true;
                    }
                }
                else
                {
                    // Mode 2: use storage existence check, no File.Exists fallback
                    if (StorageManager!.HasDocument(documentIdOrFilePath))
                    {
                        presentation = OpenDocument(documentIdOrFilePath);
                        isTemporary = true; // transient copy from storage — must be closed
                    }
                }

                if (presentation == null)
                    return AgentToolResult.Fail($"Presentation not found: {documentIdOrFilePath}");

                // Export to Markdown format
                using (MemoryStream stream = new MemoryStream())
                {
                    presentation.Save(stream, Syncfusion.Presentation.FormatType.Markdown);
                    stream.Position = 0;
                    string markdownContent = System.Text.Encoding.UTF8.GetString(stream.ToArray());

                    if (isTemporary)
                        presentation.Close();

                    return AgentToolResult.Ok($"Generated Markdown content from {documentIdOrFilePath} " + markdownContent, new { MarkdownContent = markdownContent });
                }
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to get Markdown: {ex.Message}");
            }
        }

        /// Exports PowerPoint presentation slides as images to the file system.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="outputDirectory">The directory to save the exported image files.</param>
        /// <param name="imageFormat">The image format: Png or Jpeg. Defaults to Png.</param>
        /// <param name="startSlideIndex">The 1-based start slide index. If null, starts from the first slide.</param>
        /// <param name="endSlideIndex">The 1-based end slide index. If null, converts up to the last slide.</param>
        /// <returns>Result containing the list of exported image file paths.</returns>
        [Tool(
            Name = "ExportAsImage",
            Description = "Exports PowerPoint presentation slides as images (PNG or JPEG) to the output directory. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Optionally specify a slide range using startSlideIndex and endSlideIndex (1-based). Returns the file paths of the exported images.")]
        public AgentToolResult ExportAsImage(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The directory to save the exported image files")]
            string outputDirectory,
            [ToolParameter(Description = "The image format: Png or Jpeg. Defaults to Png")]
            string? imageFormat = "Png",
            [ToolParameter(Description = "The 1-based start slide index. If null, starts from the first slide")]
            int? startSlideIndex = null,
            [ToolParameter(Description = "The 1-based end slide index. If null, converts up to the last slide")]
            int? endSlideIndex = null)
        {
            try
            {
                var presentation = OpenDocument(documentIdOrFilePath);
                if (presentation == null)
                    return AgentToolResult.Fail($"Presentation not found: {documentIdOrFilePath}");

                // Determine image type
                ExportImageFormat exportFormat = imageFormat?.Equals("Jpeg", StringComparison.OrdinalIgnoreCase) == true
                    ? ExportImageFormat.Jpeg
                    : ExportImageFormat.Png;

                string fileExtension = exportFormat == ExportImageFormat.Jpeg ? ".jpeg" : ".png";

                int totalSlides = presentation.Slides.Count;
                int start = startSlideIndex ?? 1;
                int end = endSlideIndex ?? totalSlides;

                // Validate slide range
                if (start < 1 || start > totalSlides)
                    return AgentToolResult.Fail($"Invalid startSlideIndex: {start}. Must be between 1 and {totalSlides}.");
                if (end < start || end > totalSlides)
                    return AgentToolResult.Fail($"Invalid endSlideIndex: {end}. Must be between {start} and {totalSlides}.");

                // Convert slides to images
                presentation.PresentationRenderer = new Syncfusion.PresentationRenderer.PresentationRenderer();

                var exportedFilePaths = new List<string>();

                for (int i = start - 1; i < end; i++)
                {
                    int slideNumber = i + 1;
                    Stream imageStream = presentation.Slides[i].ConvertToImage(exportFormat);

                    string fileName = $"{Path.GetFileNameWithoutExtension(documentIdOrFilePath)}_Slide{slideNumber}{fileExtension}";
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
                            imageStream.Position = 0;
                            imageStream.CopyTo(fileStream);
                        }
                    }
                    else
                    {
                        using (var memoryStream = new MemoryStream())
                        {
                            imageStream.Position = 0;
                            imageStream.CopyTo(memoryStream);
                            memoryStream.Position = 0;
                            SaveFile(fullPath, memoryStream);
                        }
                    }

                    imageStream.Dispose();
                    exportedFilePaths.Add(fullPath);
                }

                if (Mode == DocumentManagerMode.DocumentStorage)
                    presentation.Close();

                return AgentToolResult.Ok(
                    $"Successfully exported {exportedFilePaths.Count} slide(s) as {imageFormat} images",
                    new { FilePaths = exportedFilePaths.ToArray(), SlideCount = exportedFilePaths.Count });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to export presentation as images: {ex.Message}");
            }
        }
        /// <summary>
        /// Parses the paste option string and converts it to a <see cref="PasteOptions"/> value.
        /// </summary>
        /// <param name="result">The parsed <see cref="PasteOptions"/> value.</param>
        /// <returns><c>true</c> if parsing succeeded; otherwise, <c>false</c>.</returns>
        private static bool TryParsePasteOption(string pasteOption, out PasteOptions result)
        {
            switch (pasteOption?.ToLowerInvariant())
            {
                case "sourceformatting":
                    result = PasteOptions.SourceFormatting;
                    return true;
                case "usedestinationtheme":
                    result = PasteOptions.UseDestinationTheme;
                    return true;
                default:
                    result = PasteOptions.SourceFormatting;
                    return false;
            }
        }
        /// <summary>
        /// Converts the document to the file system in the specified format (DocumentStorage mode only).
        /// </summary>
        [Tool(
            Name = "ConvertPresentation",
            Description = "Converts the presentation to the file system in the specified format. Works only in DocumentStorage mode. documentIdOrFilePath: The input file path from storage. Supported formats: PPTX, PPTM, POTX, POTM, MD.")]
        public AgentToolResult ConvertPresentation(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The file path to export to")]
            string filePath,
            [ToolParameter(Description = "The format: PPTX, PPTM, POTX, POTM, MD. Defaults to Pptx")]
            string? formatType = "Pptx")
        {
            try
            {

                // Open the document from storage
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                // Ensure correct file extension based on format
                string extension = formatType?.ToUpperInvariant() switch
                {
                    "PPTX" => ".pptx",
                    "PPTM" => ".pptm",
                    "POTX" => ".potx",
                    "POTM" => ".potm",
                    "MD" => ".md",
                    _ => ".pptx"
                };

                string outputPath = filePath;
                if (!outputPath.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
                {
                    outputPath = Path.ChangeExtension(outputPath, extension);
                }

                // Save the document to storage 
                SaveDocument(outputPath, document);

                return AgentToolResult.Ok($"Document exported successfully to {outputPath}", new { FilePath = outputPath });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to export document: {ex.Message}");
            }
        }
    }
}
