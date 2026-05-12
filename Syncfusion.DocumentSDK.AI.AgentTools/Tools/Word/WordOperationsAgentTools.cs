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
using Syncfusion.DocIORenderer;
using System;
using System.IO;
using System.Text.RegularExpressions;

namespace Syncfusion.AI.AgentTools.Word
{
    /// <summary>
    /// Provides agent tools for document manipulation and comparison operations.
    /// Handles merging, splitting, and comparing Word documents.
    /// </summary>
    public class WordOperationsAgentTools : AgentToolBase<WordDocument>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WordOperationsAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The document manager for managing Word documents.</param>
        public WordOperationsAgentTools(WordDocumentManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordOperationsAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public WordOperationsAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Merges multiple Word documents into a single destination document.
        /// </summary>
        /// <param name="destinationDocumentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the destination WordDocument.</param>
        /// <param name="documentIdsOrFilePaths">A collection of document IDs or file paths to merge into the destination document.</param>
        /// <param name="pasteOption">Specifies how formatting is handled during merge, such as using destination styles or keeping source formatting.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(
            Name = "MergeDocuments",
            Description = "Merges multiple Word documents into a single destination document. destinationDocumentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode) of the destination WordDocument. documentIds/filePaths: A collection of document IDs or file paths to merge into the destination document.")]
        public AgentToolResult MergeDocuments(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the destination WordDocument")]
            string destinationDocumentIdOrFilePath,
            [ToolParameter(Description = "A collection of document IDs or file paths to merge")]
            string[] documentIdsOrFilePaths,
            [ToolParameter(Description = "Import option to control formatting when merging. Accepted values: 'UseDestinationStyles' (default) - applies destination document styles; 'KeepSourceFormatting' - applies default Normal style and preserves all other formatting as direct formatting; 'KeepTextOnly' - imports only plain text without any formatting; 'ListContinueNumbering' - imports content and continues existing list numbering; 'ListRestartNumbering' - imports content and restarts list numbering; 'MergeFormatting' - applies the formatting of surrounding content in the destination document.")] 
            string pasteOption = "UseDestinationStyles",
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var destinationDoc = OpenDocument(destinationDocumentIdOrFilePath);
                if (destinationDoc == null)
                    return AgentToolResult.Fail($"Destination document not found: {destinationDocumentIdOrFilePath}");

                ImportOptions importOptions = pasteOption.ToLowerInvariant() switch
                {
                    "keepsourceformatting" => ImportOptions.KeepSourceFormatting,
                    "keeptextonly" => ImportOptions.KeepTextOnly,
                    "listcontinuenumbering" => ImportOptions.ListContinueNumbering,
                    "listrestartnumbering" => ImportOptions.ListRestartNumbering,
                    "mergeformatting" => ImportOptions.MergeFormatting,
                    _ => ImportOptions.UseDestinationStyles
                };

                foreach (var idOrPath in documentIdsOrFilePaths)
                {
                    WordDocument? sourceDoc = null;
                    bool isTemporary = false;

                    // Try to open the document using the base class helper
                    sourceDoc = OpenDocument(idOrPath);
                    if (sourceDoc == null)
                    {
                        // Fallback to direct file open if not in manager
                        sourceDoc = new WordDocument(idOrPath);
                        isTemporary = true;
                    }

                    if (sourceDoc != null)
                    {
                        destinationDoc.ImportContent(sourceDoc, importOptions);

                        if (isTemporary)
                            sourceDoc.Close();
                    }
                }

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_merged.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, destinationDoc);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = destinationDocumentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Documents merged successfully into {outputKey} using {importOptions}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to merge documents: {ex.Message}");
            }
        }

        /// <summary>
        /// Splits a Word document into multiple documents based on split rules.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the main document to split.</param>
        /// <param name="splitRules">Defines how the document is split — by sections, headings, bookmarks, or placeholder text.</param>
        /// <param name="placeholderText">The regex pattern used to identify placeholder markers when splitRules is 'SplitByPlaceholderText'. Defaults to "&lt;&lt;(.*)&gt;&gt;".</param>
        /// <param name="outputFilePath">Output file path prefix for saving the split results. If empty, generated file names are used.</param>
        /// <returns>Result containing the IDs of generated documents.</returns>
        [Tool(
            Name = "SplitDocument",
            Description = "Splits a single Word document into multiple documents based on the specified splitRules. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Returns the IDs of the documents generated after splitting. Accepted values for splitRules: 'SplitBySection' (default) - splits into one document per section; 'SplitByHeading' - splits at each Heading 1 paragraph; 'SplitByBookmark' - splits by extracting each bookmark's content; 'SplitByPlaceholderText' - splits by extracting content between paired placeholder markers (e.g. <<start>> ... <<end>>); 'SplitByPageBreak' - splits at each explicit page break or paragraph with PageBreakBefore formatting.")]
        public AgentToolResult SplitDocument(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the main document to split")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "Defines how the document is split. Accepted values: 'SplitBySection' (default) - splits the document at each section break; 'SplitByHeading' - splits the document at each Heading 1 paragraph; 'SplitByBookmark' - splits the document by extracting content within each bookmark; 'SplitByPlaceholderText' - splits the document by extracting content between paired placeholder markers; 'SplitByPageBreak' - splits the document at each explicit page break (Break entity with BreakType.PageBreak) or paragraph formatted with PageBreakBefore.")]
            string splitRules = "SplitBySection",
            [ToolParameter(Description = "The regex pattern used to find placeholder markers when splitRules is 'SplitByPlaceholderText'. Defaults to '<<(.*)>>' which matches paired markers like <<SectionStart>> and <<SectionEnd>>.")]
            string placeholderText = "<<(.*)>>",
            [ToolParameter(Description = "Output file path prefix for saving the split results.")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                List<string> splitDocumentIds = splitRules.ToLowerInvariant() switch
                {
                    "splitbysection" or "sections"                => SplitBySection(document, outputFilePath),
                    "splitbyheading" or "headings"                => SplitByHeading(document, outputFilePath),
                    "splitbybookmark" or "bookmark"               => SplitByBookmark(document, outputFilePath),
                    "splitbyplaceholdertext" or "placeholdertext" => SplitByPlaceholderText(document, placeholderText, outputFilePath),
                    "splitbypagebreak" or "pagebreak"             => SplitByPageBreak(document, outputFilePath),
                    _                                             => SplitBySection(document, outputFilePath)
                };

                return AgentToolResult.Ok(
                    $"Successfully split document into {splitDocumentIds.Count} document(s) using {splitRules}",
                    new { SplitDocumentIds = splitDocumentIds.ToArray(), Count = splitDocumentIds.Count, SplitRules = splitRules });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to split document: {ex.Message}");
            }
        }

        /// <summary>
        /// Create the split key for the split document based on the output file path and document ID. 
        /// </summary>
        string CreateSplitKey(string outputFilePath, string documentId)
        {           
            
            string ext = Path.GetExtension(outputFilePath+ documentId);
            if (!string.IsNullOrEmpty(ext))
            {
                ext = ext.StartsWith(".") ? ext : "." + ext;

            }
            else
            {
                ext = ".docx";
            }
            string nameWithoutExt = Path.GetFileNameWithoutExtension(documentId);
            string dir = Path.GetDirectoryName(outputFilePath) ?? string.Empty;
            return Path.Combine(dir, $"{nameWithoutExt}_{ext}");
        }
        /// <summary>
        /// Splits the Word document into one document per section.
        /// </summary>
        private List<string> SplitBySection(WordDocument document, string? outputFilePath = null)
        {
            var splitDocumentIds = new List<string>();
            //Iterate each section from Word document
            for (int i = 0; i < document.Sections.Count; i++)
            {
                //Create new Word document
                WordDocument newDocument = new WordDocument();
                //Add cloned section into new Word document
                newDocument.Sections.Add(document.Sections[i].Clone());
                
                string newDocumentId;
                if (Mode == DocumentManagerMode.InMemory)
                {
                    ((WordDocumentManager)InMemoryManager!).AddDocument(newDocument);
                    newDocumentId = InMemoryManager.ActiveDocumentId ?? throw new InvalidOperationException("Failed to create split Word document");
                }
                else
                {
                    newDocumentId = $"split_section_{i + 1}_{Guid.NewGuid():N}.docx";
                    newDocumentId= CreateSplitKey(outputFilePath, newDocumentId);
                    SaveDocument(newDocumentId, newDocument);
                }
                splitDocumentIds.Add(newDocumentId);
            }
            return splitDocumentIds;
        }

        /// <summary>
        /// Splits the Word document at each Heading 1 paragraph into separate documents.
        /// </summary>
        private List<string> SplitByHeading(WordDocument document, string? outputFilePath = null)
        {
            var splitDocumentIds = new List<string>();
            WordDocument newDocument = null;
            WSection newSection = null;

            foreach (WSection section in document.Sections)
            {
                // Clone the section and add into new document.
                if (newDocument != null)
                    newSection = AddSection(newDocument, section);
                //Iterate each child entity in the Word document.
                foreach (TextBodyItem item in section.Body.ChildEntities)
                {
                    //If item is paragraph, then check for heading style and split.
                    //else, add the item into new document.
                    if (item is WParagraph paragraph)
                    {
                        //If paragraph has Heading 1 style, save traversed content as a separate document
                        //and create a new document for the new heading's content.
                        if (paragraph.StyleName == "Heading 1")
                        {
                            if (newDocument != null)
                            {
                                string newDocumentId;
                                if (Mode == DocumentManagerMode.InMemory)
                                {
                                    newDocumentId = InMemoryManager!.ActiveDocumentId ?? throw new InvalidOperationException("Failed to create split Word document");
                                }
                                else
                                {
                                    newDocumentId = $"split_{Guid.NewGuid():N}.docx";
                                    newDocumentId = CreateSplitKey(outputFilePath, newDocumentId);
                                    SaveDocument(newDocumentId, newDocument);
                                }
                                splitDocumentIds.Add(newDocumentId);
                            }
                            //Create new document for new heading content.
                            newDocument = new WordDocument();
                            if (Mode == DocumentManagerMode.InMemory)
                            {
                                ((WordDocumentManager)InMemoryManager!).AddDocument(newDocument);
                            }
                            newSection = AddSection(newDocument, section);
                            AddEntity(newSection, paragraph);
                        }
                        else if (newDocument != null)
                            AddEntity(newSection, paragraph);
                    }
                    else
                        AddEntity(newSection, item);
                }
            }
            //Save the remaining content as a separate document.
            if (newDocument != null)
            {
                string newDocumentId;
                if (Mode == DocumentManagerMode.InMemory)
                {
                    newDocumentId = InMemoryManager!.ActiveDocumentId ?? throw new InvalidOperationException("Failed to create split Word document");
                }
                else
                {
                    newDocumentId = $"split_{Guid.NewGuid():N}.docx";
                    newDocumentId = CreateSplitKey(outputFilePath, newDocumentId);
                    SaveDocument(newDocumentId, newDocument);
                }
                splitDocumentIds.Add(newDocumentId);
            }
            return splitDocumentIds;
        }

        /// <summary>
        /// Splits the Word document by extracting the content within each bookmark into a separate document.
        /// </summary>
        private List<string> SplitByBookmark(WordDocument document, string? outputFilePath = null)
        {
            var splitDocumentIds = new List<string>();
            //Create the bookmark navigator instance to access the bookmarks.
            BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
            BookmarkCollection bookmarkCollection = document.Bookmarks;
            //Iterate each bookmark in the Word document.
            foreach (Bookmark bookmark in bookmarkCollection)
            {
                //Move the virtual cursor to the location before the end of the bookmark.
                bookmarksNavigator.MoveToBookmark(bookmark.Name);
                //Get the bookmark content as WordDocumentPart.
                WordDocumentPart documentPart = bookmarksNavigator.GetContent();
                WordDocument newDocument = documentPart.GetAsWordDocument();
                
                string newDocumentId;
                if (Mode == DocumentManagerMode.InMemory)
                {
                     ((WordDocumentManager)InMemoryManager!).AddDocument(newDocument);
                    newDocumentId = InMemoryManager.ActiveDocumentId ?? throw new InvalidOperationException("Failed to create split Word document");
                }
                else
                {
                    newDocumentId = $"split_bookmark_{bookmark.Name}_{Guid.NewGuid():N}.docx";
                    newDocumentId = CreateSplitKey(outputFilePath, newDocumentId);
                    SaveDocument(newDocumentId, newDocument);
                }
                splitDocumentIds.Add(newDocumentId);
            }
            return splitDocumentIds;
        }

        /// <summary>
        /// Splits the Word document by extracting content between paired placeholder markers into separate documents.
        /// </summary>
        /// <param name="document">The source Word document.</param>
        /// <param name="placeholderPattern">The regex pattern used to identify the placeholder markers (e.g. "{{(.*)}}"). Defaults to "&lt;&lt;(.*)&gt;&gt;".</param>
        /// <param name="outputFilePath">Output file path prefix for saving split documents (DocumentStorage mode only). If empty, generated file names are used.</param>
        private List<string> SplitByPlaceholderText(WordDocument document, string placeholderPattern = "<<(.*)>>", string? outputFilePath = null)
        {
            var splitDocumentIds = new List<string>();
            //Finds all the placeholder text in the Word document.
            TextSelection[] textSelections = document.FindAll(new Regex(placeholderPattern));
            if (textSelections == null)
                return splitDocumentIds;

            //Unique ID for each bookmark.
            int bkmkId = 1;
            //Collection to hold the inserted bookmarks.
            List<string> bookmarks = new List<string>();
            //Iterate each text selection in pairs (start / end placeholders).
            for (int i = 0; i < textSelections.Length; i++)
            {
                //Get the start placeholder as WTextRange.
                WTextRange textRange = textSelections[i].GetAsOneRange();
                //Get the index of the start placeholder text.
                WParagraph startParagraph = textRange.OwnerParagraph;
                int index = startParagraph.ChildEntities.IndexOf(textRange);
                string bookmarkName = "Bookmark_" + bkmkId;
                //Add new bookmark name to bookmarks collection.
                bookmarks.Add(bookmarkName);
                //Create bookmark start.
                BookmarkStart bkmkStart = new BookmarkStart(document, bookmarkName);
                //Insert the bookmark start before the start placeholder.
                startParagraph.ChildEntities.Insert(index, bkmkStart);
                //Remove the start placeholder text.
                textRange.Text = string.Empty;

                i++;
                //Get the end placeholder as WTextRange.
                textRange = textSelections[i].GetAsOneRange();
                //Get the index of the end placeholder text.
                WParagraph endParagraph = textRange.OwnerParagraph;
                index = endParagraph.ChildEntities.IndexOf(textRange);
                //Create bookmark end.
                BookmarkEnd bkmkEnd = new BookmarkEnd(document, bookmarkName);
                //Insert the bookmark end after the end placeholder.
                endParagraph.ChildEntities.Insert(index + 1, bkmkEnd);
                bkmkId++;
                //Remove the end placeholder text.
                textRange.Text = string.Empty;
            }

            BookmarksNavigator bookmarksNavigator = new BookmarksNavigator(document);
            foreach (string bookmark in bookmarks)
            {
                //Move the virtual cursor to the location before the end of the bookmark.
                bookmarksNavigator.MoveToBookmark(bookmark);
                //Get the bookmark content as WordDocumentPart.
                WordDocumentPart wordDocumentPart = bookmarksNavigator.GetContent();
                WordDocument newDocument = wordDocumentPart.GetAsWordDocument();
                
                string newDocumentId;
                if (Mode == DocumentManagerMode.InMemory)
                {
                   ((WordDocumentManager)InMemoryManager!).AddDocument(newDocument);
                    newDocumentId = InMemoryManager.ActiveDocumentId ?? throw new InvalidOperationException("Failed to create split Word document");
                }
                else
                {
                    newDocumentId = $"split_placeholder_{bookmark}_{Guid.NewGuid():N}.docx";
                    newDocumentId = CreateSplitKey(outputFilePath, newDocumentId);
                    SaveDocument(newDocumentId, newDocument);
                }
                splitDocumentIds.Add(newDocumentId);
            }
            return splitDocumentIds;
        }

        /// <summary>
        /// Splits the Word document at each explicit page break or paragraph with PageBreakBefore
        /// formatting into separate documents.
        /// <para>
        /// Two kinds of page breaks are detected:
        /// <list type="bullet">
        ///   <item><description>A <see cref="WParagraph"/> whose <c>ParagraphFormat.PageBreakBefore</c> is <c>true</c>.</description></item>
        ///   <item><description>A <see cref="WParagraph"/> that contains a <see cref="Break"/> child entity with <c>BreakType.PageBreak</c>.</description></item>
        /// </list>
        /// For inline page breaks the paragraph is split: content before the break is added to the
        /// current document; content after the break opens the next document.
        /// </para>
        /// </summary>
        private List<string> SplitByPageBreak(WordDocument document, string? outputFilePath = null)
        {
            var splitDocumentIds = new List<string>();

            // Create the first accumulator document.
            WordDocument newDocument;
            string currentDocId;
            if (Mode == DocumentManagerMode.InMemory)
            {
                newDocument = InMemoryManager!.CreateDocument();
                currentDocId = InMemoryManager.ActiveDocumentId ?? throw new InvalidOperationException("Failed to create split Word document");
            }
            else
            {
                newDocument = new WordDocument();
                currentDocId = $"split_pagebreak_1_{Guid.NewGuid():N}.docx";
            }

            foreach (WSection section in document.Sections)
            {
                // Ensure the current accumulator has at least one section.
                WSection newSection = section.Clone();
                newSection.Body.ChildEntities.Clear();
                newDocument.Sections.Add(newSection);

                foreach (TextBodyItem item in section.Body.ChildEntities)
                {
                    if (item is not WParagraph paragraph)
                    {
                        // Tables and other body items are added as-is.
                        AddEntity(newSection, item);
                        continue;
                    }

                    // ── Case 1: PageBreakBefore paragraph format ──────────────────────
                    // The paragraph starts a new logical page; treat it as the first
                    // paragraph of the NEXT document (mirrors SplitByHeading behavior).
                    if (paragraph.ParagraphFormat.PageBreakBefore)
                    {
                        // Save the current document
                        if (Mode == DocumentManagerMode.DocumentStorage)
                        {
                            currentDocId = CreateSplitKey(outputFilePath, currentDocId);
                            SaveDocument(currentDocId, newDocument);
                        }
                        splitDocumentIds.Add(currentDocId);

                        // Create new document
                        if (Mode == DocumentManagerMode.InMemory)
                        {
                            newDocument = InMemoryManager!.CreateDocument();
                            currentDocId = InMemoryManager.ActiveDocumentId!;
                        }
                        else
                        {
                            newDocument = new WordDocument();
                            currentDocId = $"split_pagebreak_{splitDocumentIds.Count + 1}_{Guid.NewGuid():N}.docx";
                        }
                        newSection = section.Clone();
                        newSection.Body.ChildEntities.Clear();
                        newDocument.Sections.Add(newSection);

                        // Add the paragraph (without the break-before flag side-effect)
                        // into the new document so its content is preserved.
                        AddEntity(newSection, paragraph);
                        continue;
                    }

                    // ── Case 2: Inline page break inside a paragraph ──────────────────
                    // Find the first Break child with BreakType.PageBreak.
                    int breakIndex = -1;
                    for (int i = 0; i < paragraph.ChildEntities.Count; i++)
                    {
                        if (paragraph.ChildEntities[i] is Break b &&
                            b.BreakType == BreakType.PageBreak)
                        {
                            breakIndex = i;
                            break;
                        }
                    }

                    if (breakIndex < 0)
                    {
                        // No page break — just add the paragraph to the current document.
                        AddEntity(newSection, paragraph);
                        continue;
                    }

                    // Clone the paragraph twice: once for content before the break and
                    // once for content after.
                    WParagraph beforeBreak = (WParagraph)paragraph.Clone();
                    WParagraph afterBreak  = (WParagraph)paragraph.Clone();

                    // beforeBreak keeps only items at indices 0 … breakIndex-1.
                    // Remove from the end so indices stay stable while removing.
                    for (int ci = beforeBreak.ChildEntities.Count - 1; ci >= breakIndex; ci--)
                        beforeBreak.ChildEntities.RemoveAt(ci);

                    // afterBreak keeps only items at indices breakIndex+1 … end
                    // (the Break entity itself is discarded).
                    for (int ci = breakIndex; ci >= 0; ci--)
                        afterBreak.ChildEntities.RemoveAt(ci);

                    // Add the pre-break fragment to the current accumulator.
                    if (beforeBreak.ChildEntities.Count > 0 || !string.IsNullOrEmpty(beforeBreak.Text))
                        newSection.Body.ChildEntities.Add(beforeBreak);

                    // Seal the current document and open the next one.
                    if (Mode == DocumentManagerMode.DocumentStorage)
                    {
                        currentDocId = CreateSplitKey(outputFilePath, currentDocId);
                        SaveDocument(currentDocId, newDocument);
                    }
                    splitDocumentIds.Add(currentDocId);

                    // Create new document
                    if (Mode == DocumentManagerMode.InMemory)
                    {
                        newDocument = InMemoryManager!.CreateDocument();
                        currentDocId = InMemoryManager.ActiveDocumentId!;
                    }
                    else
                    {
                        newDocument = new WordDocument();
                        currentDocId = $"split_pagebreak_{splitDocumentIds.Count + 1}_{Guid.NewGuid():N}.docx";
                    }
                    newSection = section.Clone();
                    newSection.Body.ChildEntities.Clear();
                    newDocument.Sections.Add(newSection);

                    // Add the post-break fragment to the new accumulator.
                    if (afterBreak.ChildEntities.Count > 0 || !string.IsNullOrEmpty(afterBreak.Text))
                        newSection.Body.ChildEntities.Add(afterBreak);
                }
            }

            // Save whatever remains in the last accumulator.
            if (Mode == DocumentManagerMode.DocumentStorage)
            {
                currentDocId = CreateSplitKey(outputFilePath, currentDocId);
                SaveDocument(currentDocId, newDocument);
            }
            splitDocumentIds.Add(currentDocId);

            return splitDocumentIds;
        }

        /// <summary>Clones a section from the source document, clears its content and headers/footers, and adds it to the target document.</summary>
        private static WSection AddSection(WordDocument newDocument, WSection section)
        {
            //Create new session based on original document
            WSection newSection = section.Clone();
            newSection.Body.ChildEntities.Clear();
            //Remove the first page header.
            newSection.HeadersFooters.FirstPageHeader.ChildEntities.Clear();
            //Remove the first page footer.
            newSection.HeadersFooters.FirstPageFooter.ChildEntities.Clear();
            //Remove the odd footer.
            newSection.HeadersFooters.OddFooter.ChildEntities.Clear();
            //Remove the odd header.
            newSection.HeadersFooters.OddHeader.ChildEntities.Clear();
            //Remove the even header.
            newSection.HeadersFooters.EvenHeader.ChildEntities.Clear();
            //Remove the even footer.
            newSection.HeadersFooters.EvenFooter.ChildEntities.Clear();
            //Add cloned section into new document
            newDocument.Sections.Add(newSection);
            return newSection;
        }

        /// <summary>Adds a cloned entity to the body of the specified section.</summary>
        private static void AddEntity(WSection newSection, Entity entity)
        {
            //Add cloned item into the newly created section
            newSection.Body.ChildEntities.Add(entity.Clone());
        }

        /// <summary>
        /// Creates a deep copy (clone) of a Word document and stores it in memory.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the source document to clone.</param>
        /// <param name="outputFilePath">Output file path for saving the cloned document (DocumentStorage mode only).</param>
        /// <returns>Result containing the new document ID of the cloned document.</returns>
        [Tool(
            Name = "CloneDocument",
            Description = "Creates a deep copy (clone) of an existing Word document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Returns the new document ID of the cloned document. Use this to duplicate a template before making changes, so the original is preserved.")]
        public AgentToolResult CloneDocument(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the source document to clone")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "Output file path for saving the cloned document (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                //Creates a deep copy of the source document.
                WordDocument clonedDocument = document.Clone();
                
                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_cloned.docx";

                string outputKey = outputFilePath;
                if (Mode == DocumentManagerMode.InMemory)
                {
                    ((WordDocumentManager)InMemoryManager!).AddDocument(clonedDocument);
                    outputKey = InMemoryManager.ActiveDocumentId!;
                }
                else
                {
                    SaveDocument(outputKey, clonedDocument);
                }

                return AgentToolResult.Ok(
                    $"Document cloned successfully. New document ID: {outputKey}",
                    new { SourceDocumentId = documentIdOrFilePath, ClonedDocumentId = outputKey });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to clone document: {ex.Message}");
            }
        }

        /// <summary>
        /// Updates all fields (DATE, TIME, DOCVARIABLE, IF, SEQ, etc.) present in a Word document.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document whose fields should be updated.</param>
        /// <param name="updatePageFields">When true, also updates page-related fields such as Page and NumPages (requires Word-to-PDF layout engine).</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(
            Name = "UpdateDocumentFields",
            Description = "Updates all fields in a Word document — DATE, TIME, DOCVARIABLE, DOCPROPERTY, IF, SEQ, NUMPAGES, Cross-Reference, CreateDate, and more. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Pass updatePageFields=true to also update Page and NumPages fields (requires additional PDF layout assemblies).")]
        public AgentToolResult UpdateDocumentFields(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document whose fields should be updated")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "When true, also updates page-related fields such as Page and NumPages. Defaults to false.")]
            bool updatePageFields = false,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                //Updates all fields present in the document.
                document.UpdateDocumentFields(updatePageFields);

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_fields_updated.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Document fields updated successfully for {outputKey}",
                    new { DocumentId = outputKey, UpdatePageFields = updatePageFields });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to update document fields: {ex.Message}");
            }
        }

        /// <summary>
        /// Unlinks all fields in a Word document, replacing each field with its most recent result text.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document whose fields should be unlinked.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success with the count of fields unlinked.</returns>
        [Tool(
            Name = "UnlinkDocumentFields",
            Description = "Unlinks all fields in a Word document by replacing each field with its current result as static text or a graphic. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). After unlinking, field values can no longer be updated automatically. XE (Index Entry) fields are skipped as they cannot be unlinked.")]
        public AgentToolResult UnlinkDocumentFields(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document whose fields should be unlinked")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                int unlinkedCount = UnlinkAllFields(document);

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_fields_unlinked.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Successfully unlinked {unlinkedCount} field(s) in {outputKey}",
                    new { DocumentId = outputKey, UnlinkedFieldCount = unlinkedCount });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to unlink document fields: {ex.Message}");
            }
        }

        /// <summary>
        /// Uses FindAllItemsByProperty to collect every WField in the document, then unlinks each one.
        /// </summary>
        private static int UnlinkAllFields(WordDocument document)
        {
            // Find all fields in the document regardless of field type (pass null for property name/value).
            List<Entity> fieldEntities = document.FindAllItemsByProperty(EntityType.Field, null, null);
            int unlinkedCount = 0;
            foreach (Entity entity in fieldEntities)
            {
                if (entity is WField field)
                {
                    try
                    {
                        field.Unlink();
                        unlinkedCount++;
                    }
                    catch { /* Skip fields that cannot be unlinked (e.g., XE index-entry fields) */ }
                }
            }
            return unlinkedCount;
        }

        /// <summary>
        /// Updates (rebuilds) the Table of Contents in a Word document.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document containing the TOC to update.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(
            Name = "UpdateTableOfContents",
            Description = "Updates (rebuilds) the Table of Contents (TOC) in a Word document based on the current heading styles and page layout. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). Call this after modifying document content to reflect the latest headings and page numbers in the TOC.")]
        public AgentToolResult UpdateTableOfContents(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document whose Table of Contents should be updated")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                //Updates the table of contents in the document.
                document.UpdateTableOfContents();

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_toc_updated.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok(
                    $"Table of contents updated successfully for {outputKey}",
                    new { DocumentId = outputKey });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to update table of contents: {ex.Message}");
            }
        }

        /// <summary>
        /// Compares two documents and marks differences as tracked changes.
        /// </summary>
        /// <param name="originalDocumentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the original document.</param>
        /// <param name="revisedDocumentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the revised document.</param>
        /// <param name="author">The author name associated with the tracked revisions.</param>
        /// <param name="dateTime">The date and time associated with the tracked revisions.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating whether the document comparison was completed successfully.</returns>
        [Tool(
            Name = "CompareDocuments",
            Description = "Compares the original document with the revised document and marks differences as tracked changes in the original document. originalDocumentIdOrFilePath and revisedDocumentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult CompareDocuments(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the original document")]
            string originalDocumentIdOrFilePath,
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the revised document")]
            string revisedDocumentIdOrFilePath,
            [ToolParameter(Description = "The revision author name")]
            string author = "Author",
            [ToolParameter(Description = "The date and time to associate with the revisions")]
            DateTime dateTime = default(DateTime),
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var originalDoc = OpenDocument(originalDocumentIdOrFilePath);
                var revisedDoc = OpenDocument(revisedDocumentIdOrFilePath);

                if (originalDoc == null || revisedDoc == null)
                    return AgentToolResult.Fail("One or both documents not found");

                originalDoc.Compare(revisedDoc, author, dateTime);

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_compared.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, originalDoc);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = originalDocumentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Documents compared successfully. Changes tracked in {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to compare documents: {ex.Message}");
            }
        }


    }
}
      