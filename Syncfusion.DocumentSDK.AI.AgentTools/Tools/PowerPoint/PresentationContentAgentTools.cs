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
using System.Text;
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.Presentation;

namespace Syncfusion.AI.AgentTools.PowerPoint
{
    /// <summary>
    /// Provides AI agent tools for PowerPoint presentation content operations.
    /// Handles text extraction and slide counting.
    /// </summary>
    public class PresentationContentAgentTools : AgentToolBase<IPresentation>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationContentAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The presentation manager for managing PowerPoint presentations.</param>
        public PresentationContentAgentTools(PresentationManager manager)
            : base(manager, DocumentType.PowerPoint) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationContentAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public PresentationContentAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.PowerPoint) { }

        /// <summary>
        /// Gets the PowerPoint presentation content as text.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <returns>Result containing the text content of the presentation.</returns>
        [Tool(Name = "GetText", Description = "Gets the PowerPoint presentation content as text. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult GetText(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath)
        {
            try
            {
                var presentation = OpenDocument(documentIdOrFilePath);
                if (presentation == null)
                    return AgentToolResult.Fail($"Presentation not found: {documentIdOrFilePath}");

                StringBuilder textBuilder = new StringBuilder();

                // Extract text from all slides
                for (int i = 0; i < presentation.Slides.Count; i++)
                {
                    ISlide slide = presentation.Slides[i];
                    textBuilder.AppendLine($"--- Slide {i + 1} ---");

                    // Extract text from all shapes in the slide
                    ExtractText(slide.Shapes as IShapes, textBuilder);

                    // Extract text from the slide notes body
                    if (slide.NotesSlide?.NotesTextBody != null)
                    {
                        foreach (IParagraph paragraph in slide.NotesSlide.NotesTextBody.Paragraphs)
                        {
                            textBuilder.AppendLine(paragraph.Text);
                        }
                    }

                    // Extract text from the slide notes shapes
                    if (slide.NotesSlide?.Shapes != null)
                    {
                        ExtractText(slide.NotesSlide.Shapes as IShapes, textBuilder);
                    }

                    // Extract text from the layout slide shapes
                    if (slide.LayoutSlide?.Shapes != null)
                    {
                        ExtractText(slide.LayoutSlide.Shapes as IShapes, textBuilder, true);

                        // Extract text from the master slide shapes
                        if (slide.LayoutSlide.MasterSlide?.Shapes != null)
                        {
                            ExtractText(slide.LayoutSlide.MasterSlide.Shapes as IShapes, textBuilder, true);
                        }
                    }

                    textBuilder.AppendLine();
                }

                string extractedText = textBuilder.ToString();
                int slideCount = presentation.Slides.Count;

                if (Mode == DocumentManagerMode.DocumentStorage)
                    presentation.Close();

                return AgentToolResult.Ok(
                    $"Successfully extracted text from {slideCount} slide(s)",
                    new { Text = extractedText, SlideCount = slideCount });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to extract text from PowerPoint presentation: {ex.Message}");
            }
        }

        /// <summary>
        /// Returns the number of slides in the presentation.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <returns>Result containing the number of slides.</returns>
        [Tool(Name = "GetSlideCount", Description = "Returns the number of slides. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult GetSlideCount(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);

                var presentation = OpenDocument(documentIdOrFilePath);
                if (presentation == null)
                    return AgentToolResult.Fail($"Presentation not found: {documentIdOrFilePath}");

                int slideCount = presentation.Slides.Count;

                if (Mode == DocumentManagerMode.DocumentStorage)
                    presentation.Close();

                return AgentToolResult.Ok(
                    $"Presentation has {slideCount} slide(s)",
                    new { SlideCount = slideCount });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to get slide count: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts text from all shapes in the given shape collection, handling tables, group shapes, and regular shapes.
        /// </summary>
        /// <param name="shapes">The shape collection to extract text from.</param>
        /// <param name="textBuilder">The StringBuilder to append extracted text to.</param>
        /// <param name="ignorePlaceHolder">Whether to ignore placeholder shapes.</param>
        private static void ExtractText(IShapes shapes, StringBuilder textBuilder, bool ignorePlaceHolder = false)
        {
            foreach (IShape shape in shapes)
            {
                if (shape is ITable)
                    ExtractTextInTable(shape, textBuilder);
                else if (shape is ISmartArt)
                    ExtractTextInSmartArt(shape, textBuilder);
                else if (shape is IGroupShape)
                    ExtractText((shape as IGroupShape).Shapes, textBuilder, ignorePlaceHolder);
                else
                    ExtractTextInShape(shape, textBuilder, ignorePlaceHolder);
            }
        }

        /// <summary>
        /// Extracts text from a regular shape's text body.
        /// </summary>
        /// <param name="shape">The shape to extract text from.</param>
        /// <param name="textBuilder">The StringBuilder to append extracted text to.</param>
        /// <param name="ignorePlaceHolder">Whether to ignore placeholder shapes.</param>
        private static void ExtractTextInShape(IShape shape, StringBuilder textBuilder, bool ignorePlaceHolder)
        {
            if (shape.TextBody == null || (ignorePlaceHolder && (shape as ISlideItem).SlideItemType == SlideItemType.Placeholder))
                return;

            foreach (IParagraph paragraph in shape.TextBody.Paragraphs)
            {
                textBuilder.AppendLine(paragraph.Text);
            }
        }

        /// <summary>
        /// Extracts text from a table shape by iterating all rows and cells.
        /// </summary>
        /// <param name="shape">The table shape to extract text from.</param>
        /// <param name="textBuilder">The StringBuilder to append extracted text to.</param>
        private static void ExtractTextInTable(IShape shape, StringBuilder textBuilder)
        {
            ITable table = shape as ITable;
            if (table == null)
                return;

            foreach (IRow row in table.Rows)
            {
                foreach (ICell cell in row.Cells)
                {
                    textBuilder.AppendLine(cell.TextBody.Text);
                }
            }
        }

        /// <summary>
        /// Extracts text from a SmartArt shape by iterating all nodes and their child nodes recursively.
        /// </summary>
        /// <param name="shape">The SmartArt shape to extract text from.</param>
        /// <param name="textBuilder">The StringBuilder to append extracted text to.</param>
        private static void ExtractTextInSmartArt(IShape shape, StringBuilder textBuilder)
        {
            ISmartArt smartArt = shape as ISmartArt;
            if (smartArt == null)
                return;

            foreach (ISmartArtNode node in smartArt.Nodes)
            {
                ExtractTextInSmartArtNode(node, textBuilder);
            }
        }

        /// <summary>
        /// Extracts text from a SmartArt node and recursively from its child nodes.
        /// </summary>
        /// <param name="node">The SmartArt node to extract text from.</param>
        /// <param name="textBuilder">The StringBuilder to append extracted text to.</param>
        private static void ExtractTextInSmartArtNode(ISmartArtNode node, StringBuilder textBuilder)
        {
            if (node.TextBody != null)
            {
                foreach (IParagraph paragraph in node.TextBody.Paragraphs)
                {
                    textBuilder.AppendLine(paragraph.Text);
                }
            }

            // Recursively extract text from child nodes
            foreach (ISmartArtNode childNode in node.ChildNodes)
            {
                ExtractTextInSmartArtNode(childNode, textBuilder);
            }
        }
    }
}
