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
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.Presentation;

namespace Syncfusion.AI.AgentTools.PowerPoint
{
    /// <summary>
    /// Provides AI agent tools for PowerPoint presentation security operations.
    /// Handles encryption, decryption, and write protection.
    /// </summary>
    public class PresentationSecurityAgentTools : AgentToolBase<IPresentation>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationSecurityAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The presentation manager for managing PowerPoint presentations.</param>
        public PresentationSecurityAgentTools(PresentationManager manager)
            : base(manager, DocumentType.PowerPoint) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationSecurityAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public PresentationSecurityAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.PowerPoint) { }

        /// <summary>
        /// Write protects the PowerPoint presentation with a password.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or the input presentation file path (DocumentStorage mode).</param>
        /// <param name="password">The password to protect with.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "ProtectPresentation", Description = "Write Protect the PowerPoint presentation. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ProtectPresentation(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath,
            [ToolParameter(Description = "The password to protect with")] string password,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);
                ArgumentNullException.ThrowIfNull(password);

                var presentation = OpenDocument(documentIdOrFilePath);
                if (presentation == null)
                    return AgentToolResult.Fail($"Presentation not found: {documentIdOrFilePath}");

                // Set write protection
                presentation.SetWriteProtection(password);

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_protected.pptx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, presentation);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"PowerPoint presentation {outputKey} write protected successfully");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to protect PowerPoint presentation: {ex.Message}");
            }
        }

        /// <summary>
        /// Encrypts the presentation with a password.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or the input presentation file path (DocumentStorage mode).</param>
        /// <param name="password">The password to encrypt with.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "EncryptPresentation", Description = "Encrypts the presentation using the provided password. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult EncryptPresentation(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath,
            [ToolParameter(Description = "The password to encrypt with")] string password,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);
                ArgumentNullException.ThrowIfNull(password);

                var presentation = OpenDocument(documentIdOrFilePath);
                if (presentation == null)
                    return AgentToolResult.Fail($"Presentation not found: {documentIdOrFilePath}");

                // Encrypt the presentation
                presentation.Encrypt(password);

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_encrypted.pptx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, presentation);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"PowerPoint presentation {outputKey} encrypted successfully");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to encrypt PowerPoint presentation: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes write protection from a presentation.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or the input presentation file path (DocumentStorage mode).</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "UnprotectPresentation", Description = "Removes write protection from a presentation. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult UnprotectPresentation(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);

                var presentation = OpenDocument(documentIdOrFilePath);
                if (presentation == null)
                    return AgentToolResult.Fail($"Presentation not found: {documentIdOrFilePath}");

                // Remove write protection (no parameter needed)
                presentation.RemoveWriteProtection();

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_unprotected.pptx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, presentation);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"PowerPoint presentation {outputKey} write protection removed successfully");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to unprotect PowerPoint presentation: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes encryption from a presentation.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or the input presentation file path (DocumentStorage mode).</param>
        /// <param name="password">The password used to decrypt the presentation.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating success or failure.</returns>
        [Tool(Name = "DecryptPresentation", Description = "Removes encryption from a presentation. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult DecryptPresentation(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath,
            [ToolParameter(Description = "The protection password")] string password,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);

                var presentation = OpenDocument(documentIdOrFilePath, password);
                if (presentation == null)
                    return AgentToolResult.Fail($"Presentation not found: {documentIdOrFilePath}");

                // Remove encryption
                presentation.RemoveEncryption();

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_decrypted.pptx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, presentation);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"PowerPoint presentation {outputKey} encryption removed successfully");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to decrypt PowerPoint presentation: {ex.Message}");
            }
        }
    }
}
