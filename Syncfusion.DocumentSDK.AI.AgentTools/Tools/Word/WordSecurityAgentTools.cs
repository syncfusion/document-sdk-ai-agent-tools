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
using System;

namespace Syncfusion.AI.AgentTools.Word
{
    /// <summary>
    /// Provides agent tools for document security, protection, and encryption.
    /// Handles password protection, encryption, and document access control.
    /// </summary>
    public class WordSecurityAgentTools : AgentToolBase<WordDocument>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WordSecurityAgentTools"/> class (Mode 1 — InMemory).
        /// </summary>
        /// <param name="manager">The document manager for managing Word documents.</param>
        public WordSecurityAgentTools(WordDocumentManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="WordSecurityAgentTools"/> class (Mode 2 — DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public WordSecurityAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.Word) { }

        /// <summary>
        /// Protects the Word document with a password and protection type.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document to protect.</param>
        /// <param name="password">The password used to apply document protection.</param>
        /// <param name="protectionType">The type of protection to apply, such as allowing comments, form fields, revisions, read-only access, or no protection.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating whether the document protection was applied successfully.</returns>
        [Tool(
            Name = "ProtectDocument",
            Description = "Protects the Word document with password and protection type. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult ProtectDocument(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document to protect")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The password used to protect the document")]
            string password,
            [ToolParameter(Description = "The type of protection to apply. Accepted values: 'AllowOnlyComments' (also: 'comments') - allows only adding or modifying comments; 'AllowOnlyFormFields' (also: 'formfields', 'form fields') - allows only modifying form field values; 'AllowOnlyRevisions' (also: 'revisions', 'trackchanges', 'track changes') - allows only tracked changes (accept/reject options are disabled in Word); 'AllowOnlyReading' (also: 'readonly', 'read only', 'read-only', 'reading') - read-only access, no editing allowed; 'NoProtection' (also: 'none', 'noprotection') - removes protection and allows full editing access.")]
            string protectionType,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                ProtectionType type = protectionType.ToLowerInvariant().Replace(" ", "").Replace("-", "") switch
                {
                    "allowonlycomments"   or "comments"                          => ProtectionType.AllowOnlyComments,
                    "allowonlyformfields" or "formfields"                        => ProtectionType.AllowOnlyFormFields,
                    "allowonlyrevisions"  or "revisions" or "trackchanges"       => ProtectionType.AllowOnlyRevisions,
                    "allowonlyreading"    or "readonly"  or "reading"            => ProtectionType.AllowOnlyReading,
                    "noprotection"        or "none"                              => ProtectionType.NoProtection,
                    _                                                            => ProtectionType.NoProtection
                };

                document.Protect(type, password);

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_protected.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Document {outputKey} protected successfully with {type}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to protect document: {ex.Message}");
            }
        }

        /// <summary>
        /// Encrypts the document with a password.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document to encrypt.</param>
        /// <param name="password">The password used to encrypt the document.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating whether the document was encrypted successfully.</returns>
        [Tool(
            Name = "EncryptDocument",
            Description = "Encrypts the document using the provided password. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult EncryptDocument(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document to encrypt")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The password used to encrypt the document")]
            string password,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                document.EncryptDocument(password);

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_encrypted.docx";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Document {outputKey} encrypted successfully");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to encrypt document: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes protection from a document using the password.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document to unprotect.</param>
        /// <param name="password">The password used to remove document protection.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating whether the document was unprotected successfully.</returns>
        [Tool(
            Name = "UnprotectDocument",
            Description = "Removes protection from the Word document using the provided password. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult UnprotectDocument(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document to unprotect")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The password used to unprotect the document")]
            string password,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                // Unprotect by setting NoProtection
                document.Protect(ProtectionType.NoProtection, password);

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_unprotected.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Document {outputKey} unprotected successfully");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to unprotect document: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes encryption from a document.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document to decrypt.</param>
        /// <param name="password">The password used to decrypt the document.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result indicating whether the document was decrypted successfully.</returns>
        [Tool(
            Name = "DecryptDocument",
            Description = "Removes encryption from the Word document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). The document must be loaded with the correct password first.")]
        public AgentToolResult DecryptDocument(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode) of the document to decrypt")]
            string documentIdOrFilePath,
            [ToolParameter(Description = "The protection password")] string? password,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
            string? outputFilePath = null)
        {
            try
            {
                var document = OpenDocument(documentIdOrFilePath, password);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                // Remove encryption by setting empty password
                document.RemoveEncryption();

                // ── Save ────────────────────────────────────────────────────────
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_decrypted.docx";
                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath; // InMemory mode always updates the same document ID

                return AgentToolResult.Ok($"Document {outputKey} decrypted successfully");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to decrypt document: {ex.Message}");
            }
        }
    }
}
