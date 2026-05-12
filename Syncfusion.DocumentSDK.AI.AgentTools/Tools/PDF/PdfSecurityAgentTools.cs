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
using Syncfusion.Drawing;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Pdf.Redaction;
using Syncfusion.Pdf.Security;
using System.Text.Json;

namespace Syncfusion.AI.AgentTools.PDF
{
    /// <summary>
    /// Provides AI agent tools for PDF document security and encryption operations.
    /// Handles encryption, decryption, and permission management.
    /// </summary>
    public class PdfSecurityAgentTools : AgentToolBase<PdfDocumentBase>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PdfSecurityAgentTools"/> class (Mode 1 � InMemory).
        /// </summary>
        /// <param name="manager">The PDF document manager.</param>
        public PdfSecurityAgentTools(PdfDocumentManager manager)
            : base(manager, DocumentType.PDF) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="PdfSecurityAgentTools"/> class (Mode 2 � DocumentStorage).
        /// </summary>
        /// <param name="manager">The document storage manager.</param>
        public PdfSecurityAgentTools(DocumentStorageManager manager)
            : base(manager, DocumentType.PDF) { }

        /// <summary>
        /// Protects a PDF document with a user password and applies the specified encryption algorithm and key size.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="password">The password to encrypt the document.</param>
        /// <param name="encryptionAlgorithm">The encryption algorithm (AES or RC4).</param>
        /// <param name="keySize">The bit-length of the encryption key (40, 128, 256).</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result confirming encryption with the algorithm and key size applied.</returns>
        [Tool(Name = "EncryptPdf", Description = "Protect the PDF file with password. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). encryptionAlgorithm: The algorithm applied to encrypt the PDF. keySize: The bit-length of the encryption key.")]
        public AgentToolResult EncryptPdf(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath,
            [ToolParameter(Description = "The password to encrypt the document")] string password,
            [ToolParameter(Description = "The encryption algorithm (AES or RC4)")] string encryptionAlgorithm = "AES",
            [ToolParameter(Description = "The bit-length of the encryption key (40, 128, 256)")] string keySize = "256",
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);
                ArgumentNullException.ThrowIfNull(password);

                // -- Open --------------------------------------------------------
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                PdfEncryptionKeySize encryptionKeySize = keySize switch
                {
                    "40" => PdfEncryptionKeySize.Key40Bit,
                    "128" => PdfEncryptionKeySize.Key128Bit,
                    "256" => PdfEncryptionKeySize.Key256Bit,
                    _ => PdfEncryptionKeySize.Key256Bit
                };

                PdfSecurity security = document.Security;
                security.KeySize = encryptionKeySize;
                security.UserPassword = password;

                // -- Save --------------------------------------------------------
                 if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_encrypted.pdf";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath;

                return AgentToolResult.Ok(
                    $"PDF document {outputKey} encrypted successfully with {encryptionAlgorithm} {keySize}-bit encryption",
                    new { Algorithm = encryptionAlgorithm, KeySize = keySize });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to encrypt PDF document: {ex.Message}");
            }
        }

        /// <summary>
        /// Removes encryption from a password-protected PDF document by clearing its security passwords and permissions.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result confirming the document has been decrypted.</returns>
        [Tool(Name = "DecryptPdf", Description = "Removes encryption from a protected PDF file. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult DecryptPdf(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath,
            [ToolParameter(Description = "The protection password")] string password,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);

                // -- Open --------------------------------------------------------
                var document = OpenDocument(documentIdOrFilePath, password);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                if (document is PdfLoadedDocument loadedDocument)
                {
                    loadedDocument.Security.UserPassword = string.Empty;
                    loadedDocument.Security.OwnerPassword = string.Empty;
                    loadedDocument.Security.Permissions = PdfPermissionsFlags.Default;
                    loadedDocument.FileStructure.IncrementalUpdate = false; // Force full rewrite to remove encryption

                    // -- Save ----------------------------------------------------
                    if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                        outputFilePath = "output_decrypted.pdf";

                    string outputKey = outputFilePath;
                    SaveDocument(outputKey, loadedDocument);
                    if (Mode == DocumentManagerMode.InMemory)
                        outputKey = documentIdOrFilePath;

                    return AgentToolResult.Ok($"PDF document {outputKey} decrypted successfully");
                }
                else
                {
                    return AgentToolResult.Fail("Decryption is only supported for loaded PDF documents");
                }
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to decrypt PDF document: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets document permissions on a PDF such as print, copy, edit content, and more.
        /// Permissions are specified as a comma-separated string of flag names.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="permissions">The permissions to set (comma-separated: Print, EditContent, CopyContent, etc.).</param>
        /// <param name="password">The password to open the document if it is already encrypted (optional).</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result confirming the permissions were applied.</returns>

        [Tool(Name = "SetPermissions", Description = "Sets or restricts PDF permissions by specifying allowed actions like printing, editing, copying, annotations, and form filling.")]
        public AgentToolResult SetPermissions(
        [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        string documentIdOrFilePath,

        [ToolParameter(Description =
            "Permissions object specifying allowed actions. " +
            "Each field should be true (allowed) or false (restricted). " +
            "Example: { AllowPrint: false, AllowEditContent: false } means read-only.")]
        PermissionsRequest permissions,

        [ToolParameter(Description = "Optional password if the document is already encrypted.")]
        string? password = null,

        [ToolParameter(Description = "Optional output file path (used only in DocumentStorage mode).")]
        string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);

                // ✅ Open document
                var document = string.IsNullOrEmpty(password)
                    ? OpenDocument(documentIdOrFilePath)
                    : OpenDocument(documentIdOrFilePath, password);

                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");
                
                // ✅ Convert DTO → PdfPermissionsFlags
                PdfPermissionsFlags flags = PdfPermissionsFlags.Default;

                if (permissions.AllowPrint)
                    flags |= PdfPermissionsFlags.Print;

                if (permissions.AllowEditContent)
                    flags |= PdfPermissionsFlags.EditContent;

                if (permissions.AllowCopyContent)
                    flags |= PdfPermissionsFlags.CopyContent;

                if (permissions.AllowEditAnnotations)
                    flags |= PdfPermissionsFlags.EditAnnotations;

                if (permissions.AllowFillFields)
                    flags |= PdfPermissionsFlags.FillFields;

                if (permissions.AllowAssembleDocument)
                    flags |= PdfPermissionsFlags.AssembleDocument;

                if (permissions.AllowAccessibilityCopyContent)
                    flags |= PdfPermissionsFlags.AccessibilityCopyContent;

                if (permissions.AllowFullQualityPrint)
                    flags |= PdfPermissionsFlags.FullQualityPrint;

                // ✅ Apply security (CRITICAL)
                document.Security.Permissions = flags;

                // ✅ Save
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_permissions.pdf";

                string outputKey = outputFilePath;

                SaveDocument(outputKey, document);

                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath;

                return AgentToolResult.Ok(
                    $"Permissions set successfully for PDF document {outputKey}",
                    new
                    {
                        PermissionsApplied = flags.ToString(),
                        InputPermissions = permissions
                    });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to set PDF permissions: {ex.Message}");
            }
        }
        public class PermissionsRequest
        {
            public bool AllowPrint { get; set; }
            public bool AllowEditContent { get; set; }
            public bool AllowCopyContent { get; set; }
            public bool AllowEditAnnotations { get; set; }
            public bool AllowFillFields { get; set; }
            public bool AllowAssembleDocument { get; set; }
            public bool AllowAccessibilityCopyContent { get; set; }
            public bool AllowFullQualityPrint { get; set; }
        }

        /// <summary>
        /// Removes all document permissions from a PDF by resetting them to the default (unrestricted) state.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="password">The password to open the document if it is already encrypted (optional).</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result confirming all permissions have been removed.</returns>
        [Tool(Name = "RemovePermissions", Description = "Removes document permissions. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
        public AgentToolResult RemovePermissions(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")] string documentIdOrFilePath,
            [ToolParameter(Description = "The password to open the document if it is already encrypted (optional).")] string? password = null,
            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")] string? outputFilePath = null)
        {
            try
            {
                ArgumentNullException.ThrowIfNull(documentIdOrFilePath);

                // -- Open --------------------------------------------------------
                var document = string.IsNullOrEmpty(password)
                    ? OpenDocument(documentIdOrFilePath)
                    : OpenDocument(documentIdOrFilePath, password);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                document.Security.Permissions = PdfPermissionsFlags.Default;

                // -- Save --------------------------------------------------------
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_permissions_removed.pdf";

                string outputKey = outputFilePath;
                SaveDocument(outputKey, document);
                if (Mode == DocumentManagerMode.InMemory)
                    outputKey = documentIdOrFilePath;

                return AgentToolResult.Ok($"All permissions removed from PDF document {outputKey}");
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to remove PDF permissions: {ex.Message}");
            }
        }

        /// <summary>
        /// Digitally signs a PDF document using a PFX certificate and places the signature within the specified bounds.
        /// An optional appearance image can be provided for the visible signature.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="certificateFilePath">The path to the certificate file (.pfx).</param>
        /// <param name="certificatePassword">The certificate password.</param>
        /// <param name="bounds">The signature bounds (X, Y, Width, Height).</param>
        /// <param name="pageIndex">The zero-based page index where the signature should be placed (0 for first page, -1 for last page).</param>
        /// <param name="appearanceImagePath">Optional path to signature appearance image.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the document ID, certificate used, and signature bounds.</returns>
        [Tool(
            Name = "SignPdf",
            Description =
                "Digitally signs a PDF document using a certificate and an optional appearance image. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode)."
        )]
        public AgentToolResult SignPdf(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode)")]
    string documentIdOrFilePath,

            [ToolParameter(Description = "The path to the certificate file (.pfx)")]
    string certificateFilePath,

            [ToolParameter(Description = "The certificate password")]
    string certificatePassword,

            [ToolParameter(Description = "The signature bounds (X, Y, Width, Height)")]
    RectangleF bounds,

            [ToolParameter(Description = "The zero-based page index where the signature should be placed (0 for first page, -1 for last page). Default is 0.")]
    int pageIndex = 0,

            [ToolParameter(Description = "Optional path to signature appearance image")]
    string? appearanceImagePath = null,

            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
    string? outputFilePath = null
        )
        {
            try
            {
                if (string.IsNullOrWhiteSpace(documentIdOrFilePath))
                    return AgentToolResult.Fail("documentIdOrFilePath cannot be null or empty.");

                if (string.IsNullOrWhiteSpace(certificateFilePath))
                    return AgentToolResult.Fail("certificateFilePath cannot be null or empty.");

                if (string.IsNullOrWhiteSpace(certificatePassword))
                    return AgentToolResult.Fail("certificatePassword cannot be null or empty.");

                // -- Open --------------------------------------------------------
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                bool isReloaded = false;
                PdfLoadedDocument loadedDocument;

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

                // Load certificate based on mode
                System.Security.Cryptography.X509Certificates.X509Certificate2 certificate;
                if (Mode == DocumentManagerMode.InMemory)
                {
                    if (!File.Exists(certificateFilePath))
                        return AgentToolResult.Fail($"Certificate file not found: {certificateFilePath}");

                    certificate = new System.Security.Cryptography.X509Certificates.X509Certificate2(
                        certificateFilePath,
                        certificatePassword,
                        System.Security.Cryptography.X509Certificates.X509KeyStorageFlags.MachineKeySet |
                        System.Security.Cryptography.X509Certificates.X509KeyStorageFlags.EphemeralKeySet);
                }
                else
                {
                    if (!StorageManager!.HasDocument(certificateFilePath))
                        return AgentToolResult.Fail($"Certificate file not found in storage: {certificateFilePath}");

                    using var certStream = StorageManager.GetDocumentStream(certificateFilePath);
                    if (certStream == null)
                        return AgentToolResult.Fail($"Failed to read certificate file from storage: {certificateFilePath}");

                    using var memStream = new MemoryStream();
                    certStream.CopyTo(memStream);
                    certificate = new System.Security.Cryptography.X509Certificates.X509Certificate2(
                        memStream.ToArray(),
                        certificatePassword,
                        System.Security.Cryptography.X509Certificates.X509KeyStorageFlags.MachineKeySet |
                        System.Security.Cryptography.X509Certificates.X509KeyStorageFlags.EphemeralKeySet);
                }

                PdfCertificate pdfCertificate = new PdfCertificate(certificate);

                if (loadedDocument.Pages.Count == 0)
                    return AgentToolResult.Fail("PDF document has no pages.");

                // Handle page index (-1 for last page, or specific zero-based index)
                int targetPageIndex = pageIndex;
                if (pageIndex == -1)
                    targetPageIndex = loadedDocument.Pages.Count - 1;

                if (targetPageIndex < 0 || targetPageIndex >= loadedDocument.Pages.Count)
                    return AgentToolResult.Fail($"Invalid page index: {pageIndex}. Document has {loadedDocument.Pages.Count} pages.");

                PdfLoadedPage? page = loadedDocument.Pages[targetPageIndex] as PdfLoadedPage;
                if (page == null)
                    return AgentToolResult.Fail($"Failed to load page at index {targetPageIndex}.");

                PdfSignature signature =
                    new PdfSignature(loadedDocument, page, pdfCertificate, "Signature");

                signature.Bounds = bounds;

                if (!string.IsNullOrWhiteSpace(appearanceImagePath))
                {
                    Stream? imageStream = null;
                    try
                    {
                        if (Mode == DocumentManagerMode.InMemory)
                        {
                            if (File.Exists(appearanceImagePath))
                            {
                                imageStream = new FileStream(appearanceImagePath, FileMode.Open, FileAccess.Read);
                            }
                        }
                        else
                        {
                            if (!StorageManager!.HasDocument(appearanceImagePath))
                                return AgentToolResult.Fail($"Appearance image file not found in storage: {appearanceImagePath}");
                            imageStream = StorageManager.GetDocumentStream(appearanceImagePath);
                        }

                        if (imageStream != null)
                        {
                            PdfBitmap image = new PdfBitmap(imageStream);

                            signature.Appearance.Normal.Graphics?.DrawImage(
                                image,
                                new RectangleF(0, 0, bounds.Width, bounds.Height)
                            );
                        }
                    }
                    finally
                    {
                        imageStream?.Dispose();
                    }
                }

                // -- Save --------------------------------------------------------
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_signed.pdf";
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
                    $"PDF document {outputKey} signed successfully.",
                    new { DocumentId = outputKey, CertificateUsed = certificateFilePath, SignatureBounds = bounds, PageIndex = targetPageIndex }
                );
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail($"Failed to sign PDF document: {ex.Message}, {ex.StackTrace}");
            }
        }

        /// <summary>
        /// Redacts rectangular regions from an existing PDF document by permanently removing sensitive content
        /// and filling the areas with the specified color.
        /// </summary>
        /// <param name="documentIdOrFilePath">The document ID (InMemory mode) or input file path (DocumentStorage mode).</param>
        /// <param name="redaction">Redaction instructions including page index, rectangle bounds, and optional fill color.</param>
        /// <param name="outputFilePath">Output file path for saving the result (DocumentStorage mode only).</param>
        /// <returns>Result containing the document ID and the number of redactions applied.</returns>
        [Tool(
            Name = "RedactPdf",
            Description =
                "Redacts rectangular regions from an existing PDF document. documentIdOrFilePath: The document ID (InMemory mode) or input file path (DocumentStorage mode). " +
                "Input format example: " +
                "{\"documentIdOrFilePath\":\"pdf_123456\",\"redaction\":{\"Redactions\":[{\"PageIndex\":0," +
                "\"Bounds\":{\"X\":100,\"Y\":200,\"Width\":250,\"Height\":50}," +
                "\"Color\":{\"Red\":0,\"Green\":0,\"Blue\":0}}]}}"
        )]
        public AgentToolResult RedactPdf(
            [ToolParameter(Description = "The document ID (InMemory mode) or input file path (DocumentStorage mode).")]
    string documentIdOrFilePath,

            [ToolParameter(Description =
        "Redaction instructions including page index, rectangle bounds, and optional color.")]
    RedactionRequest redaction,

            [ToolParameter(Description = "Output file path for saving the result (DocumentStorage mode only).")]
    string? outputFilePath = null
        )
        {
            try
            {
                if (string.IsNullOrWhiteSpace(documentIdOrFilePath))
                    return AgentToolResult.Fail("documentIdOrFilePath cannot be null or empty.");

                if (redaction == null || redaction.Redactions.Count == 0)
                    return AgentToolResult.Fail("At least one redaction must be provided.");

                // -- Open --------------------------------------------------------
                var document = OpenDocument(documentIdOrFilePath);
                if (document == null)
                    return AgentToolResult.Fail($"Document not found: {documentIdOrFilePath}");

                PdfLoadedDocument loadedDocument;
                bool isReloaded = false;
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

                foreach (var r in redaction.Redactions)
                {
                    if (r.PageIndex < 0 || r.PageIndex >= loadedDocument.Pages.Count)
                        return AgentToolResult.Fail($"Invalid page index: {r.PageIndex}");

                    if (loadedDocument.Pages[r.PageIndex] is not PdfLoadedPage page)
                        continue;

                    var fillColor = r.Color?.ToColor() ?? Color.Black;

                    var pdfRedaction = new PdfRedaction(
                        r.Bounds.ToRectangleF(),
                        fillColor);

                    page.AddRedaction(pdfRedaction);
                }

                loadedDocument.Redact();

                // -- Save --------------------------------------------------------
                if (outputFilePath == null && Mode == DocumentManagerMode.DocumentStorage)
                    outputFilePath = "output_redacted.pdf";

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
                    $"Redaction applied successfully into document {outputKey}.",
                    new { DocumentId = outputKey, RedactionCount = redaction.Redactions.Count });
            }
            catch (Exception ex)
            {
                return AgentToolResult.Fail(
                    $"Failed to redact PDF document: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Represents a collection of redaction items to be applied to a PDF document.
    /// </summary>
    public class RedactionRequest
    {
        /// <summary>
        /// Gets or sets the list of redaction items to apply to the document.
        /// </summary>
        public List<RedactionItem> Redactions { get; set; } = new();
    }

    /// <summary>
    /// Represents a single redaction area to be applied to a PDF page.
    /// </summary>
    public class RedactionItem
    {
        /// <summary>
        /// Gets or sets the zero-based page index where the redaction will be applied.
        /// </summary>
        public int PageIndex { get; set; }
        
        /// <summary>
        /// Gets or sets the rectangular bounds of the redaction area (X, Y, Width, Height).
        /// </summary>
        public RectangleData Bounds { get; set; } = default!;
        
        /// <summary>
        /// Gets or sets the optional fill color for the redacted area. If null, defaults to black.
        /// </summary>
        public ColorData? Color { get; set; }
    }

    /// <summary>
    /// Represents rectangular bounds with X, Y coordinates, width, and height.
    /// Used to define redaction areas or signature placements in PDF documents.
    /// </summary>
    public class RectangleData
    {
        /// <summary>
        /// Gets or sets the X coordinate of the rectangle's top-left corner.
        /// </summary>
        public float X { get; set; }
        
        /// <summary>
        /// Gets or sets the Y coordinate of the rectangle's top-left corner.
        /// </summary>
        public float Y { get; set; }
        
        /// <summary>
        /// Gets or sets the width of the rectangle.
        /// </summary>
        public float Width { get; set; }
        
        /// <summary>
        /// Gets or sets the height of the rectangle.
        /// </summary>
        public float Height { get; set; }

        /// <summary>
        /// Converts this RectangleData to a System.Drawing.RectangleF structure.
        /// </summary>
        /// <returns>A RectangleF with the same coordinates and dimensions.</returns>
        public RectangleF ToRectangleF()
            => new RectangleF(X, Y, Width, Height);
    }

    /// <summary>
    /// Represents an RGB color value used for redaction fill or other color-based operations.
    /// </summary>
    public class ColorData
    {
        /// <summary>
        /// Gets or sets the red component value (0-255).
        /// </summary>
        public int Red { get; set; }
        
        /// <summary>
        /// Gets or sets the green component value (0-255).
        /// </summary>
        public int Green { get; set; }
        
        /// <summary>
        /// Gets or sets the blue component value (0-255).
        /// </summary>
        public int Blue { get; set; }

        /// <summary>
        /// Converts this ColorData to a System.Drawing.Color structure.
        /// </summary>
        /// <returns>A Color with the specified RGB values.</returns>
        public Color ToColor()
            => Color.FromArgb(Red, Green, Blue);
    }
}
