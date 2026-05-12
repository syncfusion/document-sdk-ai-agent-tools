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
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;

namespace Syncfusion.AI.AgentTools.PDF
{
    /// <summary>
    /// Manager for handling PDF documents in memory during AI agent operations.
    /// Provides document lifecycle management with automatic cleanup.
    /// Supports both new PDF documents and loaded existing PDFs.
    /// </summary>
    public class PdfDocumentManager : DocumentManagerBase<PdfDocumentBase>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PdfDocumentManager"/> class.
        /// </summary>
        /// <param name="expirationTime">Time before documents are automatically cleaned up. Default is 30 minutes.</param>
        public PdfDocumentManager(TimeSpan? expirationTime = null)
            : base(expirationTime)
        {
        }

        /// <summary>
        /// Gets the document type managed by this manager.
        /// </summary>
        /// <returns>DocumentType.PDF indicating this manager handles PDF documents.</returns>
        public override DocumentType DocumentType => DocumentType.PDF;

        /// <summary>
        /// Gets the prefix used for generating unique PDF document identifiers.
        /// </summary>
        /// <returns>The prefix string "pdf_" used for PDF document IDs.</returns>
        protected override string DocumentIdPrefix => "pdf_";

        /// <summary>
        /// Creates a new empty PDF document instance.
        /// </summary>
        /// <returns>A new PdfDocument instance ready for editing.</returns>
        protected override PdfDocumentBase CreateDocumentInstance()
        {
            return new PdfDocument();
        }

        /// <summary>
        /// Imports an existing PDF document from the specified file path.
        /// </summary>
        /// <param name="filePath">The file path to the PDF document to load.</param>
        /// <returns>A PdfLoadedDocument instance loaded from the specified file.</returns>
        protected override PdfDocumentBase ImportDocumentInstance(string filePath)
        {
            return new PdfLoadedDocument(filePath, true);
        }

        /// <summary>
        /// Imports an encrypted PDF document from the specified file path using the provided password.
        /// </summary>
        /// <param name="filePath">The file path to the encrypted PDF document.</param>
        /// <param name="password">The password required to open the encrypted document.</param>
        /// <returns>A PdfLoadedDocument instance loaded from the specified file with decryption applied.</returns>
        protected override PdfDocumentBase ImportDocumentInstance(string filePath, string password)
        {
            return new PdfLoadedDocument(filePath, password, true);
        }

        /// <summary>
        /// Exports the PDF document to the specified file path.
        /// </summary>
        /// <param name="document">The PdfDocumentBase instance to export.</param>
        /// <param name="filePath">The file path where the document will be saved.</param>
        protected override void ExportDocumentInstance(PdfDocumentBase document, string filePath)
        {
            // Save PDF document
            using FileStream outputStream = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            outputStream.SetLength(0);
            document.Save(outputStream);
        }

        /// <summary>
        /// Closes and releases the resources associated with the PDF document.
        /// </summary>
        /// <param name="document">The PdfDocumentBase instance to close.</param>
        protected override void CloseDocument(PdfDocumentBase document)
        {
            document.Close(true);
        }

        /// <summary>
        /// Imports an existing PdfDocument instance directly into the manager without file I/O.
        /// This is useful for scenarios where a PdfDocument is created programmatically 
        /// (e.g., from Office document conversion) and needs to be stored in the manager.
        /// </summary>
        /// <param name="document">The PdfDocument instance to import.</param>
        /// <returns>The document ID assigned to the imported document.</returns>
        internal string ImportDocumentInstance(PdfDocumentBase document)
        {
            return ImportDocument(document);
        }

        /// <summary>
        /// Disposes the PDF Document Manager and all managed documents.
        /// </summary>
        public void Dispose()
        {
            Clear();
        }
    }
}
