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
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Syncfusion.AI.AgentTools.Word
{
    /// <summary>
    /// Manager for handling Word documents in memory during AI agent operations.
    /// Provides document lifecycle management with automatic cleanup.
    /// </summary>
    public class WordDocumentManager : DocumentManagerBase<WordDocument>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="WordDocumentManager"/> class.
        /// </summary>
        /// <param name="expirationTime">Time before documents are automatically cleaned up. Default is 30 minutes.</param>
        public WordDocumentManager(TimeSpan? expirationTime = null)
            : base(expirationTime)
        {
        }

        /// <summary>
        /// Gets the document type managed by this manager.
        /// </summary>
        /// <returns>DocumentType.Word indicating this manager handles Word documents.</returns>
        public override DocumentType DocumentType => DocumentType.Word;

        /// <summary>
        /// Gets the prefix used for generating unique document identifiers.
        /// </summary>
        /// <returns>The prefix string "doc_" used for Word document IDs.</returns>
        protected override string DocumentIdPrefix => "doc_";

        /// <summary>
        /// Creates a new empty Word document instance.
        /// </summary>
        /// <returns>A new WordDocument instance ready for editing.</returns>
        protected override WordDocument CreateDocumentInstance()
        {
            return new WordDocument();
        }

        /// <summary>
        /// Stores an existing <see cref="WordDocument"/> instance in the manager
        /// and marks it as the active document. This is useful when a document is
        /// created programmatically and needs to be kept in memory for further operations.
        /// </summary>
        /// <param name="document">The <see cref="WordDocument"/> instance to store.</param>
        /// <returns>The same <see cref="WordDocument"/> instance that was stored.</returns>
        internal WordDocument AddDocument(WordDocument document)
        {
            ArgumentNullException.ThrowIfNull(document);

            // Import the instance into the base manager which handles id generation,
            // expiration cleanup and active document tracking.
            ImportDocument(document);

            return document;
        }

        /// <summary>
        /// Imports an existing Word document from the specified file path.
        /// </summary>
        /// <param name="filePath">The file path to the Word document to load.</param>
        /// <returns>A WordDocument instance loaded from the specified file.</returns>
        protected override WordDocument ImportDocumentInstance(string filePath)
        {
            return new WordDocument(filePath);
        }

        /// <summary>
        /// Imports an encrypted Word document from the specified file path using the provided password.
        /// </summary>
        /// <param name="filePath">The file path to the encrypted Word document.</param>
        /// <param name="password">The password required to open the encrypted document.</param>
        /// <returns>A WordDocument instance loaded from the specified file with decryption applied.</returns>
        protected override WordDocument ImportDocumentInstance(string filePath, string password)
        {
            // Open encrypted Word document with password
            return new WordDocument(filePath, password);
        }
        /// <summary>
        /// Exports the Word document to the specified file path.
        /// </summary>
        /// <param name="document">The WordDocument instance to export.</param>
        /// <param name="filePath">The file path where the document will be saved.</param>
        protected override void ExportDocumentInstance(WordDocument document, string filePath)
        {
            document.Save(filePath);
        }

        /// <summary>
        /// Closes and releases the resources associated with the Word document.
        /// </summary>
        /// <param name="document">The WordDocument instance to close.</param>
        protected override void CloseDocument(WordDocument document)
        {
            document.Close();
        }

        /// <summary>
        /// Disposes the Word Document Manager and it's documents.
        /// </summary>
        public void Dispose()
        {
            Clear();
        }
    }
}
