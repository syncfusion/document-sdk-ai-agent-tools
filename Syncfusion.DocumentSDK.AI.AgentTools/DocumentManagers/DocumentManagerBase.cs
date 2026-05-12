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

namespace Syncfusion.AI.AgentTools.Core
{
    /// <summary>
    /// Abstract base class for document managers that provides common functionality
    /// for managing documents in memory with automatic cleanup.
    /// </summary>
    /// <typeparam name="TDocument">The type of document managed by this manager.</typeparam>
    public abstract class DocumentManagerBase<TDocument> : IDocumentManager
        where TDocument : class
    {
        private readonly Dictionary<string, TDocument> _documents = new Dictionary<string, TDocument>();
        private readonly Dictionary<string, DateTime> _lastAccessed = new Dictionary<string, DateTime>();
        private readonly TimeSpan _expirationTime;
        private string? _activeDocumentId;

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentManagerBase{TDocument}"/> class.
        /// </summary>
        /// <param name="expirationTime">Time before documents are automatically cleaned up. Default is 10 minutes.</param>
        protected DocumentManagerBase(TimeSpan? expirationTime = null)
        {
            _expirationTime = expirationTime ?? TimeSpan.FromMinutes(10);
        }

        /// <inheritdoc/>
        public abstract DocumentType DocumentType { get; }

        /// <summary>
        /// Gets the prefix used for generating document IDs (e.g., "doc_", "pdf_", "ppt_").
        /// </summary>
        /// <returns>The document ID prefix specific to the document type.</returns>
        protected abstract string DocumentIdPrefix { get; }

        /// <summary>
        /// Creates a new document instance.
        /// </summary>
        /// <returns>A new document of type TDocument.</returns>
        protected abstract TDocument CreateDocumentInstance();

        /// <summary>
        /// Imports a document from a file path.
        /// </summary>
        /// <param name="filePath">The file path to import from.</param>
        /// <returns>The imported document.</returns>
        protected abstract TDocument ImportDocumentInstance(string filePath);

        /// <summary>
        /// Imports a document from a file path with password support.
        /// </summary>
        /// <param name="filePath">The file path to import from.</param>
        /// <param name="password">Password for encrypted document.</param>
        /// <returns>The imported document.</returns>
        protected abstract TDocument ImportDocumentInstance(string filePath, string password);

        /// <summary>
        /// Exports a document to a file path.
        /// </summary>
        /// <param name="document">The document to export.</param>
        /// <param name="filePath">The file path to export to.</param>
        protected abstract void ExportDocumentInstance(TDocument document, string filePath);

        /// <summary>
        /// Closes a document and releases its resources.
        /// </summary>
        /// <param name="document">The document to close.</param>
        protected abstract void CloseDocument(TDocument document);

        /// <summary>
        /// Creates a new document and stores it in the manager with an auto-generated or specified ID.
        /// </summary>
        /// <param name="documentId">Optional document ID. If null, a unique ID is auto-generated.</param>
        /// <returns>The newly created document instance.</returns>
        internal TDocument CreateDocument(string? documentId = null)
        {
            CleanupExpiredDocuments();

            documentId ??= GenerateDocumentId();

            var document = CreateDocumentInstance();
            _documents[documentId] = document;
            _lastAccessed[documentId] = DateTime.UtcNow;
            _activeDocumentId = documentId;

            return document;
        }

        /// <summary>
        /// Imports a document from a file path and stores it in the manager.
        /// </summary>
        /// <param name="filePath">The file path to import from.</param>
        /// <returns>The imported document instance.</returns>
        /// <exception cref="ArgumentNullException">Thrown when filePath is null.</exception>
        /// <exception cref="FileNotFoundException">Thrown when the file does not exist.</exception>
        internal TDocument ImportDocument(string filePath)
        {
            ArgumentNullException.ThrowIfNull(filePath);

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File not found: {filePath}", filePath);

            CleanupExpiredDocuments();

            string documentId = GenerateDocumentId();

            var document = ImportDocumentInstance(filePath);
            _documents[documentId] = document;
            _lastAccessed[documentId] = DateTime.UtcNow;
            _activeDocumentId = documentId;

            return document;
        }

        /// <summary>
        /// Imports a document into the manager from a file with password support.
        /// </summary>
        /// <param name="filePath">The file path to import from.</param>
        /// <param name="password">Password for encrypted document.</param>
        /// <returns>The imported document.</returns>
        internal TDocument ImportDocument(string filePath, string password)
        {
            ArgumentNullException.ThrowIfNull(filePath);
            ArgumentNullException.ThrowIfNull(password);

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"File not found: {filePath}", filePath);

            CleanupExpiredDocuments();

            string documentId = GenerateDocumentId();

            var document = ImportDocumentInstance(filePath, password);
            _documents[documentId] = document;
            _lastAccessed[documentId] = DateTime.UtcNow;
            _activeDocumentId = documentId;

            return document;
        }

        /// <summary>
        /// Imports an existing document instance directly into the manager without file I/O.
        /// This is useful for scenarios where a document is created programmatically 
        /// (e.g., from conversions, merges, or other operations) and needs to be stored in the manager.
        /// </summary>
        /// <param name="document">The document instance to import.</param>
        /// <returns>The document ID assigned to the imported document.</returns>
        protected string ImportDocument(TDocument document)
        {
            ArgumentNullException.ThrowIfNull(document);

            CleanupExpiredDocuments();

            string documentId = GenerateDocumentId();

            _documents[documentId] = document;
            _lastAccessed[documentId] = DateTime.UtcNow;
            _activeDocumentId = documentId;

            return documentId;
        }

        /// <summary>
        /// Exports a document to a file path using the manager's storage.
        /// </summary>
        /// <param name="filePath">The file path to export to.</param>
        /// <param name="documentId">Optional document ID. If null, uses the active document.</param>
        /// <exception cref="ArgumentNullException">Thrown when filePath is null.</exception>
        /// <exception cref="InvalidOperationException">Thrown when the document is not found.</exception>
        internal void ExportDocument(string filePath, string? documentId = null)
        {
            ArgumentNullException.ThrowIfNull(filePath);

            var document = GetDocument(documentId);
            if (document == null)
                throw new InvalidOperationException($"Document not found: {documentId ?? "(active)"}");

            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            ExportDocumentInstance(document, filePath);
        }

        /// <summary>
        /// Exports the provided document to a file path.
        /// </summary>
        /// <param name="filePath">The file path to export to.</param>
        /// <param name="document">The document instance to export.</param>
        /// <exception cref="ArgumentNullException">Thrown when filePath or document is null.</exception>
        internal void ExportDocument(string filePath, TDocument document)
        {
            ArgumentNullException.ThrowIfNull(filePath);

            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            ExportDocumentInstance(document, filePath);
        }

        /// <summary>
        /// Retrieves all document IDs currently managed by this manager.
        /// </summary>
        /// <returns>A list of all document IDs.</returns>
        internal List<string> GetAllDocumentIds()
        {
            CleanupExpiredDocuments();
            return new List<string>(_documents.Keys);
        }

        /// <summary>
        /// Checks whether a document with the specified ID exists in the manager.
        /// </summary>
        /// <param name="documentId">The document ID to check.</param>
        /// <returns>True if the document exists; otherwise, false.</returns>
        internal bool HasDocument(string documentId)
        {
            ArgumentNullException.ThrowIfNull(documentId);
            CleanupExpiredDocuments();
            return _documents.ContainsKey(documentId);
        }

        /// <summary>
        /// Removes a document from the manager and closes it.
        /// </summary>
        /// <param name="documentId">The document ID to remove.</param>
        /// <returns>True if the document was removed; false if it was not found.</returns>
        internal bool RemoveDocument(string documentId)
        {
            ArgumentNullException.ThrowIfNull(documentId);

            if (_documents.TryGetValue(documentId, out var document))
            {
                CloseDocument(document);
                _documents.Remove(documentId);
                _lastAccessed.Remove(documentId);

                if (_activeDocumentId == documentId)
                    _activeDocumentId = null;

                return true;
            }

            return false;
        }

        /// <summary>
        /// Removes and closes all documents from the manager.
        /// </summary>
        internal void Clear()
        {
            foreach (var document in _documents.Values)
            {
                CloseDocument(document);
            }

            _documents.Clear();
            _lastAccessed.Clear();
            _activeDocumentId = null;
        }

        /// <summary>
        /// Retrieves a document from the manager by ID or returns the active document if no ID is specified.
        /// </summary>
        /// <param name="documentId">Optional document ID. If null, returns the active document.</param>
        /// <returns>The document instance, or null if not found.</returns>
        internal TDocument? GetDocument(string? documentId = null)
        {
            CleanupExpiredDocuments();

            documentId ??= _activeDocumentId;

            if (documentId != null && _documents.TryGetValue(documentId, out var document))
            {
                _lastAccessed[documentId] = DateTime.UtcNow;
                return document;
            }

            return null;
        }

        /// <summary>
        /// Sets the specified document as the active document.
        /// </summary>
        /// <param name="documentId">The document ID to set as active.</param>
        /// <exception cref="ArgumentNullException">Thrown when documentId is null.</exception>
        /// <exception cref="InvalidOperationException">Thrown when the document is not found.</exception>
        internal void SetActiveDocument(string documentId)
        {
            ArgumentNullException.ThrowIfNull(documentId);

            if (!_documents.ContainsKey(documentId))
                throw new InvalidOperationException($"Document not found: {documentId}");

            _activeDocumentId = documentId;
        }

        /// <summary>
        /// Gets the currently active document ID.
        /// </summary>
        /// <returns>The active document ID, or null if no document is active.</returns>
        internal string? ActiveDocumentId => _activeDocumentId;

        /// <summary>
        /// Generates a unique document ID using the configured prefix and a GUID.
        /// </summary>
        /// <returns>A unique document ID string.</returns>
        private string GenerateDocumentId()
        {
            return $"{DocumentIdPrefix}{Guid.NewGuid():N}";
        }

        /// <summary>
        /// Removes and closes documents that have exceeded their expiration time.
        /// </summary>
        private void CleanupExpiredDocuments()
        {
            var now = DateTime.UtcNow;
            var expiredIds = new List<string>();

            foreach (var kvp in _lastAccessed)
            {
                if (now - kvp.Value > _expirationTime)
                {
                    expiredIds.Add(kvp.Key);
                }
            }

            foreach (var id in expiredIds)
            {
                RemoveDocument(id);
            }
        }
    }
}
