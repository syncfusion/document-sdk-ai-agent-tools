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

namespace Syncfusion.AI.AgentTools.Core
{
	/// <summary>
    /// Identifies the operating mode of a document manager as seen by tool classes.
    /// </summary>
    public enum DocumentManagerMode
    {
        /// <summary>Documents are held in memory for the session lifetime.</summary>
        InMemory,

        /// <summary>Documents are read from and written to an IDocumentStorage on every tool call.</summary>
        DocumentStorage
    }
	
    /// <summary>
    /// Document type enumeration for manager.
    /// </summary>
    public enum DocumentType
    {
        /// <summary>
        /// Word document (flow document).
        /// </summary>
        Word,

        /// <summary>
        /// PDF document.
        /// </summary>
        PDF,

        /// <summary>
        /// Excel workbook.
        /// </summary>
        Excel,

        /// <summary>
        /// PowerPoint presentation.
        /// </summary>
        PowerPoint
    }

    /// <summary>
    /// Base interface for all document managers. A manager is an in-memory
    /// container where documents are stored and managed during processing.
    /// </summary>
    public interface IDocumentManager
    {
        /// <summary>
        /// Gets the document type handled by this manager.
        /// </summary>
        DocumentType DocumentType { get; }
    }

    /// <summary>
    /// A centralized collection that maintains one manager for each document type.
    /// Enables higher-level components to dynamically pick the appropriate manager
    /// at runtime based on the document type being handled.
    /// </summary>
    public class DocumentManagerCollection
    {
        private readonly Dictionary<DocumentType, IDocumentManager> _managers = new Dictionary<DocumentType, IDocumentManager>();

        /// <summary>
        /// Adds a manager for a specific document type.
        /// </summary>
        /// <param name="documentType">The document type.</param>
        /// <param name="manager">The manager instance.</param>
        public void AddManager(DocumentType documentType, IDocumentManager manager)
        {
            ArgumentNullException.ThrowIfNull(manager);
            _managers[documentType] = manager;
        }

        /// <summary>
        /// Retrieves a manager for a specific document type.
        /// </summary>
        /// <param name="documentType">The document type.</param>
        /// <returns>The manager, or null if not registered.</returns>
        internal IDocumentManager? GetManager(DocumentType documentType)
        {
            return _managers.TryGetValue(documentType, out var manager) ? manager : null;
        }

        /// <summary>
        /// Retrieves a typed manager for a specific document type.
        /// </summary>
        /// <typeparam name="TDocument">The document type class.</typeparam>
        /// <param name="documentType">The document type.</param>
        /// <returns>The typed manager, or null if not registered or type mismatch.</returns>
        internal DocumentManagerBase<TDocument>? GetManager<TDocument>(DocumentType documentType)
            where TDocument : class
        {
            return GetManager(documentType) as DocumentManagerBase<TDocument>;
        }

        /// <summary>
        /// Checks if a manager is registered for a document type.
        /// </summary>
        /// <param name="documentType">The document type.</param>
        /// <returns>True if registered, false otherwise.</returns>
        internal bool HasManager(DocumentType documentType)
        {
            return _managers.ContainsKey(documentType);
        }
    }
}
