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
using Syncfusion.Presentation;

namespace Syncfusion.AI.AgentTools.PowerPoint
{
    /// <summary>
    /// Manager for handling PowerPoint presentations in memory during AI agent operations.
    /// Provides presentation lifecycle management with automatic cleanup.
    /// </summary>
    public class PresentationManager : DocumentManagerBase<IPresentation>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PresentationManager"/> class.
        /// </summary>
        /// <param name="expirationTime">Time before presentations are automatically cleaned up. Default is 30 minutes.</param>
        public PresentationManager(TimeSpan? expirationTime = null)
            : base(expirationTime)
        {
        }

        /// <summary>
        /// Gets the document type managed by this manager.
        /// </summary>
        /// <returns>DocumentType.PowerPoint indicating this manager handles PowerPoint presentations.</returns>
        public override DocumentType DocumentType => DocumentType.PowerPoint;

        /// <summary>
        /// Gets the prefix used for generating unique presentation identifiers.
        /// </summary>
        /// <returns>The prefix string "ppt_" used for PowerPoint presentation IDs.</returns>
        protected override string DocumentIdPrefix => "ppt_";

        /// <summary>
        /// Creates a new empty PowerPoint presentation instance.
        /// </summary>
        /// <returns>A new IPresentation instance ready for editing.</returns>
        protected override IPresentation CreateDocumentInstance()
        {
            return Syncfusion.Presentation.Presentation.Create();
        }

        /// <summary>
        /// Imports an existing PowerPoint presentation from the specified file path.
        /// </summary>
        /// <param name="filePath">The file path to the PowerPoint presentation to load.</param>
        /// <returns>An IPresentation instance loaded from the specified file.</returns>
        protected override IPresentation ImportDocumentInstance(string filePath)
        {
            // Open existing PowerPoint presentation
            using FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            return Syncfusion.Presentation.Presentation.Open(fileStream);
        }

        /// <summary>
        /// Imports an encrypted PowerPoint presentation from the specified file path using the provided password.
        /// </summary>
        /// <param name="filePath">The file path to the encrypted PowerPoint presentation.</param>
        /// <param name="password">The password required to open the encrypted presentation.</param>
        /// <returns>An IPresentation instance loaded from the specified file with decryption applied.</returns>
        protected override IPresentation ImportDocumentInstance(string filePath, string password)
        {
            // Open encrypted presentation with password
            using FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            return Syncfusion.Presentation.Presentation.Open(fileStream, password);
        }

        /// <summary>
        /// Exports the PowerPoint presentation to the specified file path.
        /// </summary>
        /// <param name="document">The IPresentation instance to export.</param>
        /// <param name="filePath">The file path where the presentation will be saved in PPTX format.</param>
        protected override void ExportDocumentInstance(IPresentation document, string filePath)
        {
            // Save PowerPoint presentation (PPTX format)
            using FileStream outputStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            document.Save(outputStream);
        }

        /// <summary>
        /// Closes and releases the resources associated with the PowerPoint presentation.
        /// </summary>
        /// <param name="document">The IPresentation instance to close.</param>
        protected override void CloseDocument(IPresentation document)
        {
            document.Close();
        }

        /// <summary>
        /// Gets the currently active presentation ID.
        /// </summary>
        /// <returns>The active presentation ID, or null if no presentation is active.</returns>
        public string? ActivePresentationId => ActiveDocumentId;

        /// <summary>
        /// Disposes the Presentation Manager and all managed presentations.
        /// </summary>
        public void Dispose()
        {
            Clear();
        }
    }
}
