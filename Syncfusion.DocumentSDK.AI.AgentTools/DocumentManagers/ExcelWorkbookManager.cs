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
using Syncfusion.AI.AgentTools.Core;
using Syncfusion.XlsIO;

namespace Syncfusion.AI.AgentTools.Excel
{
    /// <summary>
    /// Manager for handling Excel workbooks in memory during AI agent operations.
    /// Provides workbook lifecycle management with automatic cleanup.
    /// </summary>
    public class ExcelWorkbookManager : DocumentManagerBase<IWorkbook>, IDisposable
    {
        private readonly ExcelEngine _excelEngine;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelWorkbookManager"/> class.
        /// </summary>
        /// <param name="expirationTime">Time before workbooks are automatically cleaned up. Default is 30 minutes.</param>
        public ExcelWorkbookManager(TimeSpan? expirationTime = null)
            : base(expirationTime)
        {
            _excelEngine = new ExcelEngine();
        }

        /// <summary>
        /// Gets the document type managed by this manager.
        /// </summary>
        /// <returns>DocumentType.Excel indicating this manager handles Excel workbooks.</returns>
        public override DocumentType DocumentType => DocumentType.Excel;

        /// <summary>
        /// Gets the prefix used for generating unique workbook identifiers.
        /// </summary>
        /// <returns>The prefix string "xl_" used for Excel workbook IDs.</returns>
        protected override string DocumentIdPrefix => "xl_";

        /// <summary>
        /// Creates a new empty Excel workbook instance.
        /// </summary>
        /// <returns>A new IWorkbook instance ready for editing.</returns>
        protected override IWorkbook CreateDocumentInstance()
        {
            return _excelEngine.Excel.Workbooks.Create(1);
        }

        /// <summary>
        /// Imports an existing Excel workbook from the specified file path.
        /// </summary>
        /// <param name="filePath">The file path to the Excel workbook to load.</param>
        /// <returns>An IWorkbook instance loaded from the specified file.</returns>
        protected override IWorkbook ImportDocumentInstance(string filePath)
        {
            using FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            return _excelEngine.Excel.Workbooks.Open(fileStream);
        }

        /// <summary>
        /// Imports an encrypted Excel workbook from the specified file path using the provided password.
        /// </summary>
        /// <param name="filePath">The file path to the encrypted Excel workbook.</param>
        /// <param name="password">The password required to open the encrypted workbook.</param>
        /// <returns>An IWorkbook instance loaded from the specified file with decryption applied.</returns>
        protected override IWorkbook ImportDocumentInstance(string filePath, string password)
        {
            return _excelEngine.Excel.Workbooks.Open(filePath, ExcelParseOptions.Default, true, password);
        }

        /// <summary>
        /// Exports the Excel workbook to the specified file path.
        /// </summary>
        /// <param name="document">The IWorkbook instance to export.</param>
        /// <param name="filePath">The file path where the workbook will be saved. CSV format is supported with .csv extension.</param>
        protected override void ExportDocumentInstance(IWorkbook document, string filePath)
        {
            string extension = Path.GetExtension(filePath);
            if (extension == ".csv")
                document.ActiveSheet.SaveAs(filePath,",");
            else
                document.SaveAs(filePath);
        }

        /// <summary>
        /// Closes and releases the resources associated with the Excel workbook.
        /// </summary>
        /// <param name="document">The IWorkbook instance to close.</param>
        protected override void CloseDocument(IWorkbook document)
        {
            document.Close();
        }

        /// <summary>
        /// Disposes the Excel engine and all managed workbooks.
        /// </summary>
        public void Dispose()
        {
            Clear();
            _excelEngine?.Dispose();
        }

        /// <summary>
        /// Gets or sets the currently active workbook ID.
        /// </summary>
        public string? ActiveWorkbookId
        {
            get => ActiveDocumentId;
            set => SetActiveDocument(value!);
        }
    }
}
