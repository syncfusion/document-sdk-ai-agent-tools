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

using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Presentation;
using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Syncfusion.AI.AgentTools.Core
{
    /// <summary>
    /// Manages document lifecycle operations using an <see cref="IDocumentStorage"/> backend
    /// for persistent, externalized storage (DocumentStorage mode / Mode 2).
    /// </summary>
    /// <remarks>
    /// <para>
    /// Unlike the in-memory manager, this class never accesses the local file system directly.
    /// All I/O is delegated to the injected <see cref="IDocumentStorage"/> implementation,
    /// which may be backed by Azure Blob Storage, Amazon S3, or any other provider.
    /// </para>
    /// <para>
    /// This manager is stateless — it does not track an "active document."
    /// Every method requires an explicit document path or ID, making it safe for concurrent use
    /// across multiple requests and horizontally scaled instances.
    /// </para>
    /// </remarks>
    public sealed class DocumentStorageManager
    {
        private readonly IDocumentStorage _storage;

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentStorageManager"/> class
        /// with the specified storage backend.
        /// </summary>
        /// <param name="storage">
        /// The <see cref="IDocumentStorage"/> implementation used for all document I/O.
        /// </param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="storage"/> is <see langword="null"/>.</exception>
        public DocumentStorageManager(IDocumentStorage storage)
        {
            ArgumentNullException.ThrowIfNull(storage);
            _storage = storage;
        }

        /// <summary>
        /// Retrieves a Syncfusion document instance from the storage after deserializing it.        
        /// </summary>
        /// <param name="filePath">The file path to retrieve.</param>
        /// <param name="documentType">The document type to deserialize.</param>
        /// <param name="password">Optional password for encrypted documents.</param>
        /// <returns>The deserialized document instance, or null if not found.</returns>
        /// The returned instance is not tracked by the manager. After making mutations,
        /// call <see cref="SaveDocument"/> to persist changes back to storage.
        /// </remarks>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="filePath"/> is null.</exception>
        internal object? GetDocumentInstance(string filePath, DocumentType documentType, string? password = null)
        {
            ArgumentNullException.ThrowIfNull(filePath);

            if (!_storage.Exists(filePath)) return null;

            if (documentType == DocumentType.PDF)
            {
                var stream = _storage.Read(filePath);
                return OpenDocumentInstance(stream, documentType, password);
            }
            else
            {
                using var stream = _storage.Read(filePath);
                return OpenDocumentInstance(stream, documentType, password);
            }
        }


        /// <summary>
        /// Returns the raw document stream from the storage without deserializing it.
        /// </summary>
        /// <param name="filePath">The storage path identifying the document to retrieve.</param>
        /// <returns>
        /// A <see cref="Stream"/> containing the raw document bytes,
        /// or <see langword="null"/> if no document exists at <paramref name="filePath"/>.
        /// The caller is responsible for disposing the returned stream.
        /// </returns>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="filePath"/> is <see langword="null"/>.</exception>
        internal Stream? GetDocumentStream(string filePath)
        {
            ArgumentNullException.ThrowIfNull(filePath);

            if (!_storage.Exists(filePath)) return null;

            return _storage.Read(filePath);
        }

        /// <summary>
        /// Serializes the document to a stream, writes it back to the storage at the
        /// specified path, and closes the document instance.
        /// </summary>
        /// <param name="filePath">The storage path to write the document to.</param>
        /// <param name="document">The document instance to serialize and persist.</param>
        /// <param name="documentType">The document type, used to select the correct serializer.</param>
        /// <remarks>
        /// This method must be called after every mutating operation on a transient document
        /// instance obtained via <see cref="GetDocumentInstance"/>. The document is closed
        /// after saving, even if the write fails.
        /// </remarks>
        /// <exception cref="ArgumentNullException">
        /// Thrown when <paramref name="filePath"/> or <paramref name="document"/> is <see langword="null"/>.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the underlying <see cref="IDocumentStorage.Write"/> call fails.
        /// </exception>
        /// <exception cref="NotSupportedException">
        /// Thrown when <paramref name="documentType"/> is not yet supported.
        /// </exception>
        internal void SaveDocument(string filePath, object document, DocumentType documentType)
        {
            ArgumentNullException.ThrowIfNull(filePath);
            ArgumentNullException.ThrowIfNull(document);

            using var stream = SaveDocumentInstance(document, documentType, filePath);
            try
            {
                if (!_storage.Write(filePath, stream))
                {
                    throw new InvalidOperationException($"Failed to save document to storage with ID: {filePath}");
                }
            }
            finally
            {
                CloseDocumentInstance(document, documentType);
            }
        }

        /// <summary>
        /// Writes a raw stream (e.g., a rendered image or exported artifact) directly into
        /// the storage without any document-level serialization.
        /// </summary>
        /// <param name="filePath">The storage path to write the stream under.</param>
        /// <param name="stream">The raw stream to persist. The stream position should be set before calling.</param>
        /// <exception cref="ArgumentNullException">
        /// Thrown when <paramref name="filePath"/> or <paramref name="stream"/> is <see langword="null"/>.
        /// </exception>
        /// <exception cref="InvalidOperationException">
        /// Thrown when the underlying <see cref="IDocumentStorage.Write"/> call fails.
        /// </exception>
        internal void WriteRawStream(string filePath, Stream stream)
        {
            ArgumentNullException.ThrowIfNull(filePath);
            ArgumentNullException.ThrowIfNull(stream);
            if (!_storage.Write(filePath, stream))
            {
                throw new InvalidOperationException($"Failed to write raw stream to storage with ID: {filePath}");
            }
        }

        /// <summary>
        /// Checks whether a document exists at the specified path in the storage.
        /// </summary>
        /// <param name="filePath">The storage path to check.</param>
        /// <returns>
        /// <see langword="true"/> if a document exists at <paramref name="filePath"/>;
        /// otherwise, <see langword="false"/>.
        /// </returns>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="filePath"/> is <see langword="null"/>.</exception>
        internal bool HasDocument(string filePath)
        {
            ArgumentNullException.ThrowIfNull(filePath);
            return _storage.Exists(filePath);
        }

        /// <summary>
        /// Serializes the document to a <see cref="MemoryStream"/> in the appropriate format.
        /// </summary>
        /// <param name="document">The document instance to serialize.</param>
        /// <param name="documentType">The document type, used to select the correct format.</param>
        /// <param name="filePath">The file path, used to determine the save format based on extension.</param>
        /// <returns>
        /// A <see cref="Stream"/> positioned at the beginning, containing the serialized document bytes.
        /// </returns>
        /// <exception cref="NotSupportedException">
        /// Thrown when <paramref name="documentType"/> is not yet supported.
        /// </exception>
        internal static Stream SaveDocumentInstance(object document, DocumentType documentType, string filePath)
        {
            var ms = new MemoryStream();
            try
            {
                switch (documentType)
                {
                    case DocumentType.Word:
                        SaveWordDocument((WordDocument)document, ms, filePath);
                        break;
                    case DocumentType.PDF:
                        ((PdfDocumentBase)document).Save(ms);
                        break;
                    case DocumentType.Excel:
                        SaveExcelDocument((IWorkbook)document, ms, filePath);
                        break;
                    case DocumentType.PowerPoint:
                        SavePresentationDocument((IPresentation)document, ms, filePath);
                        break;
                    default:
                        throw new NotSupportedException($"DocumentType '{documentType}' is not supported.");
                }
                ms.Position = 0;
                return ms;
            }
            catch
            {
                ms.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Saves a Word document to the specified stream in the format determined by the file path extension.
        /// </summary>
        /// <param name="document">The Word document to save.</param>
        /// <param name="stream">The stream to save the document to.</param>
        /// <param name="filePath">The file path with the extension that determines the save format.</param>
        private static void SaveWordDocument(WordDocument document, Stream stream, string filePath)
        {
            var formatType = GetWordFormatType(filePath);
            document.Save(stream, formatType);
        }

        /// <summary>
        /// Determines the appropriate <see cref="FormatType"/> based on the file extension.
        /// </summary>
        /// <param name="filePath">The file path with the extension to check.</param>
        /// <returns>The corresponding <see cref="FormatType"/> for the file extension.</returns>
        private static DocIO.FormatType GetWordFormatType(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension switch
            {
                ".docx" => DocIO.FormatType.Docx,
                ".doc" => DocIO.FormatType.Doc,
                ".dot" => DocIO.FormatType.Dot,
                ".dotx" => DocIO.FormatType.Dotx,
                ".dotm" => DocIO.FormatType.Dotm,
                ".docm" => DocIO.FormatType.Docm,
                ".xml" => DocIO.FormatType.WordML,
                ".odt" => DocIO.FormatType.Odt,
                ".md" => DocIO.FormatType.Markdown,
                ".html" => DocIO.FormatType.Html,
                ".rtf" => DocIO.FormatType.Rtf,
                ".txt" => DocIO.FormatType.Txt,
                _ => DocIO.FormatType.Docx // Default to .docx format
            };
        }

        /// <summary>
        /// Saves a presentation document to the specified stream in the format determined by the file path extension.
        /// </summary>
        /// <param name="document">The presentation document to save.</param>
        /// <param name="stream">The stream to save the document to.</param>
        /// <param name="filePath">The file path with the extension that determines the save format.</param>
        private static void SavePresentationDocument(IPresentation document, Stream stream, string filePath)
        {
            var formatType = GetPresentationFormatType(filePath);
            document.Save(stream, formatType);
        }

        /// <summary>
        /// Determines the appropriate <see cref="Presentation.FormatType"/> based on the file extension.
        /// </summary>
        /// <param name="filePath">The file path with the extension to check.</param>
        /// <returns>The corresponding <see cref="Presentation.FormatType"/> for the file extension.</returns>
        private static Presentation.FormatType GetPresentationFormatType(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension switch
            {
                ".pptx" => Presentation.FormatType.Pptx,
                ".pptm" => Presentation.FormatType.Pptm,
                ".potm" => Presentation.FormatType.Potm,
                ".potx" => Presentation.FormatType.Potx,
                ".md" => Presentation.FormatType.Markdown,
                _ => Presentation.FormatType.Pptx // Default to .pptx format
            };
        }
        /// <summary>
        /// Saves an Excel workbook to the specified stream in the format determined by the file path extension.
        /// </summary>
        /// <param name="workbook">The Excel workbook to save.</param>
        /// <param name="stream">The stream to save the workbook to.</param>
        /// <param name="filePath">The file path with the extension that determines the save format.</param>
        private static void SaveExcelDocument(IWorkbook workbook, Stream stream, string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            
            if (extension == ".csv")
            {
                workbook.SaveAs(stream, ",");
            }
            else if (extension == ".tsv")
            {
                workbook.SaveAs(stream, "\t");
            }
            else
            {
                var saveType = GetExcelSaveType(filePath);
                workbook.SaveAs(stream,saveType);
            }
        }

        /// <summary>
        /// Determines the appropriate <see cref="ExcelSaveType"/> based on the file extension.
        /// </summary>
        /// <param name="filePath">The file path with the extension to check.</param>
        /// <returns>The corresponding <see cref="ExcelSaveType"/> for the file extension.</returns>
        private static ExcelSaveType GetExcelSaveType(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension switch
            {
                ".xlsx" => ExcelSaveType.SaveAsXLS,
                ".xls" => ExcelSaveType.SaveAsXLS,
                ".xlsm" => ExcelSaveType.SaveAsMacro,
                ".xlsb" => ExcelSaveType.SaveAsXLSB,
                ".xlt" => ExcelSaveType.SaveAsTemplate,
                ".xltx" => ExcelSaveType.SaveAsTemplate,
                ".xltm" => ExcelSaveType.SaveAsMacroTemplate,
                ".ods" => ExcelSaveType.SaveAsODS,
                ".md" => ExcelSaveType.Markdown,
                _ => ExcelSaveType.SaveAsXLS // Default to .xlsx format
            };
        }

        /// <summary>
        /// Deserializes a document from the given stream based on the specified document type.
        /// </summary>
        /// <param name="stream">The stream containing the raw document bytes.</param>
        /// <param name="documentType">The document type, used to select the correct deserializer.</param>
        /// <returns>The deserialized document instance.</returns>
        /// <exception cref="NotSupportedException">
        /// Thrown when <paramref name="documentType"/> is not yet supported.
        /// </exception>
        /// <summary>
        /// Deserializes a document from the given stream based on the specified document type.
        /// </summary>
        /// <param name="stream">The stream containing the raw document bytes.</param>
        /// <param name="documentType">The document type, used to select the correct deserializer.</param>
        /// <param name="password">Optional password for decrypting encrypted documents.</param>
        /// <returns>The deserialized document instance.</returns>
        /// <exception cref="NotSupportedException">
        /// Thrown when <paramref name="documentType"/> is not yet supported.
        /// </exception>
        private static object OpenDocumentInstance(Stream stream, DocumentType documentType, string? password = null)
        {
            if (password != null)
            {
                return documentType switch
                {
                    DocumentType.Word => new WordDocument(stream, DocIO.FormatType.Automatic, password),
                    DocumentType.PDF => new PdfLoadedDocument(stream, password),
                    DocumentType.Excel => new ExcelEngine().Excel.Workbooks.Open(stream, ExcelParseOptions.Default, false, password),
                    DocumentType.PowerPoint => Syncfusion.Presentation.Presentation.Open(stream, password),
                    _ => throw new NotSupportedException($"DocumentType '{documentType}' is not supported.")
                };
            }
            else
            {
                return documentType switch
                {
                    DocumentType.Word => new WordDocument(stream, DocIO.FormatType.Automatic),
                    DocumentType.PDF => new PdfLoadedDocument(stream),
                    DocumentType.Excel => new ExcelEngine().Excel.Workbooks.Open(stream),
                    DocumentType.PowerPoint => Syncfusion.Presentation.Presentation.Open(stream),
                    _ => throw new NotSupportedException($"DocumentType '{documentType}' is not supported.")
                };
            }
        }
        /// <summary>
        /// Closes and releases resources held by the document instance.
        /// </summary>
        /// <param name="document">The document instance to close.</param>
        /// <param name="documentType">The document type, used to invoke the correct close method.</param>
        /// <exception cref="NotSupportedException">
        /// Thrown when <paramref name="documentType"/> is not yet supported.
        /// </exception>
        private static void CloseDocumentInstance(object document, DocumentType documentType)
        {
            switch (documentType)
            {
                case DocumentType.Word:
                    ((WordDocument)document).Close();
                    break;
                case DocumentType.PDF:
                    ((PdfDocumentBase)document).Close(true);
                    break;
                case DocumentType.Excel:
                    ((IWorkbook)document).Close();
                    break;
                case DocumentType.PowerPoint:
                    ((IPresentation)document).Close();
                    break;
                default:
                    throw new NotSupportedException($"DocumentType '{documentType}' is not supported.");
            }
        }
    }
}
