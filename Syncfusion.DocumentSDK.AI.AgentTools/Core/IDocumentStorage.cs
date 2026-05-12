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

namespace Syncfusion.AI.AgentTools.Core
{
    /// <summary>
    /// Defines a contract for reading, writing, and deleting documents in a storage provider.
    /// </summary>
    /// <remarks>
    /// Implement this interface to integrate with any storage backend such as Azure Blob Storage,
    /// Amazon S3, local file system, or any other custom provider. The implementation is used by
    /// <see cref="DocumentStorageManager"/> in DocumentStorage mode (Mode 2) to manage document
    /// lifecycle operations.
    /// </remarks>
    public interface IDocumentStorage
    {
        /// <summary>
        /// Reads a document from the storage and returns its content as a stream.
        /// </summary>
        /// <param name="filePath">The relative or absolute path identifying the document in the storage.</param>
        /// <returns>
        /// A <see cref="Stream"/> containing the document content. The caller is responsible for disposing the stream.
        /// </returns>
        Stream Read(string filePath);

        /// <summary>
        /// Writes a document to the storage, creating or overwriting the file at the specified path.
        /// </summary>
        /// <param name="filePath">The relative or absolute path identifying the document in the storage.</param>
        /// <param name="documentStream">The stream containing the document content to write.</param>
        /// <returns>
        /// <see langword="true"/> if the write operation succeeded; otherwise, <see langword="false"/>.
        /// </returns>
        bool Write(string filePath, Stream documentStream);

        /// <summary>
        /// Checks whether a document exists at the specified path in the storage.
        /// </summary>
        /// <param name="filePath">The relative or absolute path identifying the document to check.</param>
        /// <returns>
        /// <see langword="true"/> if the document exists; otherwise, <see langword="false"/>.
        /// </returns>
        bool Exists(string filePath);
    }
}
