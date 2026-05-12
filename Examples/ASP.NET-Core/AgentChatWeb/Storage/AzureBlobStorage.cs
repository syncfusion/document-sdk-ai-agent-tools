using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using Syncfusion.AI.AgentTools.Core;

namespace AgentChatApp.Storage;

/// <summary>
/// Azure Blob Storage implementation of <see cref="IDocumentStorage"/>.
/// All document I/O flows through the configured blob container.
/// </summary>
public sealed class AzureBlobStorage : IDocumentStorage
{
    private readonly BlobContainerClient _container;

    public AzureBlobStorage(BlobContainerClient container)
    {
        ArgumentNullException.ThrowIfNull(container);
        _container = container;
        // Ensure the container exists (no-op if it already does).
        _container.CreateIfNotExists(PublicAccessType.None);
    }

    /// <inheritdoc/>
    public Stream Read(string filePath)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        var blob = _container.GetBlobClient(filePath);
        var ms = new MemoryStream();
        blob.DownloadTo(ms);
        ms.Position = 0;
        return ms;
    }

    /// <inheritdoc/>
    public bool Write(string filePath, Stream documentStream)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        ArgumentNullException.ThrowIfNull(documentStream);
        try
        {
            documentStream.Position = 0;
            var blob = _container.GetBlobClient(filePath);
            blob.Upload(documentStream, overwrite: true);
            return true;
        }
        catch
        {
            return false;
        }
    }



    /// <inheritdoc/>
    public bool Exists(string filePath)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        return _container.GetBlobClient(filePath).Exists();
    }

    /// <summary>
    /// Returns metadata for all blobs: name, size, content type, and last modified date.
    /// Used by the Files API to populate the Documents folder listing.
    /// Supports folder hierarchy using blob name prefixes (e.g., "Input/", "Output/").
    /// </summary>
    public IReadOnlyList<BlobItemInfo> GetAllBlobItems()
    {
        return _container.GetBlobs(BlobTraits.Metadata)
            .Select(b => new BlobItemInfo(
                b.Name,
                b.Properties.ContentLength ?? 0,
                b.Properties.ContentType ?? "application/octet-stream",
                b.Properties.LastModified?.LocalDateTime ?? DateTime.MinValue))
            .OrderBy(b => b.Name)
            .ToList();
    }

    /// <summary>
    /// Returns a hierarchical structure of folders and files in the blob container.
    /// Folders are identified by the presence of blobs with matching prefixes.
    /// </summary>
    public FileSystemStructure GetFileSystemStructure()
    {
        var blobs = _container.GetBlobs(BlobTraits.Metadata).ToList();
        var root = new FileSystemStructure();

        // Group blobs by their folder structure
        foreach (var blob in blobs)
        {
            var parts = blob.Name.Split('/');
            
            if (parts.Length == 1)
            {
                // File at root level
                root.Files.Add(new BlobItemInfo(
                    blob.Name,
                    blob.Properties.ContentLength ?? 0,
                    blob.Properties.ContentType ?? "application/octet-stream",
                    blob.Properties.LastModified?.LocalDateTime ?? DateTime.MinValue));
            }
            else
            {
                // File in a folder
                var folderName = parts[0];
                if (!root.Folders.ContainsKey(folderName))
                {
                    root.Folders[folderName] = new FolderInfo
                    {
                        Name = folderName,
                        Files = new List<BlobItemInfo>()
                    };
                }

                root.Folders[folderName].Files.Add(new BlobItemInfo(
                    blob.Name,
                    blob.Properties.ContentLength ?? 0,
                    blob.Properties.ContentType ?? "application/octet-stream",
                    blob.Properties.LastModified?.LocalDateTime ?? DateTime.MinValue));
            }
        }

        return root;
    }

    /// <summary>
    /// Uploads a stream as a new blob with the given file path (blob name).
    /// Used by the Files API upload endpoint.
    /// Supports folder paths (e.g., "Input/template.docx").
    /// </summary>
    public void Upload(string filePath, Stream stream, string contentType)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        ArgumentNullException.ThrowIfNull(stream);
        stream.Position = 0;
        var options = new BlobUploadOptions
        {
            HttpHeaders = new BlobHttpHeaders { ContentType = contentType }
        };
        _container.GetBlobClient(filePath).Upload(stream, options);
    }

    /// <summary>
    /// Downloads the blob as a byte array for serving in an HTTP response.
    /// Supports folder paths (e.g., "Input/template.docx").
    /// </summary>
    public byte[] Download(string filePath)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        var ms = new MemoryStream();
        _container.GetBlobClient(filePath).DownloadTo(ms);
        return ms.ToArray();
    }
    
    /// <inheritdoc/>
    internal bool Delete(string filePath)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        try
        {
            var blob = _container.GetBlobClient(filePath);
            return blob.DeleteIfExists();
        }
        catch (Exception ex)
        {
            // Log the exception for debugging
            System.Diagnostics.Debug.WriteLine($"Failed to delete blob '{filePath}': {ex.Message}");
            return false;
        }
    }
}

/// <summary>Lightweight blob metadata DTO returned by <see cref="AzureBlobStorage.GetAllBlobItems"/>.</summary>
public sealed record BlobItemInfo(string Name, long Size, string ContentType, DateTime LastModified);

/// <summary>Represents the file system structure with folders and files.</summary>
public sealed class FileSystemStructure
{
    public Dictionary<string, FolderInfo> Folders { get; set; } = new();
    public List<BlobItemInfo> Files { get; set; } = new();
}

/// <summary>Represents a folder with its contained files.</summary>
public sealed class FolderInfo
{
    public string Name { get; set; } = string.Empty;
    public List<BlobItemInfo> Files { get; set; } = new();
}
