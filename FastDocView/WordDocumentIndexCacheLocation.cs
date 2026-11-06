using System.Security.Cryptography;
using System.Text;

namespace FastDocView;

/// <summary>
/// Where a document's sidecar index lives. Never next to the document (read-only folders, network
/// shares): under %LOCALAPPDATA% keyed by the full path. The index itself also records the file's
/// size and last-write time, so a changed document is detected and re-indexed. The folder is kept bounded by
/// <see cref="WordDocumentIndexCache"/> (LRU, caps in <see cref="WordDocumentIndexCache.Limits"/>).
/// </summary>
public static class WordDocumentIndexCacheLocation
{
    public static string DefaultCacheDirectory { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "KitveiHakodesh", "FastDocView");

    public static string IndexPathFor(string documentPath, string? cacheDirectory = null)
    {
        string fullPath = Path.GetFullPath(documentPath);
        string hash;
        using (var sha = SHA256.Create())
        {
            byte[] digest = sha.ComputeHash(Encoding.UTF8.GetBytes(fullPath.ToUpperInvariant()));
            var sb = new StringBuilder(32);
            for (int i = 0; i < 16; i++) sb.Append(digest[i].ToString("x2"));
            hash = sb.ToString();
        }
        // The hash of the full path alone: unique per document, no name cleaning; the record beside it names the document.
        string indexPath = Path.Combine(cacheDirectory ?? DefaultCacheDirectory, hash + WordDocumentIndexCache.SidecarExtension);
        WordDocumentIndexCache.RememberDocument(indexPath, fullPath);     // its record names the document (stale-sidecar cleanup)
        return indexPath;
    }
}
