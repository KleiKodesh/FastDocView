using System.Security.Cryptography;
using System.Text;

namespace FastDocView;

/// <summary>
/// Where a document's sidecar index lives. Never next to the document (read-only folders, network
/// shares): under %LOCALAPPDATA% keyed by the full path. The index itself also records the file's
/// size and last-write time, so a changed document is detected and re-indexed.
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
        string fileName = Path.GetFileNameWithoutExtension(fullPath);
        if (fileName.Length > 40) fileName = fileName.Substring(0, 40);
        foreach (char invalid in Path.GetInvalidFileNameChars()) fileName = fileName.Replace(invalid, '_');
        return Path.Combine(cacheDirectory ?? DefaultCacheDirectory, $"{fileName}.{hash}.wordindex");
    }
}
