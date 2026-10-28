namespace WordDocumentRandomAccess;

/// <summary>
/// Thrown by RenderBlocks when the file changed on disk since it was opened (Word saved it): the
/// index's offsets belong to the old version, so reading on would show wrong text. The host
/// reopens the document, which re-indexes it, and keeps the reader's position.
/// </summary>
public sealed class WordDocumentSourceChangedException : IOException
{
    public string FilePath { get; }

    public WordDocumentSourceChangedException(string filePath)
        : base($"The document changed on disk since it was opened: {filePath}") => FilePath = filePath;
}
