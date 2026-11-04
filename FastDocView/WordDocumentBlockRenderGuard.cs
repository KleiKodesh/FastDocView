namespace FastDocView;

/// <summary>
/// Keeps one damaged block from failing its whole window. A window renders in one pass as before; only when that
/// pass throws a content error is the window rendered again block by block, and each block that still fails becomes
/// the usual "unsupported" placeholder. A changed / closed / unreadable file (any <see cref="IOException"/> but a
/// truncated stream), cancellation and out-of-memory are never caught: the host must see those.
/// Each failed block is reported once per document through <see cref="BlockRenderFailed"/>.
/// </summary>
public sealed class WordDocumentBlockRenderGuard(string filePath)
{
    public const string PlaceholderHtml = "<p><span class=\"unsupported\">[this part of the document could not be read]</span></p>";

    /// <summary>Raised once per document and block that could not be rendered: file path, block index, cause.</summary>
    public static event Action<string, int, Exception>? BlockRenderFailed;

    private readonly HashSet<int> _reportedBlocks = [];

    public static bool IsContentError(Exception exception) => exception switch
    {
        EndOfStreamException => true,
        IOException or UnauthorizedAccessException or OperationCanceledException or ObjectDisposedException
            or OutOfMemoryException or StackOverflowException or ThreadAbortException => false,
        _ => true,
    };

    /// <param name="renderLocked">The document's own window render (called under its read lock).</param>
    /// <param name="kindOf">The block's kind from the index, for the placeholder.</param>
    public IReadOnlyList<WordDocumentBlockHtml> Render(int firstBlockIndex, int count, int blockCount,
        Func<int, int, IReadOnlyList<WordDocumentBlockHtml>> renderLocked, Func<int, WordDocumentBlockKind> kindOf)
    {
        try { return renderLocked(firstBlockIndex, count); }
        catch (Exception exception) when (IsContentError(exception) && count > 1) { }
        catch (Exception exception) when (IsContentError(exception)) { return [Placeholder(firstBlockIndex, exception, kindOf)]; }

        int first = Math.Max(0, firstBlockIndex);
        int end = (int)Math.Min(blockCount, (long)firstBlockIndex + count);
        var result = new List<WordDocumentBlockHtml>(Math.Max(0, end - first));
        for (int blockIndex = first; blockIndex < end; blockIndex++)
        {
            try { result.AddRange(renderLocked(blockIndex, 1)); }
            catch (Exception exception) when (IsContentError(exception)) { result.Add(Placeholder(blockIndex, exception, kindOf)); }
        }
        return result;
    }

    private WordDocumentBlockHtml Placeholder(int blockIndex, Exception cause, Func<int, WordDocumentBlockKind> kindOf)
    {
        if (_reportedBlocks.Add(blockIndex)) BlockRenderFailed?.Invoke(filePath, blockIndex, cause);
        WordDocumentBlockKind kind;
        try { kind = kindOf(blockIndex); }
        catch (Exception exception) when (IsContentError(exception)) { kind = WordDocumentBlockKind.Paragraph; }
        // a table row placeholder would break the row grouping on the page: show it as a paragraph
        if (kind is WordDocumentBlockKind.TableRow or WordDocumentBlockKind.Table) kind = WordDocumentBlockKind.Paragraph;
        return new WordDocumentBlockHtml(blockIndex, kind, PlaceholderHtml, []);
    }
}
