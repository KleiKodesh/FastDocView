using System.Text;
using FastDocView.DocBinary;
using FastDocView.Docx;
using FastDocView.Markup;
using FastDocView.Odt;
using FastDocView.Rtf;

namespace FastDocView;

/// <summary>
/// Plain text of a document, read through the in-document search's own text path (no rendering of the body, no second
/// extractor): the block scan of <see cref="IWordWindowedDocument.Search"/> and the notes scan of
/// <see cref="IWordWindowedDocument.SearchNotes"/> run with a capturing matcher
/// (<see cref="WordDocumentSearchMatcher.ForTextCapture"/>), so the text is exactly what the search reads - and what the
/// renderer shows. Blocks are read in bounded batches (the search's own windowed reads); only the result string grows.
/// <para>
/// Output: one line per paragraph / cell / line break as the search separates them; a block or note with no text is
/// left out. Footnotes and endnotes are one list in body order, as the notes pane and the notes search show them.
/// <see cref="GetTextWithNotes(string, CancellationToken)"/> puts a blank line between the body and the notes.
/// Hebrew marks and quote glyphs are kept as written (the search folds them only when it compares). No positions,
/// structure or metadata.
/// </para>
/// </summary>
public static class WordDocumentText
{
    private const int BlocksPerBatch = 1000;
    private const int NotesPerBatch = 200;

    /// <summary>The body text only (every format <see cref="WordWindowedDocumentOpener"/> opens).</summary>
    public static string GetText(string filePath, CancellationToken cancellation = default)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(filePath, cancellation: cancellation);
        return GetText(document, cancellation);
    }

    /// <summary>The footnotes and endnotes only, in body order.</summary>
    public static string GetNotes(string filePath, CancellationToken cancellation = default)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(filePath, cancellation: cancellation);
        return GetNotes(document, cancellation);
    }

    /// <summary>The body text, then a blank line, then the notes (only the body when there are no notes).</summary>
    public static string GetTextWithNotes(string filePath, CancellationToken cancellation = default)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(filePath, cancellation: cancellation);
        return GetTextWithNotes(document, cancellation);
    }

    /// <summary><see cref="GetText(string, CancellationToken)"/> of an open document.</summary>
    public static string GetText(IWordWindowedDocument document, CancellationToken cancellation = default)
    {
        Func<WordDocumentSearchMatcher, int, int, WordDocumentSearchBatch> scan = document switch
        {
            DocxWindowedDocument docx => (matcher, from, count) => docx.SearchWith(matcher, from, count, DocxWindowedDocument.DocxSearchTextMode.Bytes),
            OdtWindowedDocument odt => (matcher, from, count) => odt.SearchWith(matcher, from, count, OdtWindowedDocument.OdtSearchTextMode.Bytes),
            DocBinaryWindowedDocument doc => (matcher, from, count) => doc.SearchBlocksWith(matcher, from, count, onePass: true),
            MarkupWindowedDocument markup => markup.SearchWith,
            RtfWindowedDocument rtf => rtf.SearchWith,
            _ => throw new NotSupportedException($"No text reader for {document.GetType().Name}."),
        };
        return Collect(scan, BlocksPerBatch, cancellation);
    }

    /// <summary><see cref="GetNotes(string, CancellationToken)"/> of an open document.</summary>
    public static string GetNotes(IWordWindowedDocument document, CancellationToken cancellation = default)
    {
        // the notes are known only as far as the index reaches: finish it first (positions only, nothing rendered)
        while (!document.IsIndexComplete)
        {
            cancellation.ThrowIfCancellationRequested();
            document.ScanAhead(BlocksPerBatch);
        }
        return Collect((matcher, from, count) => WordDocumentNoteSearch.Search(document, matcher, from, count), NotesPerBatch, cancellation);
    }

    /// <summary><see cref="GetTextWithNotes(string, CancellationToken)"/> of an open document.</summary>
    public static string GetTextWithNotes(IWordWindowedDocument document, CancellationToken cancellation = default)
    {
        string body = GetText(document, cancellation);
        string notes = GetNotes(document, cancellation);
        return notes.Length == 0 ? body : body.Length == 0 ? notes : body + "\n\n" + notes;
    }

    /// <summary>Runs a search scan batch by batch with a capturing matcher; each hit's count names its text entry.</summary>
    private static string Collect(Func<WordDocumentSearchMatcher, int, int, WordDocumentSearchBatch> scan, int batchSize, CancellationToken cancellation)
    {
        var text = new StringBuilder();
        int from = 0;
        while (true)
        {
            cancellation.ThrowIfCancellationRequested();
            WordDocumentSearchMatcher matcher = WordDocumentSearchMatcher.ForTextCapture();   // per batch: entries are dropped once copied
            WordDocumentSearchBatch batch = scan(matcher, from, batchSize);
            foreach (int count in batch.HitCounts)
            {
                string entry = matcher.CapturedText(count);
                if (entry.Length == 0) continue;
                if (text.Length > 0) text.Append('\n');
                text.Append(entry);
            }
            if (batch.ReachedEnd || batch.NextBlock < 0) break;
            from = batch.NextBlock;
        }
        return text.ToString();
    }
}
