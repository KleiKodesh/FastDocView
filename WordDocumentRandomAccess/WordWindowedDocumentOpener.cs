using WordDocumentRandomAccess.DocBinary;
using WordDocumentRandomAccess.Docx;
using WordDocumentRandomAccess.Markup;
using WordDocumentRandomAccess.Odt;
using WordDocumentRandomAccess.Rtf;

namespace WordDocumentRandomAccess;

/// <summary>
/// File-system entry point. Picks the format by signature (zip → .docx family, or OpenDocument text
/// when its `mimetype` entry says so; OLE compound file → .doc; XML rooted in office:document → .fodt; `{\rtf` → RTF),
/// never by extension, and lets the <see cref="WordIndexPolicy"/> decide whether an index
/// is worth building. <see cref="Decide"/> answers the same question without opening the document,
/// for the host's "which files to index" picker.
/// </summary>
public static class WordWindowedDocumentOpener
{
    public static readonly string[] SupportedExtensions = [".docx", ".docm", ".dotx", ".dotm", ".doc", ".dot", ".odt", ".ott", ".fodt", ".mht", ".mhtml", ".txt", ".rtf"];

    public enum WordFileFormat { Docx, DocBinary, WebArchive, RichText, OpenDocument, FlatOpenDocument, Unknown, PlainText }

    public static WordFileFormat DetectFormat(string filePath)
    {
        var header = new byte[4];
        using FileStream file = WordDocumentFileStream.OpenShared(filePath);
        _ = file.Read(header, 0, 4);
        if (Path.GetExtension(filePath).Equals(".txt", StringComparison.OrdinalIgnoreCase)) return WordFileFormat.PlainText;   // any bytes are text; no signature applies
        if (header[0] == (byte)'P' && header[1] == (byte)'K')
            return OdtPackage.IsOpenDocumentTextPackage(file) ? WordFileFormat.OpenDocument : WordFileFormat.Docx;
        if ((header[0] == (byte)'<' || (header[0] == 0xEF && header[1] == 0xBB && header[2] == 0xBF)) && OdtPackage.IsFlatOpenDocumentText(file))
            return WordFileFormat.FlatOpenDocument;
        if (header[0] == 0xD0 && header[1] == 0xCF && header[2] == 0x11 && header[3] == 0xE0) return WordFileFormat.DocBinary;
        if (header[0] == (byte)'M' && header[1] == (byte)'I' && header[2] == (byte)'M' && header[3] == (byte)'E') return WordFileFormat.WebArchive;
        if (header[0] == (byte)'{' && header[1] == (byte)'\\') return WordFileFormat.RichText;
        string extension = Path.GetExtension(filePath).ToLowerInvariant();
        if (extension is ".mht" or ".mhtml") return WordFileFormat.WebArchive;        // archives may start with other headers (From:, Content-Type:)
        return WordFileFormat.Unknown;
    }

    public static IWordWindowedDocument Open(string filePath, WordIndexPolicy? policy = null,
                                             IProgress<WordDocumentIndexProgress>? progress = null,
                                             CancellationToken cancellation = default)
    {
        try { return OpenByFormat(filePath, policy, progress, cancellation); }
        catch (IOException exception) when (exception is not WordDocumentSourceChangedException && IsNotFullyLocal(filePath))
        {
            throw OnlineOnly(filePath, exception);
        }
    }

    private static IWordWindowedDocument OpenByFormat(string filePath, WordIndexPolicy? policy,
                                                      IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
    {
        return DetectFormat(filePath) switch
        {
            WordFileFormat.Docx => DocxWindowedDocument.Open(filePath, policy, progress, cancellation),
            WordFileFormat.DocBinary => DocBinaryWindowedDocument.Open(filePath, policy, progress, cancellation),
            WordFileFormat.OpenDocument => OdtWindowedDocument.Open(filePath, isFlat: false, policy, progress, cancellation),
            WordFileFormat.FlatOpenDocument => OdtWindowedDocument.Open(filePath, isFlat: true, policy, progress, cancellation),
            WordFileFormat.WebArchive => MarkupWindowedDocument.Open(filePath, MarkupSourceKind.WebArchive, policy, progress, cancellation),
            WordFileFormat.PlainText => MarkupWindowedDocument.Open(filePath, MarkupSourceKind.PlainText, policy, progress, cancellation),
            WordFileFormat.RichText => RtfWindowedDocument.Open(filePath, policy, progress, cancellation),
            _ => throw DescribeUnknown(filePath),
        };
    }

    // Windows file attributes of a cloud placeholder (WinNT.h, learn.microsoft.com "File Attribute Constants"):
    // RECALL_ON_DATA_ACCESS 0x00400000 (not all data is local), RECALL_ON_OPEN 0x00040000, OFFLINE 0x00001000.
    private const int NotFullyLocalAttributes = 0x00400000 | 0x00040000 | 0x00001000;

    /// <summary>True for a file whose data is (partly) in the cloud only, e.g. OneDrive "Files On-Demand".</summary>
    private static bool IsNotFullyLocal(string filePath)
    {
        try { return ((int)File.GetAttributes(filePath) & NotFullyLocalAttributes) != 0; }
        catch (IOException) { return false; }
        catch (UnauthorizedAccessException) { return false; }
    }

    private static IOException OnlineOnly(string filePath, IOException cause) => new(
        $"\"{Path.GetFileName(filePath)}\" is stored online only (for example OneDrive Files On-Demand) and could not be downloaded: "
        + $"{cause.Message} Make it available offline, or start the cloud app, then open it again.", cause);

    /// <summary>
    /// The error for a file no reader takes, naming what it is when its first bytes tell: Word 2003
    /// XML, HTML, or a Word for Windows 1.x / 2.0 file (a FIB with no compound file around it; wIdent
    /// 0xA59B/0xA59C = Word 1, 0xA5DB = Word 2 per LibreOffice ww8scan.cxx WW8Fib::GetFIBVersion).
    /// </summary>
    private static Exception DescribeUnknown(string filePath)
    {
        var head = new byte[4096];
        int count;
        using (FileStream file = WordDocumentFileStream.OpenShared(filePath)) count = file.Read(head, 0, head.Length);
        ushort identifier = count >= 2 ? BitConverter.ToUInt16(head, 0) : (ushort)0;
        if (identifier is 0xA59B or 0xA59C) return new NotSupportedException("Word for Windows 1.x documents are not supported.");
        if (identifier == 0xA5DB) return new NotSupportedException("Word for Windows 2.0 documents are not supported.");
        string text = System.Text.Encoding.UTF8.GetString(head, 0, count);
        if (text.Contains("schemas.microsoft.com/office/word/2003/wordml"))
            return new NotSupportedException("Word 2003 XML documents (WordprocessingML 2003) are not supported.");
        if (text.IndexOf("<html", StringComparison.OrdinalIgnoreCase) >= 0 || text.IndexOf("<!DOCTYPE html", StringComparison.OrdinalIgnoreCase) >= 0)
            return new NotSupportedException("HTML pages saved with a Word extension are not supported.");
        return new InvalidDataException("Not a Word, RTF, text or OpenDocument file (no known signature).");
    }

    /// <summary>What opening this file would do under the policy — reads only headers, no index pass.</summary>
    public static WordIndexDecision Decide(string filePath, WordIndexPolicy? policy = null)
    {
        policy ??= WordIndexPolicy.Default;
        try { return DecideByFormat(filePath, policy); }
        catch (IOException exception) when (IsNotFullyLocal(filePath)) { throw OnlineOnly(filePath, exception); }
    }

    private static WordIndexDecision DecideByFormat(string filePath, WordIndexPolicy policy)
    {
        return DetectFormat(filePath) switch
        {
            WordFileFormat.Docx => DocxWindowedDocument.Decide(filePath, policy),
            WordFileFormat.DocBinary => DocBinaryWindowedDocument.Decide(filePath, policy),
            WordFileFormat.OpenDocument => OdtWindowedDocument.Decide(filePath, policy, isFlat: false),
            WordFileFormat.FlatOpenDocument => OdtWindowedDocument.Decide(filePath, policy, isFlat: true),
            WordFileFormat.WebArchive => MarkupWindowedDocument.Decide(filePath, policy, MarkupSourceKind.WebArchive),
            WordFileFormat.PlainText => MarkupWindowedDocument.Decide(filePath, policy, MarkupSourceKind.PlainText),
            WordFileFormat.RichText => RtfWindowedDocument.Decide(filePath, policy),
            var format => new WordIndexDecision(filePath, WordIndexMode.InMemory, false, 0, 0, $"unsupported ({format})"),
        };
    }

    /// <summary>Decisions for every Word file under a folder — the host shows this list and lets the user pick what to index.</summary>
    public static IEnumerable<WordIndexDecision> DecideFolder(string folderPath, WordIndexPolicy? policy = null, bool recursive = true, int maxFiles = 2000)
    {
        int count = 0;
        foreach (string path in EnumerateFilesSafely(folderPath, recursive))
        {
            if (!SupportedExtensions.Contains(Path.GetExtension(path).ToLowerInvariant())) continue;
            if (Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal)) continue;      // Word lock files
            if (++count > maxFiles) yield break;
            WordIndexDecision decision;
            try { decision = Decide(path, policy); }
            catch (Exception exception) when (exception is IOException or InvalidDataException or NotSupportedException)
            {
                decision = new WordIndexDecision(path, WordIndexMode.InMemory, false, 0, 0, $"cannot read: {exception.Message}");
            }
            yield return decision;
        }
    }

    /// <summary>Directory walk that skips folders it may not read (net48 has no EnumerationOptions.IgnoreInaccessible).</summary>
    private static IEnumerable<string> EnumerateFilesSafely(string folderPath, bool recursive)
    {
        var pending = new Stack<string>();
        pending.Push(folderPath);
        while (pending.Count > 0)
        {
            string folder = pending.Pop();
            string[] files;
            try { files = Directory.GetFiles(folder); }
            catch (Exception exception) when (exception is UnauthorizedAccessException or IOException) { continue; }
            foreach (string file in files) yield return file;
            if (!recursive) continue;
            string[] subfolders;
            try { subfolders = Directory.GetDirectories(folder); }
            catch (Exception exception) when (exception is UnauthorizedAccessException or IOException) { continue; }
            for (int i = subfolders.Length - 1; i >= 0; i--) pending.Push(subfolders[i]);
        }
    }

    /// <summary>Builds the sidecar index now for files the policy marks as worth it (the host's "index these" action).</summary>
    public static void BuildSidecarIndexes(IEnumerable<WordIndexDecision> decisions, WordIndexPolicy? policy = null,
                                           IProgress<(string file, WordDocumentIndexProgress progress)>? progress = null,
                                           CancellationToken cancellation = default)
    {
        policy ??= WordIndexPolicy.Default;
        foreach (WordIndexDecision decision in decisions)
        {
            if (decision.Mode != WordIndexMode.Sidecar || decision.HasCachedIndex) continue;
            cancellation.ThrowIfCancellationRequested();
            var fileProgress = progress is null ? null : new Progress<WordDocumentIndexProgress>(p => progress.Report((decision.FilePath, p)));
            using IWordWindowedDocument document = Open(decision.FilePath, policy, fileProgress, cancellation);
        }
    }
}
