using WordDocumentRandomAccess.DocBinary;
using WordDocumentRandomAccess.Docx;

namespace WordDocumentRandomAccess;

/// <summary>Opens a Word file by its signature (zip → .docx family, OLE compound file → .doc), not by extension.</summary>
public static class WordWindowedDocumentOpener
{
    public static IWordWindowedDocument Open(string filePath, string? cacheDirectory = null,
                                             IProgress<WordDocumentIndexProgress>? progress = null,
                                             CancellationToken cancellation = default)
    {
        var header = new byte[4];
        using (var file = File.OpenRead(filePath)) _ = file.Read(header, 0, 4);
        if (header[0] == (byte)'P' && header[1] == (byte)'K')
            return DocxWindowedDocument.Open(filePath, cacheDirectory, progress, cancellation);
        if (header[0] == 0xD0 && header[1] == 0xCF && header[2] == 0x11 && header[3] == 0xE0)
            return DocBinaryWindowedDocument.Open(filePath, cacheDirectory, progress, cancellation);
        if (header[0] == (byte)'M' && header[1] == (byte)'I' && header[2] == (byte)'M' && header[3] == (byte)'E')
            throw new NotSupportedException("This is a Word web archive (MHTML saved with a .doc extension), not a binary Word file.");
        if (header[0] == (byte)'{' && header[1] == (byte)'\\')
            throw new NotSupportedException("This is an RTF file saved with a Word extension; RTF is not supported.");
        throw new InvalidDataException("Not a Word document (neither a zip package nor an OLE compound file).");
    }
}
