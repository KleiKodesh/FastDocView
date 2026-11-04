namespace FastDocView;

/// <summary>Why a document could not be opened, in terms a reader understands. A host maps each value to its own message.</summary>
public enum WordDocumentOpenFailure
{
    /// <summary>Anything not classified below (a bug or an unexpected system error).</summary>
    Unknown,
    /// <summary>Password-protected (encrypted) file.</summary>
    Encrypted,
    /// <summary>The file's bytes are damaged or not what its format promises.</summary>
    Corrupt,
    /// <summary>A cloud placeholder (e.g. OneDrive Files On-Demand) that could not be downloaded.</summary>
    OnlineOnly,
    /// <summary>The file or its folder does not exist (moved, renamed, deleted).</summary>
    NotFound,
    /// <summary>A format or format version the library does not read.</summary>
    UnsupportedFormat,
    /// <summary>The user may not read the file.</summary>
    AccessDenied,
}

/// <summary>A password-protected document. Still a <see cref="NotSupportedException"/> for callers that catch that.</summary>
public sealed class WordDocumentEncryptedException(string message) : NotSupportedException(message);

/// <summary>A cloud placeholder whose data could not be downloaded.</summary>
public sealed class WordDocumentOnlineOnlyException(string message, Exception cause) : IOException(message, cause);

/// <summary>Maps an exception thrown by <see cref="WordWindowedDocumentOpener.Open"/> (or by a later read) to a <see cref="WordDocumentOpenFailure"/>.</summary>
public static class WordDocumentOpenFailures
{
    public static WordDocumentOpenFailure Classify(Exception exception) => exception switch
    {
        WordDocumentEncryptedException => WordDocumentOpenFailure.Encrypted,
        WordDocumentOnlineOnlyException => WordDocumentOpenFailure.OnlineOnly,
        FileNotFoundException or DirectoryNotFoundException => WordDocumentOpenFailure.NotFound,
        UnauthorizedAccessException or System.Security.SecurityException => WordDocumentOpenFailure.AccessDenied,
        NotSupportedException => WordDocumentOpenFailure.UnsupportedFormat,
        InvalidDataException or EndOfStreamException or System.Xml.XmlException or FormatException
            or IndexOutOfRangeException or ArgumentOutOfRangeException or OverflowException => WordDocumentOpenFailure.Corrupt,
        _ => WordDocumentOpenFailure.Unknown,
    };
}
