namespace FastDocView;

/// <summary>
/// What a document file says about itself - the values Word's File &gt; Info / Properties dialog shows (title, author,
/// dates, the statistics Word saved). Format-neutral: each reader fills what its format stores
/// (<see cref="IWordWindowedDocument.ReadProperties"/>). A value the file does not hold is null, never guessed or
/// computed from the text. The statistics are the ones the file was saved with (Word's count at its last save), not a
/// count of the text as shown.
/// </summary>
public sealed class WordDocumentProperties
{
    public static readonly WordDocumentProperties None = new();

    public string? Title { get; init; }
    public string? Subject { get; init; }
    /// <summary>Word's "Tags".</summary>
    public string? Keywords { get; init; }
    public string? Category { get; init; }
    /// <summary>Word's "Comments" (the description).</summary>
    public string? Comments { get; init; }
    /// <summary>Word's "Status".</summary>
    public string? ContentStatus { get; init; }
    public string? Author { get; init; }
    public string? LastModifiedBy { get; init; }
    public string? Company { get; init; }
    public string? Manager { get; init; }
    /// <summary>The revision number as the file stores it (a text in .docx / .doc).</summary>
    public string? Revision { get; init; }
    /// <summary>The program that saved the file ("Microsoft Office Word", "LibreOffice/7.6...").</summary>
    public string? Application { get; init; }
    public string? Template { get; init; }

    /// <summary>Dates: <see cref="DateTimeKind.Utc"/> when the file states the moment (a zone or a FILETIME),
    /// <see cref="DateTimeKind.Unspecified"/> when it stores only a wall-clock time (RTF, an .odt date without a zone).</summary>
    public DateTime? Created { get; init; }
    public DateTime? Modified { get; init; }
    public DateTime? LastPrinted { get; init; }

    public int? Pages { get; init; }
    public int? Words { get; init; }
    /// <summary>Characters without spaces (Word's "Characters (no spaces)").</summary>
    public int? Characters { get; init; }
    public int? CharactersWithSpaces { get; init; }
    public int? Lines { get; init; }
    public int? Paragraphs { get; init; }
    /// <summary>Total editing time in minutes.</summary>
    public int? TotalEditingMinutes { get; init; }

    /// <summary>A text value as a reader stores it: null for absent or blank (only spaces), trimmed otherwise.</summary>
    internal static string? Text(string? value) => string.IsNullOrWhiteSpace(value) ? null : value!.Trim();

    /// <summary>A non-negative integer stored as text (app.xml, meta:document-statistic); null when absent or not a number.</summary>
    internal static int? Count(string? value) =>
        int.TryParse(value?.Trim(), System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int count) ? count : null;
}
