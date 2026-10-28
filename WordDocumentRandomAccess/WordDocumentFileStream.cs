namespace WordDocumentRandomAccess;

/// <summary>
/// Read-only access to a document file that never gets in Word's way.
///   • Every open shares read, write and delete (<see cref="OpenShared"/>), so Word can open the
///     file for editing and save it — including its "write a temp file, rename over the original"
///     save — while we read it.
///   • A <see cref="WordDocumentFileStream"/> holds the OS handle only while it is being read:
///     the document calls <see cref="Release"/> after each window, so between windows the file is
///     not open at all. The next read reopens it transparently at the same position.
///   • <see cref="HasSourceChanged"/> compares the file's current size and write time with the ones
///     seen at open, so a host can tell that Word saved a new version and reopen (re-index) it.
/// </summary>
public sealed class WordDocumentFileStream : Stream
{
    private const int BufferSize = 1 << 16;

    private readonly long _lengthAtOpen;
    private readonly long _lastWriteTicksAtOpen;
    private FileStream? _file;
    private long _position;

    public string Path { get; }

    /// <summary>Opens a file for reading while letting any other process read, write, rename or delete it.</summary>
    public static FileStream OpenShared(string path)
        => new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, BufferSize);

    public WordDocumentFileStream(string path)
    {
        Path = path;
        var info = new FileInfo(path);
        if (!info.Exists) throw new FileNotFoundException("Document not found.", path);
        _lengthAtOpen = info.Length;
        _lastWriteTicksAtOpen = info.LastWriteTimeUtc.Ticks;
    }

    /// <summary>True when the file on disk is no longer the version this stream was opened on (or is gone).</summary>
    public bool HasSourceChanged
    {
        get
        {
            var info = new FileInfo(Path);
            return !info.Exists || info.Length != _lengthAtOpen || info.LastWriteTimeUtc.Ticks != _lastWriteTicksAtOpen;
        }
    }

    /// <summary>Closes the OS handle; the next read reopens the file. Safe to call at any time.</summary>
    public void Release()
    {
        _file?.Dispose();
        _file = null;
    }

    private FileStream File => _file ??= OpenShared(Path);

    public override int Read(byte[] buffer, int offset, int count)
    {
        FileStream file = File;
        if (file.Position != _position) file.Position = _position;
        int read = file.Read(buffer, offset, count);
        _position += read;
        return read;
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        _position = origin switch
        {
            SeekOrigin.Begin => offset,
            SeekOrigin.Current => _position + offset,
            _ => _lengthAtOpen + offset,
        };
        return _position;
    }

    public override bool CanRead => true;
    public override bool CanSeek => true;
    public override bool CanWrite => false;
    public override long Length => _lengthAtOpen;
    public override long Position { get => _position; set => _position = value; }
    public override void Flush() { }
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing)
    {
        if (disposing) Release();
        base.Dispose(disposing);
    }
}
