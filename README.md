# WordDocumentRandomAccess

Random access into Word files for instant windowed HTML rendering of huge documents — no
decompressed copy, no page layout.

| Project | What |
|---|---|
| `WordDocumentRandomAccess/` | Library (net48 + net10.0, zero dependencies). `WordWindowedDocumentOpener.Open(path)` → `IWordWindowedDocument` with `RenderBlocks(first, count)`. |
| `WordDocumentRandomAccess.Tests/` | Console harness: deflate round-trip + random seeks, synthetic 50k-block docx, real-file smoke run. |
| `WordDocumentRandomAccess.Viewer/` | WinForms + WebView2 windowed viewer (test stand-in for the app's Vue scroller). |

```
cd WordDocumentRandomAccess.Tests
dotnet run -c Release -- --real "C:\path\book.docx" "C:\path\old.doc"
cd ../WordDocumentRandomAccess.Viewer
dotnet run -c Release -- "C:\path\book.docx"            # add --screenshot out.png to capture and exit
```

Design, measurements, limitations and the specifications used: [NOTES.md](NOTES.md).
