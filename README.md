# FastDocView

Random access into Word files for instant windowed HTML rendering of huge documents — no
decompressed copy, no page layout.

| Project | What |
|---|---|
| `SeekableDeflate/` | **Separate repo, a submodule of this one.** Seekable deflate + zip-entry location the library builds on. |
| `FastDocView/` | Library (net48 + net10.0, zero dependencies). `WordWindowedDocumentOpener.Open(path)` → `IWordWindowedDocument` with `RenderBlocks(first, count)`. |
| `FastDocView.Tests/` | Console harness: deflate round-trip + random seeks, synthetic 50k-block docx, real-file smoke run. |
| `FastDocView.DevServer/` | C# localhost JSON host (`http://127.0.0.1:5188/api/…`), the dev counterpart of the app's service. |
| `vue-demo/` | **The live rig**: Vite + Vue using the app's own `WindowedScroller` (copied), folder picker with index decisions, reload, server memory readout. `npm run dev` spawns the C# host. |
| `FastDocView.Viewer/` | Old WinForms + WebView2 viewer — superseded by `vue-demo`, kept only until the Vue rig covers everything. |

```
cd FastDocView.Tests
dotnet run -c Release -- --real "C:\path\book.docx" "C:\path\old.doc"      # harness (+ --huge 1000000)

cd ../vue-demo
npm install
npm run dev                                   # http://localhost:5189 — the app's own FastDocView page; starts the C# host too
node playwright/appCheck.mjs "C:\path\book.docx" [--headed]   # live DOM/console check, no screenshots
```

## Using it

One file:

```csharp
using var document = WordWindowedDocumentOpener.Open(path);
IReadOnlyList<WordDocumentBlockHtml> window = document.RenderBlocks(0, 50);   // HTML of blocks 0-49, nothing else read
```

A host (service, dev server) that keeps documents open by id for a page:

```csharp
using var documents = new WordWindowedDocumentSet(cacheDirectory);   // locks per document, idle close, sidecar cache
string id = documents.Open(path);
var window = documents.Use(id, document => document.RenderBlocks(first, count));
documents.Use(id, document => document.ScanFor(5000, TimeSpan.FromMilliseconds(20)));   // keep building a first-open index
// WordDocumentClosedException from Use = closed, idle-closed or changed on disk: reopen
```

Plain text (C# only; no positions or structure), read through the search's own text path, so it is exactly what
the search finds and the view shows:

```csharp
string body  = WordDocumentText.GetText(path, cancellation);           // body only
string notes = WordDocumentText.GetNotes(path, cancellation);          // footnotes + endnotes, one list in body order
string all   = WordDocumentText.GetTextWithNotes(path, cancellation);  // body, a blank line, the notes
```

One line per paragraph / cell / line break; blocks without text are left out. Overloads take an open
`IWordWindowedDocument`. Details: NOTES.md, "Plain-text API".

`FastDocView.DevServer/Program.cs` is a complete host in ~150 lines.

Design, measurements, limitations and the specifications used: [NOTES.md](NOTES.md).

## Feature support

What each format renders, verified in the code: [docs/SUPPORTED-FEATURES.md](docs/SUPPORTED-FEATURES.md).
What is missing or partial, with the reason and status: [docs/UNSUPPORTED-FEATURES.md](docs/UNSUPPORTED-FEATURES.md).
