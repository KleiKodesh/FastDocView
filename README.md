# WordDocumentRandomAccess

Random access into Word files for instant windowed HTML rendering of huge documents — no
decompressed copy, no page layout.

| Project | What |
|---|---|
| `WordDocumentRandomAccess/` | Library (net48 + net10.0, zero dependencies). `WordWindowedDocumentOpener.Open(path)` → `IWordWindowedDocument` with `RenderBlocks(first, count)`. |
| `WordDocumentRandomAccess.Tests/` | Console harness: deflate round-trip + random seeks, synthetic 50k-block docx, real-file smoke run. |
| `WordDocumentRandomAccess.DevServer/` | C# localhost JSON host (`http://127.0.0.1:5188/api/…`), the dev counterpart of the app's service. |
| `vue-demo/` | **The live rig**: Vite + Vue using the app's own `WindowedScroller` (copied), folder picker with index decisions, reload, server memory readout. `npm run dev` spawns the C# host. |
| `WordDocumentRandomAccess.Viewer/` | Old WinForms + WebView2 viewer — superseded by `vue-demo`, kept only until the Vue rig covers everything. |

```
cd WordDocumentRandomAccess.Tests
dotnet run -c Release -- --real "C:\path\book.docx" "C:\path\old.doc"      # harness (+ --huge 1000000)

cd ../vue-demo
npm install
npm run dev                                   # http://localhost:5189 — starts the C# host too
node playwright/liveCheck.mjs "C:\path\book.docx" [--headed]   # live DOM/console check, no screenshots
```

Design, measurements, limitations and the specifications used: [NOTES.md](NOTES.md).
