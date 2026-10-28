# WordDocumentRandomAccess

Standalone library (own git, sibling of KitveiHakodesh). **Stated goal: fast preview of .doc and
.docx files** — any window of paragraphs rendered to HTML in milliseconds, without a decompressed
copy, inside the app's book-view (lines pane + windowed scroller + TOC from headings).

Read first: [NOTES.md](NOTES.md) (design, measurements, limitations, specs) and
[LEARNINGS.md](LEARNINGS.md) (what we found out the hard way — append, never rewrite history).
Naming and style follow `KitveiHakodesh/.kiro/steering/` (full words, siblings share a prefix,
one concept one name, no `utils`).

## Direction (user, 2026-10-28)

- A **full library, multiplatform**: net48 host (VSTO), native AOT service, **WASM** in the
  browser. Keep the core free of file-system assumptions (`Stream` in, everything else pure
  managed), no P/Invoke, no reflection, no `System.Drawing`. `IsAotCompatible` stays on.
- **Index intelligently**: the user chooses which files/folders are indexed; index only when it
  gains something. Small files (zran below ~50 MB of XML is overkill) should open by a plain
  one-pass parse; big files get the sidecar index. The decision belongs in the library, exposed as
  a policy the host can override.
- The final reader **looks like today's book-view**; the block API (`index` + `html`) mirrors
  `seforimApi.getLinesPaged` on purpose. The WinForms/WebView2 viewer is a test stand-in only.

## Layout

| Folder | Owns |
|---|---|
| `WordDocumentRandomAccess/Deflate` | managed RFC 1951 inflater with zran checkpoints, seekable stream |
| `…/Zip` | central directory → entry data offsets (Zip64) |
| `…/Docx` | body-block scanner, index model/file, builder, block reader, `DocxWindowedDocument` |
| `…/Html`, `…/DocxHtml` | format-neutral formatting models + CSS; styles.xml cascade; block → HTML |
| `…/DocBinary` | compound file, FIB, piece table, FKPs, style sheet, sprms, `DocBinaryWindowedDocument` |
| `WordDocumentRandomAccess.Tests` | console harness (`dotnet run -c Release -- --real <files>`); must print ALL PASSED |
| `WordDocumentRandomAccess.DevServer` | C# localhost JSON host on 5188 (dev only; the real host is the KitveiHakodesh service) |
| `vue-demo` | **the live rig** — Vite on 5189 + the app's `WindowedScroller` (copied from `KitveiHakodesh/vue-frontend/src/components`, keep in sync); `window.__wordDemo` hooks; `playwright/liveCheck.mjs` |
| `WordDocumentRandomAccess.Viewer` | old WinForms viewer — superseded, do not extend |

**Debugging is live only**: `npm run dev` in `vue-demo`, then Playwright against `localhost:5189`
(DOM + console + `__wordDemo.state()`). Never screenshots. Demo UI stays plain — it is a rig,
not the product; the product is the app's book-view.

## Rules

- Every new format feature needs a check in the test harness (synthetic docx or a real file) and a
  line in NOTES.md "Deliberately not done yet" removed or added.
- Never hold whole-document text or XML in memory; bounded caches only (same rule as the app's
  windowed data).
- Long runs report progress per step (`IProgress<WordDocumentIndexProgress>`).
- Default sidecar threshold is 8 MB of body (`WordIndexPolicy`); change it with a measurement, not a guess.
- Formatting coverage follows `--audit` on real files: keep both audits at 0 unhandled; anything
  ignored on purpose goes in the audit's ignored-by-design set with a reason.
- Notes follow Word's draft view: references inline, bodies delivered with their block.
- Don't build or verify unless asked; when asked, run the harness and the viewer screenshot.
