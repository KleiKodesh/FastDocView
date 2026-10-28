# FastDocView

Standalone library — its own **public** repo (github.com/KleiKodesh/FastDocView),
sibling folder of KitveiHakodesh; nothing of it lives in the KitveiHakodesh repo.

**Goal: instant preview of .doc and .docx files.** Any window of paragraphs rendered to HTML in
milliseconds, for files of any size, inside the app's book-view (lines pane + windowed scroller +
TOC from headings + notes). **Speed is the goal** (user, 2026-10-28): instant loading beats
formatting coverage when the two compete.

Read first: [REFERENCES.md](REFERENCES.md) (the authoritative sources, in order — Word's source is
not available; Microsoft's Open Specifications are the closest), [NOTES.md](NOTES.md) (design,
measurements, limitations) and
[LEARNINGS.md](LEARNINGS.md) (what we found out the hard way — append, never rewrite history).
Naming and style follow `KitveiHakodesh/.kiro/steering/` (full words, siblings share a prefix,
one concept one name, no `utils`, a name must match its contents).

## The paradigm — never break it

**VITAL (user, 2026-10-28): data is loaded WINDOWED, always. Never, ever cross that line — it is the
whole point of FastDocView and what makes it fast.** Only the blocks near the viewport are read,
parsed, rendered and sent; the document is never loaded, streamed or sent whole — not on a first
open, not "for a moment", not because a file is small, not to the page. A first open builds its
index while the reader reads windows from the partial index; the index-building call returns only
counts, never content. Any design that needs the whole document at once is wrong — find the
index-time or per-window form instead.

One **streaming index pass** records everything a reader needs to jump anywhere; afterwards only
the blocks near the viewport are read, parsed and rendered — the same model as the app's lines
view over SQLite.

- .docx: zran-style deflate checkpoints (own managed inflater) over document.xml **and**
  footnotes.xml / endnotes.xml; per-block byte offsets, kinds, styles, list labels, headings,
  note references, text lengths. Notes are read on demand like blocks.
- .doc: the piece table already gives random access; the index pass records blocks (paragraphs,
  whole table rows), headings; formatting comes from FKP pages on demand.
- **Never hold whole-document text or XML** (body or notes) in memory — bounded caches only.
  A part being "small" in one sample is not a reason to load it whole.
- Anything that needs the whole document (numbering, note numbers, SEQ fields, table row counts,
  open fields at a block start) is computed **at index time** and stored per block.
- Techniques from other projects (LibreOffice, Open-Xml-PowerTools, Mammoth, POI, pandoc) are
  adopted only when they fit index-time or render-time-per-block; whole-document DOM passes are
  rejected or rewritten as per-block resolvers + caches.

## Direction (user, 2026-10-28)

- **Full multiplatform library**: net48 host (VSTO), native AOT service, **WASM**. Core is
  `Stream` in, pure managed: no P/Invoke, no reflection, no `System.Drawing`; `IsAotCompatible` on.
- **Index intelligently**: `WordIndexPolicy` — small files open by one in-memory pass, big ones
  get a sidecar; the policy weighs body + notes; the host can force a mode or pick files/folders.
- The final reader **looks like today's book-view**; the block API (`index` + `html` + `notes`)
  mirrors `seforimApi.getLinesPaged` on purpose. Integration into the app comes after speed.
- Footnotes/endnotes follow **Word's continuous (draft) view**: references inline, bodies
  delivered with their block, shown in a notes pane / tooltip — never page-bottom layout.
- **Current priority order** (speed first): show text before indexing finishes; faster managed
  inflater; cache parsed styles/numbering with the index; fewer round trips to first paint;
  background pre-indexing. Formatting work (conditional table styles, themes, full numbering
  rules, fields, CSS class interning, fast-saved / Word 6 .doc) follows — see NOTES.md.

## Layout

| Folder | Owns |
|---|---|
| `..\SeekableDeflate` (own repo, sibling folder) | managed RFC 1951 inflater with zran checkpoints, seekable stream, zip-entry location — referenced as a `ProjectReference`, no NuGet. Clone it next to this repo. |
| `..\TextEncodingDetection` (own repo, sibling folder) | .txt encoding detection identical to Notepad++ (ported Notepad++ uchardet) — `ProjectReference`, no NuGet. Clone it next to this repo (worktrees: next to them, in `..\FastDocView-work`). |
| `…/Docx` | body + notes scanners, index (format v4) + builder, block / note readers, drawing stripper, `DocxWindowedDocument` |
| `…/Html`, `…/DocxHtml` | format-neutral formatting models + CSS; styles / table styles cascade; block, table, note → HTML |
| `…/Odt` | OpenDocument text (.odt/.ott/.fodt): package detection, body scanner, list numbering, index + builder, style cascade, block reader + renderer, `OdtWindowedDocument` |
| `…/Markup` | .txt read as HTML (one block per line, encoding by TextEncodingDetection = Notepad++) and Word web archives (MHTML: streaming quoted-printable / base64 with per-block resume points); scanner, renderer, index + builder, `MarkupWindowedDocument` |
| `…/Rtf` | RTF (.rtf): streaming parser (group-state stack, control-word table), index + builder (byte ranges + interned start-state snapshots), header tables, list numbering, block renderer, `RtfWindowedDocument` |
| `…/DocBinary` | compound file, FIB, piece table, FKPs (+ huge PAPX), styles, fonts, sprms, notes, block index, `DocBinaryWindowedDocument` |
| `samples/` | the two real test files (דור המלקטים .doc 40 MB, .docx 87 MB body) — the demo's default |
| `FastDocView.Tests` | console harness; must print ALL PASSED. `--real`, `--audit`, `--odt [N]`, `--text [N]`, `--mhtml [N]`, `--text-corpus <list>`, `--mime-scan <list>`, `--rtf [N]` (with `--corpus <list>`: the RTF corpus pass), `--rtf-audit`, `--rtf-profile`, `--corpus <list>`, `--huge N`, `--notes-streaming N` |
| `FastDocView.DevServer` | C# localhost JSON host on 5188 (dev only; the real host will be the KitveiHakodesh service) |
| `vue-demo` | **the live rig** — Vite on 5189 + the app's `WindowedScroller` (copied from `KitveiHakodesh/vue-frontend/src/components`, keep in sync — the copy carries one fix to port back: `landAt` re-lands on the new last row when fetching rows shrinks the count); `window.__wordDemo` hooks; `playwright/` checks |
| `FastDocView.Viewer` | old WinForms viewer — superseded, do not extend |

## Verifying

- **Live only, never screenshots**: `npm run dev` in `vue-demo` (spawns the dev server), then
  Playwright against `localhost:5189` — `liveCheck.mjs` (jumps, bounded window, reload keeps the
  same block at the same pixel), `notesCheck.mjs` (numbering, notes pane, click, hover). Read DOM,
  console and `__wordDemo.state()`.
- **Third-party stress documents** live in `test-data/` (LibreOffice + Apache POI + Apache Tika, every format we read, licenses inside):
  `--corpus ../test-data` runs all ~5,000 — part of QA for any reader/renderer change. 0 crashes;
  broken/encrypted files must fail with a clear error.
- **Coverage follows real files**: `--audit` on the samples and `--corpus <list>` over every Word
  file on the machine (list exported from the KitveiHakodesh service's file index — a full-drive
  search is too slow). Keep unhandled at 0; anything ignored on purpose goes in the audit's
  ignored-by-design set with a reason. The corpus pass also opens + renders each file: 0 crashes.
- Speed claims need numbers: server open/reopen ms, text-shown ms in the browser, window ms, heap.
- Synthetic fixtures carry a version suffix in their file name; bump it when the generator changes.

## Rules

- **Never get in Word's way** (user, 2026-10-28). Every file open goes through
  `WordDocumentFileStream` (`OpenShared` = read access, sharing read + write + delete) — never
  `File.OpenRead` / `FileShare.Read`, which block Word from saving. A document holds no handle
  between windows (`Release()` after each render), notices a save (`HasSourceChanged`) and refuses
  to render the new file with the old index (`WordDocumentSourceChangedException`) — the host
  reopens. `WordCoexistenceTest` (exclusive open + save-by-replace while we read) must pass.

- **Always double-check against the source; never rely on memory alone** (user, 2026-10-28).
  Format facts (sprm opcodes, element semantics, FIB offsets, LibreOffice behaviour) must be
  verified in a source from REFERENCES.md before they are implemented — for Word-specific .docx
  behaviour check [MS-OI29500] first; when sources disagree or are silent, test in Word itself — a remembered
  opcode already cost us silently dropped raised/lowered text (0x4840 vs 0x4845). Research
  reports must say which facts were read from source and which came from memory; the
  from-memory ones are checked before use.

- Every new format feature needs a harness check (synthetic or real file) and a NOTES.md update.
- Long runs report progress per step (`IProgress<WordDocumentIndexProgress>`, every 100 files in
  corpus mode) and stay light: below-normal priority, one file at a time, bounded memory.
- Default sidecar threshold is 8 MB of text (body + notes); change it with a measurement.
- Demo UI stays plain — it is a rig, not the product.
- Clean up what you start (dev server, Vite, Word automation instances) — by PID, never by name.
- **QA before each commit** (user, 2026-10-28): before every commit run the harness
  (`dotnet run -c Release -- --real samples/…doc samples/…docx --audit <both>`), and when the
  change touches rendering or the demo, the Playwright checks against the live rig. Then have a
  fresh subagent review the staged diff for correctness (not style) and fix what it finds. A
  commit that was not QA'd is not pushed.
- **Work in parallel where pieces are independent** (user, 2026-10-28): separate formats,
  speed vs coverage, research vs implementation go to parallel subagents with narrow briefs;
  only sequential, shared-file work stays in the main thread. **At most 2 agents at a time**
  in a session (user, 2026-10-28) — the machine must stay usable; queue the rest. Agents work
  in their own worktree under `..\FastDocView-work\<topic>` (branch `work/<topic>`),
  commit there, never push or merge; the main thread merges after QA.
- **Commit as you go** (user, 2026-10-28): small commits after each verified step (a feature,
  a fix, a measured speed-up), not one big commit at the end — each QA'd per the rule above.
- **No speed hacks — no threading to hide slowness** (user, 2026-10-28): if something is slow,
  the implementation is bad; fix the algorithm / data structure / allocations. No extra threads,
  producer–consumer pipelines or parallel loops to mask a slow pass. (A two-thread inflate/scan
  hand-off was tried and reverted for this reason.)
- **Test runs take a few minutes at most** (user, 2026-10-28): iterate and QA on samples — the files
  relevant to the change plus a fixed-seed random sample (~300) of `test-data/` / real files. Never
  block on full-corpus runs.
- **Book view is NEVER FastDocView** (user, 2026-10-28): the app's book view (library books) and the
  FastDocView page (local document files) are different things — own feature folder, route, stores and
  names. FastDocView only borrows the book view's design and the shared components (WindowedScroller,
  SplitPane, TreeView, useZoom…). Never call this work "book-view integration".
- **No hacks — fix the bug** (user, 2026-10-28): when a test file breaks the system, fix the cause,
  never paper over it with a cap, limit, cut-off or fallback that hides the input. Stress files
  (5,000 nested tables, huge lines, deep tags) exist to prove the design holds. Example: deep
  nesting overflowing the stack is fixed by walking with an explicit stack, not by flattening
  content past depth N. Limits are only legitimate when the format or Word itself defines them
  (cite the source).
- **A failed build stops the commit**: never chain `build | grep … && git commit` — grep
  succeeds even when the build fails. Check the build's exit status, both targets (net48 + net10).
- Commits: author KleiKodesh, **no Co-Authored-By trailer**; push freely to this repo.
- Don't build or verify unless asked; when asked, run the harness and the Playwright checks.
