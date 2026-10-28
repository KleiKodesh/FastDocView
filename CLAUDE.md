# WordDocumentRandomAccess

Standalone library — its own **public** repo (github.com/KleiKodesh/WordDocumentRandomAccess),
sibling folder of KitveiHakodesh; nothing of it lives in the KitveiHakodesh repo.

**Goal: instant preview of .doc and .docx files.** Any window of paragraphs rendered to HTML in
milliseconds, for files of any size, inside the app's book-view (lines pane + windowed scroller +
TOC from headings + notes). **Speed is the goal** (user, 2026-10-28): instant loading beats
formatting coverage when the two compete.

Read first: [NOTES.md](NOTES.md) (design, measurements, limitations, specs) and
[LEARNINGS.md](LEARNINGS.md) (what we found out the hard way — append, never rewrite history).
Naming and style follow `KitveiHakodesh/.kiro/steering/` (full words, siblings share a prefix,
one concept one name, no `utils`, a name must match its contents).

## The paradigm — never break it

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
| `…/Docx` | body + notes scanners, index (format v4) + builder, block / note readers, drawing stripper, `DocxWindowedDocument` |
| `…/Html`, `…/DocxHtml` | format-neutral formatting models + CSS; styles / table styles cascade; block, table, note → HTML |
| `…/DocBinary` | compound file, FIB, piece table, FKPs (+ huge PAPX), styles, fonts, sprms, notes, block index, `DocBinaryWindowedDocument` |
| `samples/` | the two real test files (דור המלקטים .doc 40 MB, .docx 87 MB body) — the demo's default |
| `WordDocumentRandomAccess.Tests` | console harness; must print ALL PASSED. `--real`, `--audit`, `--corpus <list>`, `--huge N`, `--notes-streaming N` |
| `WordDocumentRandomAccess.DevServer` | C# localhost JSON host on 5188 (dev only; the real host will be the KitveiHakodesh service) |
| `vue-demo` | **the live rig** — Vite on 5189 + the app's `WindowedScroller` (copied from `KitveiHakodesh/vue-frontend/src/components`, keep in sync); `window.__wordDemo` hooks; `playwright/` checks |
| `WordDocumentRandomAccess.Viewer` | old WinForms viewer — superseded, do not extend |

## Verifying

- **Live only, never screenshots**: `npm run dev` in `vue-demo` (spawns the dev server), then
  Playwright against `localhost:5189` — `liveCheck.mjs` (jumps, bounded window, reload keeps the
  same block at the same pixel), `notesCheck.mjs` (numbering, notes pane, click, hover). Read DOM,
  console and `__wordDemo.state()`.
- **Coverage follows real files**: `--audit` on the samples and `--corpus <list>` over every Word
  file on the machine (list exported from the KitveiHakodesh service's file index — a full-drive
  search is too slow). Keep unhandled at 0; anything ignored on purpose goes in the audit's
  ignored-by-design set with a reason. The corpus pass also opens + renders each file: 0 crashes.
- Speed claims need numbers: server open/reopen ms, text-shown ms in the browser, window ms, heap.
- Synthetic fixtures carry a version suffix in their file name; bump it when the generator changes.

## Rules

- **Always double-check against the source; never rely on memory alone** (user, 2026-10-28).
  Format facts (sprm opcodes, element semantics, FIB offsets, LibreOffice behaviour) must be
  verified in the spec or the real upstream code before they are implemented — a remembered
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
  in their own worktree under `..\WordDocumentRandomAccess-work\<topic>` (branch `work/<topic>`),
  commit there, never push or merge; the main thread merges after QA.
- **Commit as you go** (user, 2026-10-28): small commits after each verified step (a feature,
  a fix, a measured speed-up), not one big commit at the end — each QA'd per the rule above.
- **No speed hacks — no threading to hide slowness** (user, 2026-10-28): if something is slow,
  the implementation is bad; fix the algorithm / data structure / allocations. No extra threads,
  producer–consumer pipelines or parallel loops to mask a slow pass. (A two-thread inflate/scan
  hand-off was tried and reverted for this reason.)
- **A failed build stops the commit**: never chain `build | grep … && git commit` — grep
  succeeds even when the build fails. Check the build's exit status, both targets (net48 + net10).
- Commits: author KleiKodesh, **no Co-Authored-By trailer**; push freely to this repo.
- Don't build or verify unless asked; when asked, run the harness and the Playwright checks.
