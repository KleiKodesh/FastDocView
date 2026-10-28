# Learnings

Append-only. Date, what we learned, what it changed.

## 2026-10-28 — first version

- **Word keeps no layout cache in the file.** `w:lastRenderedPageBreak` is written, never read
  back, and drifts with fonts/printer/version. Page-exact rendering needs Word (COM) or our own
  layout; "fast preview" therefore means flow rendering, no pages.
- **Deflate can't seek; zran can.** Checkpoints at block boundaries (compressed byte + leftover
  bits + 32 KB window) every ~1 MB cost ~3 % of the XML size and give ms seeks. .NET's
  `DeflateStream` can't be primed, so a managed inflater was written (~400 lines). It runs at
  ~40 MB/s — fine for a one-time index, slow for a full-document scan on every open.
- **zran is only worth it for big files.** For a 1 MB docx the index pass costs more than parsing
  the whole body. Threshold to be measured; expect tens of MB of uncompressed XML.
- **.doc already has random access** via the piece table; the only pass is the paragraph-mark
  scan (linear over text, no decompression). Formatting is looked up per paragraph through
  PlcfBtePapx/PlcfBteChpx → 512-byte FKP pages.
- **`.doc` extension lies.** Found MHTML ("MIME-Version") and RTF (`{\rtf`) saved as `.doc`.
  Detect by signature, never by extension.
- **`IsOn(null)` trap.** For `<w:b/>` a missing `w:val` means on; for `w:default` a missing
  attribute means *not* default. Applied blindly it made Heading1 the default paragraph style
  and Emphasis the default character style — everything rendered bold-italic-blue. Any
  attribute-presence check must be explicit.
- **Styles.xml is small, body is huge.** Load styles/numbering whole once; stream the body.
  Body-level block offsets + `w:pStyle` captured at index time give headings/TOC without reading
  blocks.
- **Namespaces for isolated blocks.** A block parsed alone needs the root element's opening tag
  (its xmlns declarations); the scanner captures it in the index pass.
- **Windowed numbering is unsolved.** A paragraph's list number depends on paragraphs never
  rendered. Needs an index-time counter per (numId, level) — cheap to add to the scanner.
- **Robocopy, not `mv`,** to move a folder that a shell has as cwd on Windows; and a stray
  `git init` in a subfolder silently breaks `git add -A` of the parent ("does not have a commit
  checked out").
- Measured (synthetic, 50 000 blocks, 22 MB XML): build 0.5 s, index 1.2 MB, reopen 50 ms,
  random 40-block window worst 30 ms, consecutive ~2 ms. Real Hebrew docx (250 KB): cold open
  120 ms, cached 9 ms, 30-block windows 6–34 ms.

## 2026-10-28 — policy, numbering, stream-based core

- **Indexing only pays above ~8 MB of body.** The one-pass open (scan + keep the inflated body)
  runs at ~40 MB/s; below 8 MB that is ≤ 200 ms and beats building, writing and loading a
  sidecar. Default threshold set there; a cached sidecar is still used below it.
- **List numbers must be counted in the same pass that finds the blocks.** Storing the finished
  label per block (string table) is simpler than storing counters and re-formatting: the label
  needs the whole counter vector for `%1.%2` patterns anyway.
- **`IsOn(null)` strikes twice.** Same class of bug would hit `w:lvlOverride`/`w:start` parsing
  — every "attribute present?" check is explicit now.
- **A stale test fixture hides a regression.** The 50k synthetic docx was reused from disk and
  had no numbering; the "lacks list marker" failures were the fixture, not the code. Fixture
  file names now carry a version suffix.
- **net48 has no `EnumerationOptions`.** Folder walks that must skip inaccessible directories
  need a hand-rolled stack walk when the library multi-targets.
- **WASM readiness is mostly about entry points.** Once the documents take a `Stream` and the
  index serialises to a `Stream`, nothing else in the core cares about files; the workload
  install is the remaining blocker.

## 2026-10-28 — real big files (found through the KitveiHakodesh service's local-file index)

- **The service index has paths, not sizes.** Listing the largest Word files = read the Lucene
  index read-only (`DirectoryReader` over `%LOCALAPPDATA%\KleiKodesh\KitveiHakodeshService\filesystemindex`,
  stored field `path`) and `FileInfo` each hit. 5,667 Word files on this laptop; a full-drive
  `Get-ChildItem -Recurse` never finished.
- Measured on real files: 40 MB `.doc` (107k paragraphs) cold 445 ms / reopen 175 ms / windows
  3–100 ms; 12 MB `.docx` with an **87 MB body** (123k blocks) index 5 s once, reopen 420 ms,
  windows 7–60 ms. Heap stays flat.
- **Block size, not body size, drives render cost.** A 15 MB docx with only 64 blocks (7 MB of
  drawing XML) takes 165–260 ms per window because each block's XML is parsed whole. Policy by
  body size alone misjudges such files; a per-block byte size is already in the index
  (start offsets) — a "largest block" signal should feed the decision and the renderer should
  skip `w:drawing` subtrees before parsing.
- `.doc` paragraphs: ~40 % of sampled blocks had no visible text (empty paragraphs / cell marks).
  Expected for table-heavy halachic compilations; reassembling `.doc` tables would fold those.
- OneDrive placeholder files throw "cloud file provider is not running" — surface as a clear
  "file is not downloaded" message rather than a failure.

## 2026-10-28 — live rig (C# localhost + Vite + the app's WindowedScroller)

- **Screenshots are out; the rig is the truth.** `vue-demo` runs the app's own `WindowedScroller`
  over `/api/documents/{id}/blocks`, Playwright reads `__wordDemo.state()` (rendered range,
  first visible, cache size, last fetch) and the console. First live run found what no harness
  or screenshot had: a jump into a run of empty `.doc` paragraphs (cell marks → zero-height
  rows) lands ~9 rows early (`scrollToIndex(104608)` → firstVisible 104599). Empty blocks need a
  minimum height in the row, or the `.doc` reader should fold cell-mark-only paragraphs.
- Live numbers through the browser: 40 MB `.doc` reopen 92 ms server / 630 ms wall (headings 250
  ms, weights 50 ms); 87 MB-body `.docx` reopen 710 ms server / 2.1 s wall — headings (1 s, reads
  every heading block) and the 123k-int weights array are now the biggest costs, not the index.
  Both belong in the index (heading text at scan time; weights as a binary endpoint or omitted).
- Block round trips are 15–90 ms, render 4–70 ms; the scroller keeps 36–105 rows mounted and the
  block cache stays in the low hundreds. Server heap 5–17 MB after GC with a document open.
- Playwright in a side project must pin the app's version (1.59.1) — its browsers are already in
  `%LOCALAPPDATA%\ms-playwright`; a newer package wants a download. Vite binds `localhost`, not
  `127.0.0.1`, and its ready line carries ANSI codes — don't grep for `Local:`.

## 2026-10-28 — notes, formatting audit, text-first open

- **Audit before implementing.** Streaming every part / every FKP run and counting unhandled
  constructs turned "full coverage" from guesswork into a 42-item (.docx) / ~80-item (.doc) list;
  most of it was pagination/proofing to ignore on purpose. Both now read 0 on the samples.
- **A desynced grpprl shows up as garbage opcodes** (0x0000, 0xFFFF, 0xD9D9…). The cause was
  `sprmTDefTable`'s 2-byte cb, which counts the operand PLUS ONE. Every sprm after it in a table
  row's PAPX was being misread — widths, borders, shading, all of it.
- **FIB arrays are longer than 0x200 bytes.** fcPlcfendRef/Txt are FibRgFcLcb97 entries 46/47
  (offset 0x212+); always bound by the file's own cbRgFcLcb.
- **.doc table rows must be one block.** Cell-mark paragraphs as rows produced runs of empty,
  zero-height rows the windowed scroller could not land on (jumps 9 rows early). Rows are now
  detected at index time (PAPX sprmPFInTable / sprmPFTtp) and rendered as one `<tr>`.
- **Notes ride with their block.** Delivering each block's note bodies in the same response keeps
  the notes pane exact for the visible text with no extra requests; numbering is index-time.
- **Text first.** Mounting the scroller before headings/weights and sending weights as raw Int32
  (not JSON) took reopen-to-text to 74 ms (.doc) / 147 ms (.docx 87 MB body) in the browser.
- **Drawing XML is skipped before parsing** (byte-level strip of w:drawing / w:pict / w:object):
  a drawing-heavy block no longer costs its DrawingML in XElement parsing.

## 2026-10-28 — notes streamed like the body

- **"Small part" is an assumption, not a fact.** footnotes.xml was 330 KB in the sample, so loading
  it whole looked harmless — but commentary-style seforim can have notes bigger than the text.
  The notes parts now get the same treatment as document.xml: scanned once (id → byte range),
  checkpointed, read one note at a time. Proven on a 113 MB footnotes part: heap +5 MB at open.
- **The policy must weigh every part a reader inflates**, not just document.xml — otherwise a
  small body with huge notes is opened "in memory" and inflates the notes whole.
- One package stream is shared by the block and note readers; their views re-seat the file
  position on every read, but concurrent requests still need a lock around rendering.

## 2026-10-28 — corpus audit over the whole laptop

- **Coverage is measured over every Word file on the machine** (5,355: 1,850 .docx, 3,339 .doc),
  listed from the KitveiHakodesh service's file index. First pass: 248 files with paragraph
  borders, 50 with text boxes (their text was being dropped), ~400 .doc with border sprms; 3 crashes.
  After the fixes: 0 unhandled in both formats (bar one revision construct, ignored on purpose), 0 crashes.
- **Drawings nest.** A picture inside a text box inside a shape: the drawing stripper's "these
  don't nest, first close tag wins" assumption broke the XML of such blocks. Matching is depth-aware
  now, and text boxes inside drawings are kept (their text is document text).
- **Table styles layer under paragraph styles** (defaults → table style → paragraph style →
  direct); cell margins / borders come from the table style first, the table's own tblPr last.
- **A .doc can have junk before its Pcdt, or no Clx at all** (non-complex file: text is one
  contiguous run from fcMin). Scanning forward for a plausible Pcdt and the contiguous fallback,
  as LibreOffice does, opened the last failing file.
- 145 files could not even be read (OneDrive placeholders / locked files): an OS state, not a
  parser bug — report it as "file not available", never as a format failure.

## 2026-10-28 — first-open profile: the scanner, not the inflater

- `--profile-index` on the 87 MB body (machine shared with two agent builds, so ±30 %): our
  inflater alone 0.5–0.9 s (100–170 MB/s), .NET's zlib 0.12–0.2 s (500–700 MB/s), the
  byte-level body scanner alone 1.8–2.5 s. **The scanner was the bottleneck**, at ~300 ns per
  tag: 4–6 state-machine iterations and a string allocation per tag.
- Fixes: text skipped with a vectorised `IndexOf('<')`, tag names compared as bytes against the
  dozen that matter, and a whole-tag fast path (one `IndexOf('>')` per tag, quote-parity check,
  fallback to the state machine for comments / CDATA / chunk-edge tags). Scanner now ~0.9–1.1 s.
- Inflater: 7-byte refills, mask instead of modulo, bulk + chunk-doubling match copies took it
  from ~35 to ~100+ MB/s. zlib is still 4–5× faster; a progressive first open (text before the
  index finishes) matters more than closing that gap, since reopen already reads no body.
## 2026-10-28 — theme fonts and colors

- **Word's Office theme has an empty `a:cs` face.** `cstheme="minorBidi"` only yields a Hebrew font
  through the supplemental `a:font script="Hebr"` entry, picked by `settings.xml`
  `w:themeFontLang/@bidi` (he-IL). Resolving `minorBidi` to `a:cs` alone — the obvious reading of
  the spec — gives "" and Hebrew falls back to Calibri. LibreOffice does the supplemental lookup
  first (ThemeHandler.cxx); so do we.
- **The spec's tint/shade formula is an image.** Every HTML rendering of ECMA-376 §17.3.2.6 loses
  it; the worked examples (C0504D tint 99 → D99694) pin it down as a per-component linear blend
  toward white / black. Verify formulas against the examples, not the prose.
- **Theme vs explicit value is spelled out per attribute** — theme wins for fonts and colors alike
  ("shall be ignored" / "is ignored"). LibreOffice still uses the explicit `val` for the character
  color (and keeps the theme reference beside it), because Word writes the precomputed value; we
  follow the spec when a theme part exists and fall back to `val` when it does not.

## 2026-10-28 — .doc list labels and fast-saved Prm

- **`lcbPlfLst` lies by design.** The LVL array that gives every list its levels follows the PlfLst
  but is not counted by its lcb; its size is only known by walking it (one LVL per simple list,
  nine otherwise, each with variable grpprls and an xst). Reading "offset .. offset + lcb" gets the
  LSTFs and nothing else.
- **A .doc list placeholder is a character whose value is the level.** The xst holds the number
  text with 16-bit level indexes at the offsets `rgbxchNums` names (one-based); there is no `%1`
  syntax. Bullets store the symbol-font code as 0xF0xx in the first char.
- **The .doc sample is 78 % list paragraphs** (siman / seif numbering), the .docx twin 81 % — so
  list labels are not a nicety for halachic compilations, they are most of the visible structure.
  Interning the labels (536 distinct for 83k items) keeps the index small; the .docx index still
  stores one string per item.
- **Counting lists in the index pass costs nothing measurable** (cold open 1.6–1.9 s with or
  without; the PAPX lookup per paragraph was already there for headings and table rows).
- **`git checkout -- <file>` to undo a one-line experiment undoes the whole file.** Lost three
  edits that way; use a targeted revert (sed back, or Edit) when the file carries uncommitted work.
- Prm0 isprm → sprm: the spec names the sprms (§2.9.215) but not their opcodes; each name was
  resolved against §2.6.1/§2.6.2 and cross-checked with LibreOffice `ww8scan.cxx` `aSprmId`
  (which uses sprmPJc80 0x2403 where the spec names sprmPJc 0x2461 — both are handled alike here).

## 2026-10-28 — every sprm opcode checked against sprmids.hxx

- After the 0x4845 slip, all 70 .doc opcodes the renderer interprets were looked up in LibreOffice's
  sw/source/filter/ww8/sprmids.hxx. Two more were wrong, both written from memory:
  0x0868 is CFUsePgsuSettings (we hid the text — 7,399 runs in 34 corpus files), web-hidden is 0x0811;
  0x245B is PFDyaBeforeAuto, contextual spacing is 0x246D. The other 68 matched.
- Lesson kept in CLAUDE.md: an opcode table is looked up, never recalled.

## 2026-10-28 — "why not zlib?" answered with measurements

- .NET's DeflateStream (zlib-ng since .NET 9) cannot checkpoint or resume: no inflatePrime /
  inflateSetDictionary / Z_BLOCK in the public API, and no proposal to add them was found.
- ZLibDotNet 0.1.1 (managed zlib 1.3.1 port, zlib license, dormant since 2024-03) has all three
  APIs and resumes correctly — but **fails with "invalid bit length repeat" on our real
  document.xml** and on zlib levels 1/6/9 output (only .NET "Fastest" output decodes), and is at best
  1–1.5× our inflater, slower on random reads. Rejected. Evaluation code kept on branch
  work/zlib-evaluation (b70a54e), not merged.
- Measured on the 91.6 MB XML: ours ~225–305 MB/s, .NET zlib ~830 MB/s, random 64 KB read ~2.2 ms.
  The path forward is our own inflater rebuilt on zlib's inflate_fast ideas (combined
  length/distance table entries, fewer bounds checks), one managed code path for every target.

## 2026-10-28 — coexisting with Word

- `File.OpenRead` / `FileShare.Read` deny other writers: with a document open in the reader, Word
  would open the file read-only or fail to save. Word saves by writing a temp file and replacing
  the original, so delete/rename must be shared too: every open is now `FileShare.ReadWrite |
  FileShare.Delete`, through one helper.
- Index-mode documents used to keep the file open for their whole life. Now the handle exists only
  while a window is read (released after every render) — measured cost: none visible.
- After Word saves, the old index's offsets point into a different file: rendering on would show
  wrong text silently. Renders now check size + write time first (microseconds) and throw
  `WordDocumentSourceChangedException`; the dev server answers 409 and the demo reopens in place.
  An in-memory .docx keeps showing the version it read; it still reports the change.
- A build in progress keeps its handle (shared) — if Word saves mid-build, the handle keeps reading
  the old file object, and the saved index fails its size/time check on the next open (re-index).
## 2026-10-28 - on-demand index for the first open

- **The scan is resumable for free.** The scanner was already chunk-fed and the inflater already
  pauses between symbols, so "index on demand" is a pull loop in the reader's thread, not a
  background pass. Checkpoints depend only on deflate block boundaries and output offsets, never on
  read sizes, so the pulled index is byte-identical to the one-go build.
- **A process-cold first render costs ~300-400 ms of JIT** (zip, XmlReader, the renderer) on top of
  the scan. Timings in the harness are taken after a small warm-up file; a real host is long-lived.
- **Heap "bounded" means "no more than a sidecar reopen".** A 1M-block index itself is ~45 MB of
  managed arrays/strings; comparing against 1/8 of the body was the wrong yardstick.
- Worktrees need `..\SeekableDeflate` next to them since the split: a junction
  `WordDocumentRandomAccess-work\SeekableDeflate -> ..\SeekableDeflate` makes them build.

## 2026-10-28 — on-demand index, live-checked

- A cold first open of the 87 MB-body sample shows text after ~0.26 s in the browser (server
  ~0.07 s) with ~440 of 123,007 blocks known; a jump to the middle pulls the scan there (~1.8 s);
  reaching the end completes and saves the index (identical to a full build).
- The live check caught what the harness could not: a jump to the *estimated* last row completed
  the index, the real count shrank (154,146 → 123,007), and the scroller mounted an empty window
  past the new end — a blank screen. Cause in the scroller's landing: its window was computed from
  the count before the rows were fetched. `landAt` now re-checks after the fetch and re-lands on the
  new last row. The app's own WindowedScroller has the same latent race when a count shrinks mid-landing.
