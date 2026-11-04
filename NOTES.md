# FastDocView — design notes

Goal: open huge Word files (.docx and old binary .doc) and render any window of paragraphs to
HTML in milliseconds, without ever holding a decompressed copy of the document, so the app's
windowed (virtual-scroll) view can show them instantly. No page layout: paragraphs flow, page
breaks are a thin rule.

## Why random access is hard for .docx

`word/document.xml` is one deflate-compressed zip entry. Deflate has no seek: every byte may
back-reference the previous 32 KB of *output*, so reaching paragraph 50 000 normally means
inflating everything before it. Word itself keeps no layout cache in the file
(`w:lastRenderedPageBreak` is write-only metadata; see conversation 2026-10-28).

## The zran technique (Deflate/)

Mark Adler's `zlib/examples/zran.c` (v1.7, 2025 — supports raw deflate):

1. One sequential pass inflates the stream. At deflate **block boundaries** (the only clean
   resume points) roughly every 1 MB of output we save a **checkpoint**: compressed byte
   position + leftover bits of the previous byte (blocks are not byte aligned), uncompressed
   offset, and the last 32 KB of output (the window later back-references copy from).
2. To read at uncompressed offset X: take the last checkpoint ≤ X, seed the inflater with its
   window, re-seat the bit reader, inflate forward ≤ 1 MB, return the bytes.

Cost: ~32 KB per checkpoint → ~3 % of the uncompressed XML. A 1 GB document.xml → ~32 MB index.
Seeks are a few ms (the inflater runs ~40–60 MB/s managed).

`DeflateCheckpointInflater` is a from-scratch managed RFC 1951 inflater (no zlib, no
SharpZipLib) because .NET's `DeflateStream` cannot be primed with a window or started at a bit
offset. It pauses between symbols, so a read can stop at any output byte, and it fires
`BlockBoundaryReached` where a checkpoint is valid. `DeflateHuffmanDecoder` uses a 10-bit fast
table plus puff.c's canonical walk for longer codes.

Prior art looked at: `GzipSeekable` (NuGet 0.1.0, managed, gzip-only wrapper), `indexed_gzip`
(Python), `jzran` (Java), `zindex`. None handles a raw-deflate zip entry directly, and the glue
(zip offsets + paragraph map) is the real work, so the inflater was written here.

**Not zlib:** .NET's zlib (DeflateStream) has no checkpoint/resume API, and the one managed
zlib port that has it (ZLibDotNet 0.1.1) mis-decodes our real files and is not faster — see
LEARNINGS 2026-10-28. The inflater stays ours; its speed work follows zlib's inflate_fast design.

## The .docx index (Docx/)

`DocxRandomAccessIndexBuilder` does the one pass. While inflating it feeds every chunk to
`DocxBodyBlockScanner`, a byte-level XML state machine (no DOM) that finds every child of
`<w:body>` (**block** = paragraph / table / content control / other) with its offset, kind,
`w:pStyle`, list label (counted here, `DocxListCounter`) and text length, plus headings, note
references, sections, bookmarks and the fields open at each block. The root element's opening
tag (its namespace declarations wrap each block so it can be parsed in isolation,
`DocxBlockXmlReader`) is kept too.

### Compact sidecar (format v13, 2026-11-02)

The sidecar holds **no per-block table**. What it stores:

- **Checkpoints** (`DeflateCheckpointTable`, SeekableDeflate): one per ~2 MB of inflated XML,
  snapped by `DeflateCheckpointSelector` to the deflate block boundary of the region's first half
  that needs the fewest window bytes (every boundary there is captured as a candidate, sealed 32 KB
  later with the count of window bytes the following output copied; a history-free boundary — 0
  bytes — is taken at once). Positions as varint deltas; all sparse windows (unreferenced bytes
  zero, history-free windows dropped) in one `DeflateWindowStore`: groups of 8 windows, each group
  one Brotli stream (net10; Deflate on net48, which has no Brotli in the BCL — the codec byte is
  stored, a Brotli store is unreadable on net48 and the index is rebuilt there with Deflate, which
  both read). Sample: 37 checkpoints, 17 history-free, 24 KB.
- **Anchors** (`DocxIndexAnchor`): one at the first block start past every 256 KB of output
  (`AnchorSpanBytes`, independent of the checkpoints — a rescan seeks to the nearest checkpoint
  anyway): the block's index + offset, the block before it (offset, kind, style — a window starting
  at an anchor's block reads its neighbour without scanning backwards) and the **list-numbering
  state** (`DocxListCounter.ExportState`: every list's counters + applied start overrides) as a
  delta from the previous anchor; in memory a full state is kept for every 8th anchor. Sample:
  348 anchors, 14 KB.
- Text-length **totals per 64 blocks** (scroll weights: group average until the span is scanned,
  exact afterwards), headings / note references / bookmarks with their blocks as **Elias-Fano**
  (`EliasFanoSequence`, ~2 + log2(n/count) bits per block index), sections (width table + byte
  ids), open-field stacks and string tables deflated.

Block i → position / kind / style / label / text length: `DocxCompactBlockTable` finds the span,
seeds a fresh `DocxBodyBlockScanner` at the anchor (`SeedAtBodyLevel`, list counter restored from
the anchor's state) and scans forward only as far as asked, keeping the span scanner for later
requests (4 spans cached, LRU). The same scanner and counter as the index pass ⇒ the same values
by construction (CompactIndexTest compares random + span-edge blocks against a full reference
scan; the golden dump of every block's HTML before/after the change was byte-identical for the
123 007-block sample). The bytes inflated while locating a window stay in a bounded ring
(256 KB, grows to 4 MB for a search batch), so the window renders from them without a second
inflate; sequential windows stay one inflate pass.

Measured, dor-hamelaktim-shabbat.docx (91.6 MB document.xml, 123 007 blocks), Release net10:

| | per-block layout (v12) | compact (v13), 2026-11-02 | compact after the speed work, 2026-11-03 |
|---|---|---|---|
| sidecar | 1 173 069 B | 59 326 B (24 KB windows, 14 KB anchors, 5 KB bookmarks, 4.4 KB weights, 3.5 KB headings, 3 KB sections, 3 KB rest) | 59 326 B (unchanged) |
| reopen from sidecar | 138 ms | 146 ms | ~145 ms |
| first window after reopen | 82 ms | 194 ms (numbering.xml parsed whole on the first window) | 60–70 ms (only the lists the window uses are parsed) |
| random single block, 200 probes | median 3.3 ms, p90 5.7 | median 7.9 ms, p90 14.7 | median 5.4 ms, p90 9.2 (same-moment base: 6.8 / 10.6) |
| sequential window of 40 | — | ~+2 ms over the render itself | ~+1 ms |
| full sequential render, 123k blocks | 13.0 s | 18.4 s | ~16.4 s (within this box's noise of the 2026-11-02 base) |
| search, 2 queries over all blocks | 2.6 / 2.7 s | 3.4 / 2.7 s | 2.4 / 2.8 s |

What the 2026-11-03 work found and fixed (golden dump of every block byte-identical throughout; the
machine was shared with another agent's builds, so the table's wall times are ± 20 %):

- **The retained-bytes ring restarted on every read that did not continue it**, so after the first
  miss every window and every search batch seeked back to a checkpoint and inflated its bytes a
  second time: 608 of 615 windows of a full render, 25 of 25 search batches (instrumented). Now
  `ReadBlockBytes` collects the bytes while the span scan locates the window's last block, whatever
  the length (a 5000-block search batch included) — one inflate pass — and the ring only holds the
  scan's own chunks. Every anchor's full list state is kept once computed (was: up to seven deltas
  replayed per span).
- **numbering.xml parsed whole on the first window** (1.15 MB, 464 lists; ~75 ms quiet, 150+ loaded):
  `DocxNumberingPartSlices` cuts the part into its top-level elements by bytes, and a num /
  abstractNum is parsed the first time a list asks for it.
- **The "95 MB/s inflater" was the tiered JIT's tier-0 code**: with tiering off the old loop ran
  at ~350–450 MB/s, the same as the rewritten one (SeekableDeflate: inflate_fast / libdeflate
  design, one linear output buffer, word copies). The rewrite is kept because it runs at that
  speed from its first call (`AggressiveOptimization`), where the old loop spent a process's
  first tens of MB — a reopen's first windows — four times slower. .NET's zlib does ~770 MB/s;
  the remaining lever is to feed the span after a checkpoint to it as a synthetic stream (a
  stored block holding the window + the compressed bytes re-packed by the leftover-bit shift).
- Scanner label path: the nine level definitions of a list resolved once per list (was nine
  dictionary walks per list paragraph), the labeller's CSS / pPr caches keyed by reference (a
  record's value hash hashed its strings per paragraph), style / numId / ilvl read from the
  attribute bytes (no string per tag). Scan of 24 MB with labels: 244 → ~200 ms; the scan is
  ~1 s of a full render, so this is small.

A random block now costs: inflate from the checkpoint (~1.2 MB at ~400 MB/s ≈ 3 ms) + rescan
from the anchor (≤ 256 KB at ~120 MB/s ≈ 1 ms) + render. Only a faster inflater (native zlib via
the synthetic stream, or a yet faster managed loop) moves the median much further.

The `.doc`, `.odt`, `.rtf` and markup (.txt / MHTML) indexes follow the same principle (anchors +
rescanning, no per-block table) since 2026-11-03; see each format's section.

The index is saved to `%LOCALAPPDATA%\KitveiHakodesh\FastDocView\<name>.<hash>.wordindex`
keyed by path, validated by file length + last write time. Cold build runs at about 40 MB/s of
XML with a few MB of RAM.

## When indexing is worth it (WordIndexPolicy)

The checkpoint index only pays off for big bodies. `WordIndexPolicy` decides per file from the
uncompressed body size (read from the zip central directory / the FIB, no index pass):

- **InMemory** (default below 8 MB of body): one pass that scans *and* keeps the inflated body
  for the session; nothing is written. ~40 MB/s, so ≤ 200 ms. A 4 MB Hebrew docx opens in
  ~250 ms this way.
- **Sidecar** (≥ 8 MB, or forced, or already cached): checkpoints + sidecar file; re-open in
  tens of ms whatever the size.

The host overrides the threshold / forces a mode (`ForcedMode`) and gets the decision with its
reason (`WordIndexDecision`). `WordWindowedDocumentOpener.DecideFolder` lists every Word file
under a folder with its decision for a picker; `BuildSidecarIndexes` builds the chosen ones with
per-file progress. The demo viewer shows this under File → Open folder.

### Sidecar cache folder: LRU (2026-11-03)

`WordDocumentIndexCache` keeps `%LOCALAPPDATA%\KitveiHakodesh\FastDocView` (or the policy's `CacheDirectory`) bounded;
it used to grow without end.
- **Record**: next to `X.wordindex` a tiny `X.wordindex.source` (document full path, length, last-write ticks at save
  time). The document comes from `IndexPathFor` (it remembers the last 64 index paths it computed), so no caller changed.
- **Last use** = the sidecar's last-write time: set on every save and on every open whose `TryLoad` validates it (one
  file-time write; the sidecar's content does not depend on it). NTFS last-access is often off, so not used.
- **Atomic writes**: every format's `Save` now goes through `WordDocumentIndexCache.Save`: a temp file unique to the
  writer (`X.wordindex.<pid>-<random>.tmp`, `FileShare.None`), then one rename over the old file (net10
  `File.Move(overwrite)`, net48 `File.Replace`). Before, all writers shared `X.wordindex.tmp` and did delete + move
  (two instances could clobber each other's temp; a reader could see no file). A failed save (disk full, old file held)
  leaves no temp file and is a cache miss, not an error (.docx / .doc callers no longer see the exception).
- **Trim** (`Trim(directory, limits)`, also callable by the host): temp files and records with no sidecar older than an
  hour; sidecars whose document is gone (only when its folder still exists: an unplugged drive or share keeps its
  sidecars) or whose length / time differ from the record; then least recently used first until both caps hold.
  Record-less (older) sidecars take part in the LRU only.
- **Caps** (`WordDocumentIndexCache.Limits`, a `WordIndexCacheLimits`; the host or Core sets it, so a
  user setting can replace them later): default 16 MB and 500 sidecars (decided 2026-11-03; was 256 MB / 4,000, too
  generous). Measured sidecars: compact formats 3-60 KB (the .docx sample 59 KB, the 40 MB .doc sample 16-17 KB), so
  16 MB holds ~270 of the largest compact ones and the byte cap can bind before the count cap; losing an old sidecar
  only costs one slower first open.
- **When**: never on the caller's thread. Once per process per folder on the first use of the cache (a validated
  open or a save), and again after a save at least 10 minutes after the last trim; on a background thread at lowest
  priority. `TrimsAutomatically = false` turns it off (the browser build has no threads: it catches that and skips).
- **Concurrency** (the service, a second app instance): a delete of a file another process holds open (readers
  `File.ReadAllBytes`, writers' temp files) fails on the sharing violation and is skipped (`SkippedInUseCount`); a
  sidecar whose last-write time moved since the listing (used again) is kept; files that vanish mid-trim are ignored.
  A reader that loses the race to a delete sees a missing file = a cache miss, rebuilt.
- **Pending delete** (2026-11-03): a delete of a file another handle holds with delete sharing (MarkUsed's
  SetLastWriteTimeUtc, a virus scanner, the indexer; every delete on a file system without POSIX delete) leaves the name
  until that handle closes, and any open, delete or rename onto it is access denied - an `UnauthorizedAccessException`,
  not an `IOException`. Every format's `TryLoad` but RTF's let it out of `File.ReadAllBytes`; all now read through
  `WordDocumentIndexCache.TryReadSidecar` (null = a cache miss). A save whose rename onto the old sidecar fails (held,
  pending delete) retries once after 1 ms, then is a cache miss. Test: `PendingDelete` in `--index-cache` makes the state
  deterministically (a delete-sharing handle + a delete-on-close open) and checks every read, save, mark, trim and open.
- Tests: `--index-cache` (also in the default run): eviction order, count cap, size cap, stale (changed, deleted,
  folder missing, no record), orphan temps / records, a sidecar held open, 3 s of saves + reads racing ~5,600 trims
  (0 exceptions, 0 torn reads), the real open path (record written, reopen marks use, change -> stale).

## Multiplatform / WASM

The core takes `Stream`s (`DocxWindowedDocument.OpenFromStream`, `DocBinaryWindowedDocument.
OpenFromStream`, `DocxRandomAccessIndex.Write/Read`) and has no P/Invoke, reflection or drawing,
so it compiles for `browser-wasm` unchanged; file-system conveniences (`Open(path)`, sidecar
cache folder) live only in the opener/documents' path overloads. A browser project is **not in
the repo yet**: the `wasm-tools` workload is not installed on this machine (installing writes to
Program Files → elevation). Next step once installed: a `net10.0-browser` project exposing
`open(bytes)` / `renderBlocks(first, count)` through `[JSExport]`, with the index kept in memory
or handed back to JS for IndexedDB.

## Rendering (Html/, DocxHtml/)

`DocxStyleCascade` loads `styles.xml` whole (it is small) and resolves docDefaults → basedOn
chain → direct pPr/rPr into the format-neutral `WordParagraphFormatting` / `WordRunFormatting`.
Runs emit only the CSS that differs from the paragraph's base run, so most runs have no inline
style. Headings become `<h1>`–`<h6>` from `w:outlineLvl` or built-in "heading N" names.

**List numbering is counted at index time.** A list paragraph only says "list 7, level 0"; its
number is the count of earlier items, which a windowed renderer never sees. So the scanner runs
`DocxListCounter` (one counter vector per abstractNum, deeper levels reset, `w:start` and
per-num `startOverride` honoured; style-inherited numPr via `DocxStyleCascade.NumberingOf`) and
stores the finished label per block (`"3."`, `"2.a"`, `"ט״ו."`). Formats: decimal, letters,
roman, hebrew1 (gematria with geresh/gershayim), hebrew2, bullets, `%1.%2` patterns. Not done:
`w:lvlRestart`, legal numbering. `.doc` lists: see "Binary .doc" below (same counter idea, LSTF/LFO tables).

Deliberately not done yet:

- **Images** — drawn since 2026-11-03 (see "Pictures, tables, list gaps" below); .doc pictures and floating RTF \shp are still `[image]`.
- **Fields** — done 2026-11-04: see "Word fields (all formats)" below.
- **Word 6/95, encrypted, MHTML-as-.doc** — rejected with a clear message (RTF, also when saved as .doc, opens as RTF).

## Coverage plan — "full flow-preview coverage" (estimate 2–3 weeks, not months)

Parsing is the cheap part; each item below is mostly the semantics → CSS table plus a test file.
Ordered by value for text-heavy Hebrew documents:

1. Footnotes/endnotes (deferred by user) — bodies from footnotes.xml, shown like commentary.
2. ~~Themes — `theme1.xml` fonts and colors~~ — done (.docx).
3. Tables, properly — `tblGrid`/`tcW` widths, borders (`tblBorders`/`tcBorders`), row heights, `tblInd`, nested tables, `.doc` rows reassembled from cell marks into `<table>`.
4. Images — `drawing`/`pict` → rels → blob endpoint (inline only; floating as inline).
5. Text boxes / shapes — `mc:AlternateContent` → `w:txbxContent`. A wps text box (DrawingML wp:anchor / wp:inline with
   wps:txbx) is a real box where it is anchored, placed by the same WordPictureHtml.AppendPlacement as pictures (wrap,
   side, dist gaps, offset): wp:extent width, height fixed or a minimum with a:spAutoFit (LibreOffice
   oox/source/drawingml/textbodypropertiescontext.cxx:255-259), bodyPr insets (:75-80; unset = ECMA-376 Part 1 21.1.2.1.1
   defaults 0.1in / 0.05in), vert (:132), a:xfrm rot, srgbClr solid fill / noFill, a:ln width / colour / prstDash. Theme
   colours: see "Drawn shapes" below. Linked boxes: the first box holds the chain's whole txbxContent, so it shows all the
   text (no split). rtf / .doc boxes: still the old inline frame; .docx shapes, groups and VML shapes: drawn as SVG ("Drawn shapes" below) (.odt / .fodt boxes and shapes: see
   "OpenDocument text boxes and shapes" below).
   Wrap on both sides of an object is not possible in CSS: CSS Shapes 1 section 2 ("does not allow content to wrap
   on both sides of a float"); see "Two-sided and contour wrap (all formats)" below.
6. Paragraph borders/shading, tab stops (`w:tabs` positions), character spacing/scale, caps/small caps, underline styles, `w:lvlRestart`, legal numbering.
7. ~~Fields~~ — done for .docx / .doc / .rtf (see "Word fields (all formats)").
8. Tracked changes toggle (show/hide `w:ins`/`w:del`), comments ignored.
9. Math (`m:oMath`) → MathML (ECMA-376 §22.1 maps almost 1:1).
10. `.doc`: ~~fonts (Sttbfffn), lists (PlcfLst/LFO → the same label counter), footnotes (PlcffndRef/Txt), tables~~ done; the remaining common sprms.

Out of scope on purpose: page layout, headers/footers, columns, exact Word metrics.

## Footnotes and endnotes (Word draft-view model)

Word's continuous/draft view keeps notes out of the text and shows them in a notes pane. Same
here: a reference renders as `<sup class="note-reference" data-note-kind data-note-number>`, and
every block is delivered with the bodies of the notes it references (`WordDocumentBlockHtml.Notes`),
so the host lists the notes of the visible blocks without a second round trip.

- **Numbering is counted at index time** (same reason as list numbers): .docx — the scanner counts
  `w:footnoteReference` / `w:endnoteReference` per kind in body order (`customMarkFollows` → no
  number, the following run is the mark), formats from settings.xml `footnotePr`/`endnotePr`
  (defaults decimal / lowerRoman). .doc — PlcffndRef/PlcfendRef order, FRD.nAuto ≠ 0 → numbered.
- **Bodies are streamed like blocks.** .docx: the index pass also scans footnotes.xml / endnotes.xml
  (`DocxNotesPartScanner`: note id → byte range, separators dropped) with deflate checkpoints
  (`DocxNotesPartIndex`, persisted in the sidecar, format v4); a note is read on demand from its
  range (`DocxNoteXmlReader`) and its rendered HTML kept in a cache bounded to ~2 M characters.
  The policy weighs body + notes, so a small body with huge notes still gets a sidecar.
  Proven by `--notes-streaming 20000`: 113 MB footnotes part → open +5 MB heap, reopen 53 ms,
  141 notes read across random windows, heap +2 MB. .doc: the CP ranges from
  PlcffndTxt / PlcfendTxt in the footnote / endnote stories (story starts from ccpText, ccpFtn,
  ccpHdd, ccpAtn), rendered paragraph by paragraph; the leading 0x02 shows the label.
- `NoteReferences` (kind, number, block) for "next footnote" navigation; `NoteCounts`.
- Demo: notes pane under the text (All footnotes / All endnotes), click highlights, hover tooltip.
  Verified live: 239 / 234 footnotes in the samples; synthetic file checks endnotes (i, ii, iii).

## Formatting coverage — audited, not guessed

`FastDocView.Tests -- --audit <files>` streams every text-bearing part (.docx) or
every PAPX/CHPX/style UPX (.doc) and lists constructs the renderers don't handle. On the
דור המלקטים samples both read **0 unhandled**. Deliberately ignored (listed in the audit as such):
pagination (widow/keep/page-break-before), proofing and language ids, revision ids (rsid), kerning,
script hints, east-asian fonts, frame/table positioning. Added for it: tab stops (`tab-size`),
contextual spacing, paragraph/character shading, small caps / caps / character spacing,
`webHidden`, section breaks, table widths / borders / cell margins / vertical alignment,
.doc fonts (SttbfFfn), .doc table rows as real rows (widths, RTL, cell shading, borders),
huge PAPX in the Data stream, `sprmCSymbol`. Found and fixed: `sprmTDefTable`'s cb counts the
operand plus one — reading it as the length desynced every later sprm of table rows.

## Binary .doc (DocBinary/)

The old format already has random access: the **piece table** (Clx → PlcPcd) maps character
positions (CP) to file offsets, each piece either 8-bit cp1252 or UTF-16. So:

- `CompoundFileReader` — [MS-CFB] header, FAT via DIFAT, directory, mini FAT; streams are
  exposed as `Stream`s mapped over their sector chains (nothing copied).
- `DocBinaryFileInformationBlock` — FIB offsets (Word 97+ layout: fcClx at 0x1A2, etc.).
- `DocBinaryParagraphIndex` — one streamed scan for paragraph marks (0x0D / 0x07 / 0x0C) →
  `int[]` of paragraph start CPs; cached as a sidecar like the docx index.
- `DocBinaryFormattedDiskPages` — PAPX/CHPX by file offset through PlcfBtePapx / PlcfBteChpx
  and the 512-byte FKPs (cached per page).
- `DocBinaryStyleSheet` — STSH: names, base chain, sti (1–9 = Heading 1–9), UPX grpprls.
- `DocBinaryPropertyModifiers` — the sprm subset that matters for HTML (jc, bidi, indents,
  spacing, bold/italic/underline/strike, size, color, highlight, super/sub, hidden, fSpec).
- `DocBinaryListDefinitions` + `DocBinaryListCounter` — **list labels, counted at index time**
  like the .docx side. PlfLst (FibRgFcLcb97 index 73: cLst + 28-byte LSTFs, then the LVL array
  that `lcbPlfLst` does not count — one LVL per simple list, nine otherwise) and PlfLfo (index
  74: lfoMac + 16-byte LFOs + LFODatas with LFOLVL overrides; extra 0xFFFFFFFF dwords before
  the first LFOLVL are skipped as LibreOffice does). A paragraph's `sprmPIlfo` (0x460B) is a
  1-based LFO index (0xF801 = none, 0xF802.. = negated index), `sprmPIlvl` (0x260A) the level
  (0x0C = skipped). Label = the LVL's xst with the placeholders at `rgbxchNums` replaced by the
  counters (a placeholder's char value is the level it shows), `nfc` (MSONFC, [MS-OSHARED]
  §2.2.1.3) mapped onto `DocxNumberingDefinitions.FormatNumber`: 0 decimal, 1/2 roman, 3/4
  letters, 0x16 decimalZero, 0x17 bullet, 0x2D hebrew1, 0x2F hebrew2, 0xFF none, others decimal.
  Restarts: a deeper level resets when a more significant one than its restart limit appears
  (`fNoRestart` / `ilvlRestartLim`); LFOLVL `fStartAt` applied the first time that LFO is met;
  `fLegal` shows inherited numbers as decimal. Stored per block as `BlockListLabelIds` into an
  interned `ListLabels` table (index format v4) and rendered as `<span class="list-marker">`.
  Measured on the 40 MB sample: 407 LSTF / 447 LFO, 83,200 of 106,442 blocks labelled (the
  .docx twin: 99,965 of 123,007), 536 distinct labels, cold open unchanged (1.6–1.9 s either way).
  Not done: the LVL's own grpprlPapx/grpprlChpx (indent of the number, its font), `ixchFollow`
  (a space always follows), list labels on the first paragraph of a table row, a list style
  linked through `rgistdPara`.
- **Fast-saved files**: each Pcd's `Prm` ([MS-DOC] §2.9.214) is expanded once when the piece
  table is read — a Prm1's `igrpprl` picks one of the Clx's Prc grpprls (kept in
  `DocBinaryPieceTable.PrcGrpprls`), a Prm0's `isprm` maps through the spec's §2.9.215 table
  (names resolved against §2.6.1/§2.6.2) to a 3-byte grpprl. `Piece.PrmGrpprl` is applied after
  the PAPX of each paragraph (`ResolveParagraph`, `ReadTableInfo`) and after the CHPX of each run
  inside that piece (§2.4.6.1 / §2.4.6.2). The samples have no Prm (0 of 1,962 pieces); the
  corpus pass is the only coverage so far.

### .doc reads: reopen and search without re-reading (2026-11-03)

A measured audit found needless re-reads; fixed (`--doc-reads <file>` reports them: an unbuffered counting stream,
so the numbers are the bytes and read calls the reader asks the file for; `--doc-search-golden <files>` compares the
one-pass search with the old path on any files).

- **Section first blocks in the index** (sidecar format v7, older sidecars are rebuilt): the index pass records the
  block each PlcfSed section starts in (`SectionFirstBlocks`, Elias-Fano), as .docx does. A reopen used to rescan a
  span per section (`BlockAt` for 278 sections through an 8-span cache).
- **Search is one pass** (`DocBinarySearchCursor`): the piece characters are read once in 64K chunks and serve both
  the block scanner (boundaries) and the matcher (each block's characters handed over when its end is known); the
  characters read past a batch stay buffered, so the next batch of the same search continues without a re-read.
  The scanner skips heading text (`CollectsHeadings = false`). The old path (compact-table block ranges + a second
  `ReadText` per block) stays as `SearchByBlockRanges`, the golden reference: `--search` checks both give the same
  batches (hit blocks, counts, next block, end) - Hebrew plain and pointed, English, mixed, match case on/off, notes
  on/off, batch sizes 500 / 37 / 1 and random starts - on the sample and 10 test-data files. Note bodies are still read
  per note (`FeedSearchText`). Theoretical difference: a surrogate pair split by a 64K chunk boundary decodes as two
  U+FFFD (ReadText decoded per piece); not seen in any file.
- **Compound file**: `SectorChainStream.Read` reads consecutive sectors that are also adjacent in the file with one call.
- **List tables**: LSTF / LVL / LFO records are read through a 16 KB `BufferedStream`, and a LVL in three reads
  (LVLF; both grpprls + cch; the characters) instead of one per field.
- **FKP page cache**: LRU (evicts the least recently used page) instead of clearing all 512 pages when full.

Measured on dor-hamelaktim-shabbat.doc (40 MB, 106,442 blocks), net10 Release, idle priority on a busy machine
(CPU time is the steadier number):

| | before | after |
|---|---|---|
| reopen from the sidecar: bytes read | 29,949,545 | 1,186,087 |
| reopen: read calls | 126,633 | 2,830 |
| reopen: wall / CPU ms | 1,860-3,300 / 1,750-2,000 | 70-140 / 30-140 |
| full search (absent word, 213 batches of 500): bytes read | 58,961,006 | 31,128,941 |
| full search: read calls | 221,716 | 7,960 |
| full search, warm: CPU ms | 2,940-3,125 | 1,200-1,330 |

The first full search after a reopen still costs more CPU (~2.6 s) than a warm one: PAPX lookups for every paragraph
warm their caches. Golden dumps (`--golden`) of the sample and 11 test-data .doc files are byte-identical before/after.

## Queued after the current work (user, 2026-10-29)

1. **Right-to-left errors:** some items render LTR in the .doc sample in the demo. The user thinks it
   is broader: run visual comparisons over many files and fix every class of difference at its cause.
   **Reference = LibreOffice headless PDF** (user, 2026-10-29; first Word's PDF, then changed to
   LibreOffice: `soffice --headless --convert-to pdf`) - the PDF's pages are the reference images, not
   screenshots. Our side: the rendered page in
   the demo. Images are allowed for this comparison; debugging a specific bug stays live/DOM.

## Page text width (user, 2026-10-29 - confirmed design, queued)

The FastDocView page's text column uses the book view's max-width mechanism (centred column; on a
narrow window it fills the width and the margins disappear) but the width comes from the document:
page width minus left/right margins as Word / LibreOffice define them (.docx sectPr pgSz/pgMar, .doc
SEP, .odt page layout, .rtf paperw/margl/margr). No page setup (.txt etc.): Word's default for a new
document (verify in a source). Per section, like Word; computed at index time. Zoom scales the width
so lines break at the same words.

Library side (branch work/page-width): `IWordWindowedDocument.Sections` - `WordDocumentSection(FirstBlockIndex,
TextWidthTwips)` in block order, the first at block 0 (`WordDocumentPageSetup`). .docx: w:sectPr (pPr = ends the
section at that paragraph; the body's own = the last section), pgSz w:w - pgMar left/right - gutter; index format 8.
.doc: PlcfSed + Sepx (sprmSXaPage / SDxaLeft / SDxaRight / SDzaGutter), read at open (one record per section);
defaults 12240 / 1800 / 1800 (POI SEPAbstractType). .rtf: \paperw \margl \margr \gutter and the \sectd-reset
per-section forms, RTF 1.5 defaults 12240 / 1800 / 1800; index format 2. .odt / .fodt: master pages from styles.xml
(or the .fodt head) - page-layout-properties page-width - margins (- loext:margin-gutter); a paragraph / table style
with style:master-page-name starts a section; text before uses "Standard"; index format 2. .txt / .mht: Letter with
1 in margins = 9360 (LibreOffice Word-import defaults, PropertyMap.cxx). Not covered: w:gutterAtTop (settings.xml),
mirrored margins, columns (the width is the whole text area, not one column).

## App integration plan — the FastDocView page (approved 2026-10-28)

**Book view is NEVER FastDocView.** FastDocView is its own page: feature folder
`vue-frontend/src/features/fast-doc-view/`, route `/fast-doc-view`, own stores and names. It only
borrows the book view's design and the shared components. (The old idea of slotting documents into
the lines view as synthetic books is dropped.)

User decisions (2026-10-28):
1. Word-family files (.doc .docx .docm .dot* .odt .ott .fodt .rtf) open in FastDocView by default;
   "open as PDF" (the existing Word → PDF conversion, print layout) stays as an option.
2. .txt moves to FastDocView too (today's `/txt-view` holds the whole file in memory).
3. Rename this repo to FastDocView before the app references it.

Backend (KitveiHakodesh: Core owns everything, the service is transport only):
- Core (net10, AOT) ProjectReferences FastDocView (+ SeekableDeflate, TextEncodingDetection).
  `Core/FastDocView/`: open-document registry with idle close; index cache under app data.
- Service ops, one Dispatcher case each: open (summary), blocks(first, count), headings, weights,
  close, and search (new library feature: windowed search over blocks, never the whole text).
  Access through `LocalFileGrants`. File changed on disk → a typed error → the page reopens.
  The hosted relay (ServiceRpcRelay) needs no change.

Frontend:
- `FastDocViewPage.vue`, a document composable (bounded block cache near the window),
  `webview-host/fastDocViewApi.ts` (modelled on `seforimApi.getLinesPaged`); route in `tabStore`
  `TabRoute`, `AppPageView` (keyed per tab), recents, restore by `localFilePath`; he + en texts.
- Routing in `localFileStore`: Word family and .txt → `/fast-doc-view`; "open as PDF" keeps the
  conversion path.

Shared with the book view — reuse as-is: WindowedScroller (+ bar: headings → `landmarkTiers`,
`EstimatedTextLength` → `rowWeights`), SplitPane (text above, notes pane below), TreeView, useZoom,
useVirtualizedTextSelection, useLineCopy, ContextMenu, FloatingSearchBar, root font/theme variables.
Move to shared first (small): text-column CSS (from BookViewLinesContent), BookViewSidePanel,
toolbar overflow/drag helpers, select-all / Ctrl+F keys, zoom key (per tab/file, not bookId),
a per-file saved-position field in tabStatePersistence.
Generalise (medium): TOC tree over headings, scrollbar heading tooltip, title-bar breadcrumb branch,
search bar UI + match navigation, footnote hover (useNoteTooltip), position capture/restore.
Own to FastDocView: a simple toolbar in the same style, the notes pane, document search.
Not shared (library-specific): commentary, annotations, bookmarks, word links, versions, related books.

Order: rename → .odt on-demand index → shared refactors (book view live-checked after) → Core +
service ops → page → routing → live Playwright checks.

## Specifications used

- ECMA-376 Part 1 (5th ed., 2016), WordprocessingML §17 — https://ecma-international.org/publications-and-standards/standards/ecma-376/
- [MS-DOCX] Word extensions to .docx — https://learn.microsoft.com/openspecs/office_standards/ms-docx/
- [MS-OE376] how Office implements ECMA-376 — https://learn.microsoft.com/openspecs/office_standards/ms-oe376/
- [MS-DOC] Word binary format (rev. 12.5, 2026) — https://learn.microsoft.com/openspecs/office_file_formats/ms-doc/
- [MS-CFB] compound file — https://learn.microsoft.com/openspecs/windows_protocols/ms-cfb/
- RFC 1951 deflate; zlib `examples/zran.c` — https://github.com/madler/zlib/blob/develop/examples/zran.c
- PKWARE APPNOTE (zip, Zip64) — https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT

## Roadmap from the LibreOffice research (2026-10-28)

Two read-only research passes over LibreOffice master (writerfilter DomainMapper, ww8par6 sprm
dispatcher, sprmids.hxx, xmloff txtprmap / XMLTextListBlockContext / txtparai, sw xmltbli).
Everything below fits the paradigm: index-time counters/state, render-time per block, no whole-
document model (LibreOffice builds a full SwDoc; we keep only the semantics).

**.docx / .doc gaps, ranked (impact for Hebrew books × cheapness):**
1. ~~.doc sprmCHpsPos is 0x4845~~ — fixed (we had 0x4840 and listed 0x4845 as kerning).
2. ~~Theme fonts (`csTheme`/`minorBidi` → theme1.xml)~~ — done 2026-10-28 (see "Theme, underline, emphasis" below).
3. .doc list labels (PlcfLst / PlfLfo / LVL) — biggest .doc gap; reuse the index-time counter.
2. Theme fonts (`csTheme`/`minorBidi` → theme1.xml) — Hebrew text can get the wrong font today. Index-time theme parse.
3. ~~.doc list labels (PlcfLst / PlfLfo / LVL)~~ — done (see "Binary .doc"); number-run formatting and `ixchFollow` remain.
4. Numbering `suff`, `lvlJc`, `lvlRestart`, `isLgl`, `numStyleLink`, picture bullets — index-time.
5. ~~Theme colors (`themeColor` / tint / shade, `clrSchemeMapping`)~~ — done 2026-10-28.
6. Hyperlinks: docx rels + HYPERLINK fields, .doc PlcfFld; open-field state per block at index time.
7. ~~.doc fast-saved files: per-piece Prm / Prc sprms~~ — done (`Piece.PrmGrpprl`); no sample with a Prm yet.
8. Inline images (docx rels, .doc PICF) through a blob endpoint — render-time.
9. .doc text boxes (txbx story) after the anchoring paragraph — index-time map.
10. `w:sym` / CSymbol glyph with its font (Wingdings bullets).
11. Table-style conditional formatting (`tblStylePr` firstRow / banding) — render-time per row; row count per table at index time.
12. Emphasis marks (`em` → `text-emphasis`) — .docx done 2026-10-28; .doc sprmCKcd still open.
13. Underline styles + color (`u@val`, `u@color` + theme) — .docx done 2026-10-28; .doc sprmCKul patterns / sprmCCvUl still open.
14. Drop caps as a floated first letter (`framePr@dropCap`, sprmPDcs 0x442C).
15. .doc field-code hiding check (CFFldVanish).

**Open formats (design: OdtBodyBlockScanner etc., same IWordWindowedDocument contract):**
1. `.docm/.dotx/.dotm` — opener only (S).
2. ~~**ODT/OTT**~~ — done 2026-10-28, see "OpenDocument text (Odt/)" below. content.xml checkpointed like document.xml; automatic styles + font faces come
   before office:body and are captured during the index pass into the sidecar; lists flattened to
   blocks with labels per XMLTextListBlockContext rules (xml:id, continue-list, continue-numbering,
   start-value); notes are inline `text:note` — their body byte ranges recorded at index time and
   rendered separately; property map from txtprmap (per-script font/size variants, text-position,
   writing-mode for RTL). Split huge tables into row blocks. (M–L)
3. ~~**FODT**~~ — done with ODT. Same scanner over the plain file: offsets only, no checkpoints — the fastest format. (S after ODT)
4. **Flat OPC** — docx parts as byte ranges inside one XML; feeds the docx pipeline via a part-source abstraction. (S–M)
5. **Word 2003 XML** — close to docx element names; inline footnotes. (M)
6. ~~**RTF**~~ — done 2026-10-28, see "RTF (Rtf/)" below. byte-offset blocks at `\par` / `\row`, with a formatting-state snapshot per block (state carries across paragraphs). (L)
7. HTML/MHTML-as-.doc, `.txt` — low priority. (S–M)
Unverified against source (from memory, check before implementing): ODF note import, ODF package
encryption rules, RTF / Word 2003 importer paths, .doc list-table details.

**Rule: always double-check against the source, never rely on memory alone.** Every format fact
above is verified in the spec ([MS-DOC], [MS-CFB], ECMA-376, ODF 1.3) or the real upstream code
(LibreOffice sprmids.hxx / ww8par6.cxx / DomainMapper.cxx / xmloff) before it is implemented.

## Theme, underline, emphasis (.docx, 2026-10-28)

`DocxTheme` (DocxHtml) loads `word/theme/theme1.xml` and `word/settings.xml` once per document in
`DocxDocumentParts` (both are a few KB) and the style cascade resolves through it; the renderer's
static `ApplyRunProperties` / `ApplyParagraphProperties` keep their signatures and gained theme-aware
overloads. Every rule below was read from the source named, not from memory:

- **Theme fonts** — `rFonts` `asciiTheme` / `hAnsiTheme` / `cstheme` win over `ascii` / `hAnsi` / `cs`
  (ECMA-376 §17.3.2.26: "If the asciiTheme attribute is also specified, then this attribute shall be
  ignored"; LibreOffice DomainMapper.cxx `LN_CT_Fonts_asciiTheme` overwrites even with an empty theme
  face, "apparently what Word 2013 does"). `minorBidi` / `majorBidi` first take the supplemental
  `a:font script="Hebr"` face selected by `settings.xml` `w:themeFontLang/@bidi` (he-IL → Hebr), then
  `a:cs` — the order of LibreOffice `ThemeHandler::getFontNameForTheme`; Word's Office theme has an
  empty `a:cs`, so without this step Hebrew would fall back to the Latin face. Our one deliberate
  deviation: a theme slot that resolves to nothing falls back to the explicit attribute instead of
  overwriting with "" (keeps the docDefaults' David rather than losing the font). `eastAsiaTheme` is
  parsed but the model has no East-Asian font slot.
- **Theme colors** — `w:color`, `w:shd` (`themeFill*`), `w:u`, `w:bdr`, `w:pBdr` edges:
  `themeColor` → `clrSchemeMapping` (background1 → bg1 → light1, text2 → t2 → dark2; spec defaults
  when the mapping is absent, §17.15.1.20) → `a:clrScheme` entry (`srgbClr/@val` or
  `sysClr/@lastClr`). The theme supersedes the explicit value when it resolves (§17.3.2.6 "If the
  themeColor attribute is specified, then the val attribute is ignored"; the same sentence on shd
  §17.3.2.32 and bdr §17.3.2.4); without a theme part the explicit `val` — Word's precomputed
  fallback, which LibreOffice uses directly — is kept.
- **Tint / shade formula** — per RGB component, tint `C' = 255 − (255 − C)·T/255`, shade
  `C' = C·S/255` (T, S the hex byte). The spec's prose formula is an image in every copy we could
  reach, so it was reconstructed from the spec's worked examples (accent1 C0504D tint 99 → D99694,
  accent6 F79646 shade 80 → 7B4B23) which it reproduces to the unit. LibreOffice scales HSL
  luminance by the same factor (`Color::ApplyTintOrShade`, tools/color.cxx); Word's own palette
  rounds slightly differently (accent1 4472C4 tint 99: ours 8FAADC, Word writes 8EAADB).
- **Underline** — `u/@val` (ST_Underline §17.18.99) → `text-decoration-style` (double / dotted /
  dashed / wavy; single, words, thick solid), thick and `*Heavy` → `text-decoration-thickness:0.15em`,
  `u/@color` (+ theme) → `text-decoration-color`; `none` keeps the underline off. `words` (skip
  spaces) has no CSS equivalent and renders as a plain underline — same as LibreOffice's WORD_MODE.
- **Emphasis marks** — `w:em` (ST_Em §17.18.24): dot / underDot → `text-emphasis: filled dot`,
  comma → `filled sesame`, circle → `open circle`; underDot adds `text-emphasis-position: under right`.
  LibreOffice maps the same four (DomainMapper::getEmphasisValue).
- Test: `DocxThemeAndUnderlineTest` over `SyntheticDocxGenerator.Write(..., withThemeSamples: true)`
  (paragraphs 45–49); the audit's handled run properties now include `em`.


## OpenDocument text (Odt/, 2026-10-28)

.odt / .ott (zip) and .fodt (flat XML) behind the same `IWordWindowedDocument` contract.

- **Detection** by content, never extension: a zip whose stored `mimetype` entry is
  `application/vnd.oasis.opendocument.text` (or `-template`) → `WordFileFormat.OpenDocument`;
  XML whose root is `office:document` with that mimetype → `FlatOpenDocument`. A manifest
  `encryption-data` on content.xml → `NotSupportedException("password-protected OpenDocument is not supported")`.
- **Index pass** (`OdtBodyBlockScanner`): byte-level like the docx scanner (vectorised text skip,
  quote-to-quote tag jumps, names compared as bytes bucketed by length, prefixes mapped through the
  root's xmlns declarations). Before the body it captures the root tag and the font-face / styles /
  automatic-styles sections and feeds them to the cascade. Blocks: `text:p`, `text:h`, whole
  `table:table`; lists flattened to one block per list paragraph — the block's byte range is the
  paragraph alone. `text:section`, `text:numbered-paragraph`, index elements (only their
  `text:index-body`) are transparent; `text:soft-page-break`, `text:tracked-changes`, `*-source`
  templates, declarations are skipped.
- **List labels** counted at index time (`OdtListNumbering`), rules read in LibreOffice
  (XMLTextListBlockContext / XMLTextListItemContext / txtlists / txtimp SetStyleAndAttrs):
  root `xml:id` = list id; `continue-list` follows the continue chain to the master;
  `continue-numbering="true"` without continue-list continues the last processed list of the same
  style; nested lists inherit id/style/restart, the 2nd+ sub-list of an item restarts; only the first
  paragraph of a list-item is numbered (list-header and paragraphs after a nested list are not);
  item `start-value`. Level formats from `text:list-style` (xmlnumi.cxx defaults: format "1",
  start 1, display-levels 1). NOT from source: deeper counters are cleared when a level is numbered
  (Writer's usual behaviour); `text:style-override`, the OOo-1.x default-list-id path and the
  MS-Office "continue by style" path are not implemented.
- **Notes** are inline (`text:note`): number per class in body order (+ notes-configuration
  start-value — the offset semantics are NOT source-verified), label = the citation text the producer
  wrote (else `text:label`, else the configured format). The note-body byte range is stored; the body
  is rendered from the parsed block and delivered as `WordDocumentNoteHtml`; note text is excluded
  from the block's text length and heading text. Counting per page/chapter is shown document-wide
  (flow view).
- **Styles** (`OdtStyleCascade`): default-style → parent chain → style; property names from
  txtprmap.cxx (text, paragraph) and sw xmlitemm.cxx (table, column, cell). Complex-script variants
  (`*-complex`) win in RTL paragraphs (no per-character script detection). The index stores the
  sections reduced to what rendering reads (no list styles, no unmapped attributes): the LibreOffice
  conversion of the sample has 14.5 MB of `office:styles`, 3058 list styles — a reopen would parse
  them every time otherwise.
- **Modes**: package in-memory (content.xml kept, small files), package sidecar (deflate
  checkpoints over content.xml), flat (the file is the view; the sidecar only saves the pass).
- **Measured** (dor-hamelaktim-shabbat.docx converted by LibreOffice 25.x headless: content.xml 55.8 MB,
  121,650 blocks, 183 headings, 234 footnotes): cold index 1.76 s (the .docx with an 87 MB body:
  1.9 s), reopen 151 ms (docx 111 ms), 30-block windows 4–19 ms. Synthetic 100k blocks: .odt cold
  ~290 ms, reopen 47 ms; .fodt (17 MB) reopen 6–8 ms; random 40-block windows ≤ 28 ms.
- **Not done**: huge tables are one block (no row split — no table near 2 MB seen yet); lists inside
  table cells and notes get a plain bullet, not their counted label; heading outline numbering
  (`text:outline-style`) is not shown except where the producer wrote `text:number`; tab stops,
  images (placeholder), list-level indents (a fixed 18 pt per level) are approximations.
## On-demand index for the first open (.docx, 2026-10-28)

The first open of a big .docx no longer waits for the whole index pass. With
`WordIndexPolicy.IndexesOnDemand` (the dev server sets it; the library default stays off so
`BuildSidecarIndexes`, the corpus pass and older hosts keep the one-go build) a sidecar open with no
cached index scans only until the first 100 blocks are known and returns.

- `DocxIncrementalIndexBuilder` owns the body's inflater (`DocxEntryInflationCursor`, the pull form of
  `DocxEntryInflation.Run`: same 256 KB chunks, same checkpoint rule) and the `DocxBodyBlockScanner`.
  It is advanced **synchronously in the caller's thread** ("until block N is known" / to the end) - no
  background thread. A block is known once the next one has started (or `</w:body>` was seen).
- `IDocxBlockTable` is what the block reader and renderer read per block; the finished
  `DocxRandomAccessIndex` and the builder (over its known blocks) both serve it. The block reader's
  `DeflateRandomAccessStream` is opened over the builder's live checkpoint list, so it already seeks
  anywhere in the part scanned so far.
- `DocxWindowedDocument`: `BlockCount` = `EstimatedBlockCount` while incomplete (known + unscanned
  body bytes / average block bytes so far), `KnownBlockCount`, `IsIndexComplete`, `CompleteIndex()`.
  `RenderBlocks(first, count)` first advances the scan to `first + count`; a jump to the end pays the
  rest of the pass once. `EstimatedTextLength` past the known blocks returns the known average;
  headings / note references are those of the known blocks.
- **Notes**: `DocxIncrementalNotesPart` pulls footnotes.xml / endnotes.xml forward only until a
  referenced note's range is found (Word writes notes in reference order, so early blocks find theirs
  near the start; every part is still inflated once in total). Chosen over "scan the notes first"
  because a notes part can be bigger than the body (`--notes-streaming`).
- On completion the finished index replaces the builder (scanner lists freed, note reader re-opened
  over the index) and `Open(path)` saves the sidecar. The index is **byte-identical** to
  `DocxRandomAccessIndexBuilder.Build` (shared `CreateIndex` / `CreateNotesPartIndex`).
- Dev server: blocks responses carry `X-Known-Block-Count`, `X-Estimated-Block-Count`,
  `X-Index-Complete`; the summary has `knownBlockCount` / `isIndexComplete`. The demo uses the
  estimate as the scroller count, corrects it from every blocks response (no polling) and refetches
  headings + weights when the index completes. Weights are applied only while their length matches
  the count. `playwright/incrementalCheck.mjs` checks it live.
- Measured (`--incremental-open`, JIT warm): 1M-block synthetic (350 MB body) open + first 40
  blocks ~25 ms (was ~10 s for the full pass), window at 10 % ~1 s, at 50 % ~1 s more, completion
  from 50 % ~1.6 s; the 87 MB-body sample first window ~80 ms, 10 % ~170 ms, 50 % ~0.8 s, end ~0.4 s.
  Heap at 50 %: 27 MB (synthetic) / 13 MB (sample).

## Plain text and Word web archives (Markup/)

**.txt is read as HTML** (user's decision): embedded tags render, `<h1>`-`<h6>` feed the outline.
One block per line; a line break inside a tag or a comment does not end a block. `MarkupBlockScanner`
scans raw bytes (single-byte code pages and UTF-8: vectorised `IndexOfAny('<', '\n', '\r')` over text
runs, no decoding) or 2-byte units (UTF-16), and stores per block the start tags open at its start
(interned), so a window parses standalone; `MarkupBlockRenderer` reopens them, escapes a '<' that does
not start a tag, drops comments / script / style, strips on* handlers and non-http(s) URL schemes
(entities decoded first), replaces images with `[image]`, closes what is left open. Both sides share
the tag rules (tag ≤ 2048 code units, comment ≤ 64 KB, else text). Lines over 64 KB split at a space,
over 256 KB anywhere outside a tag (UTF-8: never inside a sequence). Implied end tags (li, dt/dd,
td/th, tr, p closed by block starts) follow the HTML rules, scoped; stack depth capped at 32.

**Encoding (.txt) = Notepad++, via the TextEncodingDetection library** (sibling repo
github.com/KleiKodesh/TextEncodingDetection, `ProjectReference` like SeekableDeflate; 2026-10-28).
Verified in Notepad++ source (Buffer.cpp `copyFileContentToScintilla`/`detectCodepage`/
`resolveLoadedEncoding`, Utf8_16.cpp `determineEncoding`/`utf8_7bits_8bits`, EncodingMapper.cpp,
Parameters.h defaults): empty → UTF-8; BOM (FE FF, FF FE, EF BB BF; no UTF-32) → uchardet on the
first 128 KB+4 (autodetect is on by default; > 0.60; TIS-620 dropped; names without a Notepad++
mapping such as KOI8-R / EUC-JP fall through) → UTF-16 LE without BOM (even length, byte0≠0,
byte1=0, IsTextUnicode statistics, emulated and checked against Win32) → UTF-8 structure (NUL or a
sequence cut by the block end = not UTF-8; 7-bit opens as UTF-8) → ANSI = windows-1255. The uchardet
is a port of the copy Notepad++ bundles (not modern uchardet), checked against that C++ compiled
with MSVC on a 2,628-case matrix + uchardet's test folder. Codes it can return: 65001, 1200, 1201,
1251, 1252, 1253, 1255, 28595, 28597, 28598, 10007, 866, 855, 932, 936, 51949, 950 — all decode on
net48 and net10 (`FromCodePage`). The index keeps the code page and the reason (e.g.
`uchardet windows-1255 (0.99)`). Change vs the old reimplementation: a UTF-8 sequence cut by the
128 KB sample end now counts against UTF-8 (as in Notepad++), but uchardet decides UTF-8 first on
real files; on 500 real files the result changed for 0. Double-byte code pages (932/936/949/950/51949)
never split a block between lead and trail byte (trails are >= 0x40, so CR/LF/space/'<' never fall
inside a character). Index format v2 (re-detects old .txt indexes). LibreOffice
(tools/source/stream/stream.cxx `SvStream::DetectEncoding`, called from sw/source/filter/basflt/iodetect.cxx
`SwIoSystem::IsDetectableText`, sw/source/filter/ascii/parasc.cxx) does BOM then ICU ucsdet with no
confidence threshold, fallback the system ANSI code page. `<meta charset>` is not honoured for .txt
(Notepad++ only sniffs it for HTML/XML file types).

**Loose .html / .htm: not supported (decided 2026-11-03).** Real web pages often need scripts, external CSS and live
resources that a static windowed reader does not run, so they would render broken. The reader, its encoding prescan
and its test were removed. Word web archives (MHTML) stay: they are self-contained static documents.

**MHTML** (.mht/.mhtml, and .doc files beginning with "MIME"): `WebArchiveContainer` reads headers only
(RFC 2045 §5.1 parameters, RFC 2046 §5.1.1 boundaries; a part may itself be multipart with its own boundary, e.g.
Internet Explorer's multipart/related > multipart/alternative > text/plain + text/html) and takes the first text/html part, depth
first; its body is
decoded streaming (`MimeTransferDecoder`, RFC 2045 §6.7 quoted-printable leniently, §6.8 base64) and
scanned in Elements mode: a block per p / h1-6 / table / div / list element outside tables. Per block
the index keeps a **resume point** = encoded offset × 4 + skip: quoted-printable restarts at the token
that produced the byte (no decoder state survives a token boundary; soft breaks produce nothing),
base64 at the 4-character group start with 0-2 bytes to skip. Charset from the part's Content-Type
(fallback windows-1255). Word's head `<style>` is kept, scoped under `.word-document`, @-rules dropped.
Measured: synthetic 20k-block QP / base64 1.1 MB: cold ~70 ms, reopen ~1 ms. The corpus has 1 such .doc
(of 5,355 Word files).

Measured (.txt, sidecar forced): largest real txt 4.6 GB / 6.9 M blocks — cold 5.7 s, reopen 101 ms,
60-block window 0.35 ms avg; synthetic 37 MB / 1 M blocks — cold 311 ms, reopen 15 ms, window 0.24 ms.
Corpus sample (largest 200 + random 300 of 11,276 txt): 0 crashes; windows-1255 234, UTF-8 201,
7-bit 44, UTF-8 BOM 18, UTF-16 LE BOM 3.

**Compact sidecar (format 6, 2026-11-03)**: no per-block table. The index keeps one `MarkupSpanAnchor` per
~64 KB of text (`AnchorSpanBytes`): a `MarkupBlockScanner` snapshot (`MarkupScanState`: the open block - index,
start, stack id, text length so far - plus the live element stack as an interned stack + its tags' lengths, the
unpushed / table depths, two flags) taken between two fed pieces where the scanner holds only text-level state
(`IsAtSpanPoint`: no tag, comment, raw text, pending CR or half a code unit); MHTML adds the decoder resume points of
the snapshot position and of the open block's start. Anchor 0 is a fresh scan from the text's start. Besides the
anchors: text-length totals per 64 blocks (scroll weights; exact once a span is scanned), the interned stacks, the
headings (Elias-Fano blocks, levels, texts deflated), the head CSS. A block's start / end / text length / stack /
resume point: `MarkupCompactBlockTable` rescans its span with the same scanner seeded from the anchor (8 spans
cached, LRU). The index pass hands finished blocks off into the groups (no per-block lists even while building).
**.txt is fed in fixed 64 KB pieces** (after the 128 KB detection sample) by both the pass and the rescans: a line
over 64 KB splits at the first space the per-byte path sees, and where that path starts depends on the piece
boundaries, so a rescan must see the same boundaries (the pass now reads full pieces even from a trickling stream).
Golden dumps (every block's HTML, headings, search hits, reopen from the sidecar) byte-identical before/after for 11
files: real Otzaria books (UTF-8 44 MB, the same text as windows-1255 19 MB and UTF-16 39 MB), the edge-case and
Shift_JIS long-line fixtures, synthetic UTF-16 BE, synthetic MHTML 20k and 300k blocks (QP + base64), Firefox MHTML.

Measured 2026-11-03 (Release net10, idle priority, 2 cores; shared box, +-20 %):

| file | sidecar before (v5) | after (v6) | % of file | reopen before / after | first window | random block median / p90 before -> after |
|---|---|---|---|---|---|---|
| real UTF-8 .txt 44.6 MB, 53,677 blocks | 439,366 B | 59,764 B | 0.13 % | 5 / 12 ms | 0.6 / 1.3 ms | 0.3 / 0.4 -> 0.5 / 0.7 ms |
| same text cp1255 19.3 MB | 723,710 B | 49,989 B | 0.26 % | 20 / 34 ms | 7.5 / 17.6 ms | 0.3 / 0.3 -> 0.7 / 1.2 ms |
| same text UTF-16 38.7 MB | 724,504 B | 51,578 B | 0.13 % | 6 / 14 ms | 0.6 / 2.0 ms | 0.3 / 0.4 -> 0.9 / 1.4 ms |
| synthetic MHTML base64 18.7 MB, 300k blocks | 1,513,976 B | 3,197 B | 0.017 % | 8 / 2 ms | 4.7 / 22 ms | 0.8 / 1.5 -> 3.6 / 5.5 ms |
| synthetic MHTML QP 17.3 MB, 300k blocks | 1,513,976 B | 2,875 B | 0.017 % | 13 / 11 ms | 3.1 / 2.9 ms | 1.5 / 1.8 -> 2.8 / 5.2 ms |

**Span rescans run lazily, MHTML anchors every 16 KB (2026-11-03, same format 6).** The v6 table above measured the
first window once per process (JIT included: the base build showed 22 ms on the same first call too). Best-of-15
reopens with TieredCompilation=0, A/B in the same minute against a worktree at b20507f, synthetic MHTML base64 300k:
first window 0.85 ms (v5) / 3.5 ms (v6) / 0.94 ms (now); random block median 0.81 / 4.06 / 1.27 ms (QP: 0.71 / 3.55 /
1.05 ms). Cause: every first touch of a span rescanned the whole 64 KB span (~2.5 ms of decode + tag-heavy scan at
~25 MB/s) even when the window needed its first 40 blocks; the 0.8 ms common to all three is the reopen of the file
handle + render. Fix: `MarkupSpanRescan` keeps the span's scanner (and MHTML decoder) alive and feeds the pass's own
pieces only until the asked-for block is complete (its next block started), resuming there on the next ask; MHTML
anchors every 16 KB (`WebArchiveAnchorSpanBytes`; .txt stays at its 64 KB piece boundaries). Sidecar 3,197 -> 3,359 B
(base64) and 2,875 -> 3,045 B (QP): the anchors compress to ~0.8 B per KB of text. Golden dumps byte-identical for
the 11 files. Older v6 sidecars still read (anchors are stored), with 64 KB spans.

**Search and windows read each byte once (2026-11-03).** `ReadText` used to read (MHTML: decode again from the first
block's resume point, in 64 KB reads) the very bytes the span rescans had just read and decoded to find the blocks'
bounds. Now `MarkupSpanRescan` hands every piece it feeds its scanner to a `WordIndexByteRing` owned by the table
(64 KB, grows to 4 MB; one contiguous range), `ReadText` scans the blocks in order and copies their text from the ring;
it reads only when the bounds were known from an earlier scan and the bytes are no longer held. Full search over a
reopened index (no span scanned yet), synthetic 20k, batches of 50 and 500, bytes read / file size: .txt (7 encodings)
2.00x -> 1.00x; MHTML quoted-printable 3.46x -> 1.18x, base64 3.21x -> 1.09x (what is left is the rescans' 4 KB
encoded pieces running past each 16 KB anchor, read again by the next span). `--text` / `--mhtml` check it
(`SearchReadsOnce`). Golden dumps byte-identical on 9 markup files (edge cases, Shift_JIS long line, search-v1,
synthetic .txt utf8 / cp1255 CR / UTF-16 BE 20k, synthetic MHTML QP / base64 20k, Firefox MHTML).

On the real books **the headings are ~90 % of the sidecar** (14,525 headings = 46.7 KB; 6,521 long ones = 53.7 KB):
the outline's own text, deflated. Anchors are 2-4 KB, the weights 1-2 KB. At 256 KB spacing the anchors were 4x
smaller but an MHTML random block cost 17-30 ms median (the scanner rescans ~25 MB/s of tag-heavy HTML), so 64 KB.

Open: a Word table as one block is not
bounded (one huge table = one big block); EstimatedTextLength counts bytes (UTF-8) or units (UTF-16).
## RTF (Rtf/, 2026-10-28)

.rtf behind the same `IWordWindowedDocument` contract. Detected by the `{\` signature (any extension).
RTF is plain bytes, so like .fodt the file itself is the random-access view: no checkpoints.

- **One streaming pass** (`RtfParser` + `RtfRandomAccessIndexBuilder`): a byte-level tokenizer with a
  group-state stack; no document model. Blocks: a paragraph ends at `\par` / `\sect` outside a table; a
  table row (`\trowd` … `\cell` … `\row`) is one block (`TableRow`, like .doc rows — a whole table would
  be unbounded). A block's range starts right after the previous terminator (or after the last header
  group at the root), and the parser's **group-state stack at that point is stored with it** (plain
  `RtfGroupState` structs as bytes, interned: the 300k-block synthetic has 25 distinct stacks, the
  largest real file 6). A window restores the stack and parses only its own bytes — a block that ends
  inside a group (`{\i … \par}`) and the next one that starts with the `}` both render right.
- **Header tables** read once by the pass into the index (`RtfDocumentTables`): fonts (name, `\fcharset` /
  `\cpg` / name suffix → code page), colors, paragraph styles (name + `\outlinelevel`), list table +
  overrides (only for counting; labels are stored per block), note formats, `\ansicpg`, `\deff`.
  In a window the header destinations are skipped.
- **Text**: 8-bit text and `\'hh` decode through the run's code page — set by `\f` from the font's
  charset (LibreOffice: only `\f`, not `\af`, changes it), `\plain` → the `\deff` font's, `\rtlch` →
  1255 when `\ansicpg1255` (LibreOffice). Bytes are decoded per run so double-byte code pages work;
  runs of `\'hh` are gathered in one go (Hebrew is written that way). `\uN` + `\ucN`: N chars skipped,
  where a byte, a `\'hh`, a control word and `\binN` + data each count one and a brace ends the skip
  (RTF 1.5 spec). Code pages come from `CodePagesEncodingProvider` (net10) / the framework (net48).
- **Fonts / RTL**: `\f \fs \b \i` and `\af \afs \ab \ai` go to the Latin or complex-script slot by
  LibreOffice's run-type rule (after `\rtlch`, or `\ltrch` then `\rtlch` → complex). A run's slot is
  picked by its characters' script (Hebrew/Arabic → complex) — but a complex slot never set since
  `\plain` falls back to the Latin value, because RichEdit writes Hebrew with `\f`/`\fs` only (our
  choice, not from a source). `\ql`/`\qr` are physical sides: Hebrew paragraphs in the 115 real files
  are `\rtlpar\qr` 2,147 times vs `\ql` 3, so `\qr` in an RTL paragraph is the start side (corpus
  evidence; the spec is silent). `\li`/`\ri` physical, `\lin`/`\rin` logical (preferred when given).
- **Lists**: labels counted at index time from `\listtable` / `\listoverridetable` (`\levelnfc` →
  LibreOffice's getNumberFormat map, `\leveltext` first char = length, placeholders at `\levelnumbers`
  positions, `\levelnorestart`, `\lfolevel` start-at once per override); when the paragraph carries
  `\listtext` / `\pntext` (Word's flat label) that text is rendered instead and no label is stored.
- **Notes**: `{\footnote …}` inline; `\ftnalt` → endnote; auto-numbered when the body had `\chftn`
  since the previous note (else a custom mark: number 0, label "*"), per kind, in body order,
  formats from `\ftnn*` / `\aftnn*` (default decimal for both — the spec is silent on the endnote
  default), `\ftnstart` / `\aftnstart`. The reference `<sup class="note-reference">` is written where
  the note group ends; the body is a `WordDocumentNoteHtml` with the block; note text is not counted
  in the block's text length or heading text.
- **Headings**: `\outlinelevel` 0-8 on the paragraph, else the style's (`\outlinelevel` in the style
  sheet or a "heading N" name); `\outlinelevel9` = body text.
- **Skipped** (whole group, fast brace scan honouring `\bin`): every LibreOffice DESTINATION we don't
  render (header/footer, info, bookmarks, fldinst, pict data, themedata, datastore, …) and any `\*`
  destination we don't know. Pictures / objects / equations → `[image]` / `[object]` / `[equation]`.
  `\upr` shows its `\ud` half. Fields show `\fldrslt`. Nested tables: `\itap`, `\nestcell`, and the
  `\nestrow` inside `{\*\nesttableprops}` (its `\trowd`/`\cellx` are not applied to the outer row).
- **Audit** (`--rtf-audit`, and the corpus pass): control words outside skipped destinations that the
  reader does not act on; each ignored one is listed with its reason in `RtfFormattingAudit.IgnoredByDesign`.
  The 115 readable real RTF files on the dev machine (118 listed; 3 OneDrive placeholders): 0
  crashes, 0 unhandled, 0 first windows that look mis-decoded (Latin-1 letters where Hebrew was meant).
- **Measured** (busy machine, JIT warm): largest real file 6.8 MB / 15,110 blocks — index pass
  ~105-150 ms, sidecar cold open 140-175 ms, reopen 17-20 ms, 40-block window in the middle 14 ms.
  Synthetic 23 MB / 300,016 blocks: full index 230-330 ms (parser alone ~135-200 ms), sidecar cold
  open ~410 ms, reopen ~46 ms, random 40-block windows <= 9 ms.
- **Not done**: paragraph borders/shading, tab stops, cell merges (`\clmgf` / `\clvmgf`), row heights
  and table borders, hyperlink targets, style
  formatting not repeated in the paragraph (Word always repeats it), `\pn` (Word 6) numbering without
  `\pntext`, Shift-JIS lead bytes that are `\`/`{`/`}`, the Chicago note format (from memory).

## Format support limits (corpus pass 2026-10-28)

- Not supported, precise NotSupportedException: Word 6.0/95 (.doc wIdent 0xA5DC / 0xA697-0xA699),
  Word for Windows 1.x/2.0, Word 2003 XML, HTML saved as .doc, encrypted .doc/.docx/.odt.
- **Live demo check of every format (2026-10-28):** .odt (121,650 blocks), .fodt, .rtf, .txt (ToratEmet),
  .mhtml and .docx all pass liveCheck.mjs. **Speed gap:** a cold .odt open shows text only after the
  whole index is built (2.6 s on the 121k-block sample) — .odt has no on-demand index like .docx yet.
- **GAP — older Word formats (open, to do):** Word 6.0/95, Word for Windows 1.x/2.0 and Word 2003
  XML (.doc that is XML) are rejected, not read. test-data holds 37 Word 6/95 files (nFib 101 / 104,
  wIdent 0xA5DC; 0xA697–0xA699), 10 Word 1/2 files and 1 Word 2003 XML. None of the user's 3,470 real
  .doc files failed this way (corpus pass 2026-10-28), so it is low priority, but it is a real gap.
  Word 6/95 is the feasible one: same piece-table/FKP idea as Word 97, but a different FIB layout,
  1-byte sprm opcodes and a different style sheet. Sources: LibreOffice ww8 "Ver67" paths
  (ww8par*.cxx, ww8scan.cxx), Apache POI HWPF old-format support (HWPFOldDocument / OldFIB). There is
  no Microsoft spec for it — treat facts as unverified until checked in both. Word 2003 XML would be a
  small streaming reader of its own (WordprocessingML 2003 schema).
- Malformed XML in any part → InvalidDataException naming the part (LibreOffice opens some of these,
  e.g. math-malformed_xml.docx, with a recovering parser — not done here).
- Malformed XML: see "Malformed XML (repair mode)" below — .docx keeps the content before the
  first error; .odt/.fodt content.xml/styles.xml still fail with InvalidDataException naming the part.

## Malformed XML (repair mode, 2026-10-28)

What LibreOffice does (read in source, master + libxml2 2.15.4 from download.lst):
- sax/source/fastparser/fastparser.cxx: libxml2 push parser, options only `XML_PARSE_NOENT |
  XML_PARSE_HUGE` — **no XML_PARSE_RECOVER**. libxml2 parserInternals.c `xmlCtxtVErr`: a fatal
  error with `recovery == 0` sets `disableSAX = 1`, so no event follows it; `xmlParseChunk` returns
  the error and fastparser throws SAXParseException.
- .docx (sw/source/writerfilter/ooxml/OOXMLDocumentImpl.cxx `resolve`): the SAXException is
  rethrown unless `SfxObjectShell::IsContinueImportOnFilterExceptions` (sfx2 objmisc.cxx) — the
  user approved the repair prompt (or repair mode is on). Then the document keeps what was imported
  before the error. Without approval the load fails: ooxmlimport.cxx testMathMalformedXml asserts
  the component is null.
- .odt/.fodt (sw/source/filter/xml/swxml.cxx ReadThroughComponent): a SAXParseException in
  styles.xml / content.xml is ERR_FORMAT_FILE_ROWCOL — the load fails (odfimport testForcepoint108,
  uiwriter testThreadedException assert failure). meta.xml / settings.xml errors are warnings.
- Error classes (libxml2 parser.c 2.15): mismatched end tag XML_ERR_TAG_NAME_MISMATCH, duplicate
  attribute xmlErrAttributeDup, undeclared entity xmlHandleUndeclaredEntity (fastparser's getEntity
  callback returns only predefined/char entities and userData != ctxt, so DTD-declared entities are
  never looked up), illegal character, premature end — all XML_ERR_FATAL: everything stops there.
  (With XML_PARSE_RECOVER, which LibreOffice does not set, an end tag closes the innermost open
  element whatever its name — xmlParseEndTag2 — and duplicate attributes are reported but parsing
  goes on.)

What we do:
- Strict XmlReader / XDocument parsing stays the normal path, unchanged. Only when it throws
  XmlException is that block / note / part re-read by `WordDocumentXml.LoadUntilFirstError`: the
  tree of every node delivered before the first error, open elements closed there. Same reader
  settings (DtdProcessing.Ignore — nothing external is ever fetched or expanded).
- Applies to .docx body blocks, notes, styles, theme, settings, numbering. ODF is unchanged (fails,
  as LibreOffice does).
- The body / notes scanners step over `<!DOCTYPE … [internal subset]>` (DocxDocumentTypeDeclarationSkipper);
  it was taken for an unterminated comment, which swallowed the body.
- Differences from LibreOffice: (1) recovery is per block — LibreOffice drops everything after the
  first error in the part, we drop the rest of the failing block only (knowing the part-wide first
  error needs a whole-part strict parse at index time); (2) text in the same text node before an
  undeclared entity is lost (XmlReader throws for the whole node; libxml2 delivers the characters
  before '&'); (3) we open without a repair prompt.
- Corpus: math-malformed_xml.docx, tdf165348_broken_package.docx, ExternalEntityInText.docx and
  clusterfuzz POIXWPFFuzzer-6733884933668864.docx (theme1.xml) now open. Speed on well-formed input:
  index build / scanner on the 87 MB sample within run-to-run noise (scanner 187–233 MB/s both).

## Word fidelity pass (.docx, branch work/oxpt-ideas, 2026-10-28/29)

Ideas taken from Open-Xml-PowerTools (MIT, read only) and re-checked in ECMA-376, [MS-OI29500], LibreOffice
writerfilter/dmapper and Word 16 (COM / PDF export of synthetic files). Regression checks: `FastDocView.Tests
--fidelity` (DocxFidelityTest*.cs, seconds).

- **Toggle properties** (b, bCs, i, iCs, caps, smallCaps, strike, emboss, imprint, outline, shadow, vanish) are kept per
  style layer (`WordToggleLayers`): last value wins inside a basedOn chain, table / paragraph / character styles XOR,
  docDefaults on stays on, direct formatting absolute (Word 16). dstrike and webHidden are not toggles.
- **Script choice**: a run with w:rtl or w:cs uses cs font / szCs / bCs / iCs for all its text; without them ascii /
  sz / b / i, even for Hebrew letters and in a bidi paragraph (Word 16 PDF). CJK characters use rFonts eastAsia.
- **Underline / strike**: lines are drawn on the runs, never on the paragraph (a child cannot remove a parent's line);
  .docx draws the underline on text pieces: not under the spaces ending the paragraph (unless compat ulTrailSpace),
  "words" skips spaces, a bare `<w:u/>` draws nothing; dstrike = double line-through on its own element.
- **Lists**: lvlOverride/lvl, numStyleLink, lvlRestart, isLgl, suffix (tab / space / nothing), label = mark rPr + lvl
  rPr; hebrew1 = Word's letters without geresh, cycling after 392. pPr/rPr formats the mark only (label, empty
  paragraph height).
- **Shading** w:shd patterns mixed as LibreOffice CellColorHandler; highlight wins over run shading.
- **w:sym** keeps its font (U+F0xx). **Links**: w:hyperlink (r:id / anchor), HYPERLINK fields (also across
  paragraphs: the index stores the fields open at each block start, format 7) and fldSimple; only http(s), mailto
  and #bookmark become `<a>`; bookmarks render as element ids.
- **Tables**: tblStylePr conditional formatting chosen and layered as LibreOffice DomainMapperTableHandler (tblLook,
  bands, corners); Word's default-paragraph-style 11/12 pt and left-alignment rule unless
  overrideTableStyleFontSizeAndJustification; deleted rows / cells hidden.
- **Neighbours**: a body block is rendered with the paragraph before and after it (read in the same pass as the
  window) - border groups (pBdr incl. between) and deleted paragraph marks (text joins the next paragraph, which keeps
  its pPr) work across blocks; inside cells / text boxes the siblings are used.
- **Tabs**: stops merge across layers ("clear" removes); a last tab to an end-aligned stop (or ptab right) becomes a
  flex filler drawn with the leader. Plain stops keep tab-size.
- **Compatibility** (settings.xml w:compat, `DocxCompatibilitySettings`, read like LibreOffice SettingsTable):
  compatibilityMode (highest valid, absent = 12), doNotExpandShiftReturn (default: lines ending in a manual break
  of a justified paragraph are stretched), autospacing (280 / 100 twips + LibreOffice's first / last / same-list
  adjustments), mode > 14 no space before a paragraph starting with a page / column break, stated mode 11-14 trims
  trailing spaces in table cells, ulTrailSpace. Not handled (no source behaviour found in LibreOffice):
  useNormalStyleForList, doNotUseIndentAsNumberingTabStop, allowSpaceOfSameStyleInTable, doNotSuppressIndentation,
  underlineTabInNumList, forgetLastTabAlignment, useWord2002TableStyleRules - the formatting audit lists them.
- Known gaps: links inside notes do not resolve r:ids (notes have their own rels); trailing-space handling is per
  paragraph end, not per wrapped line; CSS has one dash pattern and no double wave.
- Demo rig ports: FASTDOCVIEW_DEV_URL / FASTDOCVIEW_DEMO_PORT / FASTDOCVIEW_DEMO_URL (defaults 5188 / 5189). On the
  real .docx the liveCheck reload-anchor checks and notesCheck "document has notes" fail on master as well (pre-existing).

## Rendering comparison against LibreOffice's PDF (branch work/word-compare, 2026-10-29)

Reference = LibreOffice headless (`soffice --convert-to pdf`, private `-env:UserInstallation` profile, below-normal
priority) - user's decision; Word COM was tried first and paginating the 40 MB .doc did not finish in 15 minutes.
Scripts (scratchpad, not in the repo): `ref.py` (convert + sample pages + PyMuPDF line runs: bbox, font, size, flags;
pieces of one visual line merged), `--word-compare-page` (harness: blocks holding the sampled lines, written as a
static page at the reference text width), `measure.mjs` (Playwright: per paragraph computed direction, alignment,
dominant font / size / weight / style by letters, line boxes), `compare.py` (match lines to paragraphs by letters;
classes: placement, indent per side relative to the page's least indented paragraph, font family / size / bold /
italic; side-by-side PNGs), `report.py` (HTML report, before / after).

Fixed causes (each with a harness check):
- sprmPJc80 and sprmPDxaLeft80 / sprmPDxaRight80 are physical left / right ([MS-DOC] 2.6.2), sprmPJc and
  sprmPDxaLeft / sprmPDxaRight logical; all were read as start / end, so RTL paragraphs had their indent on the
  wrong side (`DocBinaryBidiTest`).
- sprmCFBoldBi / sprmCFItalicBi were merged into bold / italic; a right-to-left run draws with the Bi set only
  (`WordRunFormatting.EffectiveIsBold`).
- A right-to-left run in a left-to-right paragraph (w:rtl, sprmCFBiDi, \rtlch) is laid out right to left with its
  punctuation (ECMA-376 1st ed. rtl; RTF 1.5 \rtlch): rendered as `unicode-bidi:embed;direction:rtl` - an embedding,
  not a `dir` span (isolate), so consecutive RTL runs keep one right-to-left order.

Known limits of the reference: LibreOffice stops drawing all formatting (underline, bold, list labels) in the .doc
sample from about page 1900; `compare.py` reports differences after such a cliff apart. LibreOffice picks the
complex-script font / size / weight by the characters' script; Word (checked for .docx, LEARNINGS 2026-10-28/29)
by the run's rtl flag - font-family / size differences on Hebrew runs without sprmCFBiDi / w:rtl are that
disagreement until checked in Word.

## List label layout (branch work/indent-parity, 2026-10-29)

Measured per paragraph against LibreOffice's PDF (scratchpad tools lref.py / lmeasure.mjs / lcompare.py: label position,
label-to-text gap, first-line and later-line start, end indent, line height, paragraph pitch, line count; per page the
median start-side shift is removed because the column estimate of a PDF is only good to a few points).

- Shared `WordListLabelHtml`: the label starts at start indent + first-line indent; after a tab suffix the text starts at
  the next stop after the label's position - a custom stop, the start indent when the paragraph hangs, else the next
  default stop (Dop dxaTab / defaultTabStop, 720). The label is an inline-block with that gap as min-width. lvlJc / LVL jc
  end or centre: a zero-width nowrap flex box (flex-end / center) so the label ends at / centres on its position. A label
  wider than the gap pushes the text on (Word would use the following stop; the label's width is not known server-side).
- .doc: the LVL's grpprlPapx (indents, list tab) applies between the style and the paragraph's own sprms (LibreOffice
  ww8par3.cxx applies it at sprmPIlfo; Word writes sprms in ascending order, so Ilfo precedes the indent sprms); grpprlChpx
  formats the label; ixchFollow 0 tab / 1 space / 2 nothing (LibreOffice ww8par3.cxx); sprmPChgTabs(Papx) fill TabStops
  (one operand layout for both, as LibreOffice Read_Tab).
- .docx: the numbering level's w:pPr (indents, tabs) and w:lvlJc are stored with each interned list label (index format 9)
  and applied by the renderer between the style and the direct pPr. Numbering via the paragraph style keeps the style's
  own w:ind (LibreOffice DomainMapper_Impl finishParagraph: "the para-style indents have priority"); direct numPr lets the
  level override the style; direct w:ind wins over both. settings.xml w:defaultTabStop drives the default stop.
  Precedence taken from LibreOffice's source (it states it emulates Word) - not re-checked in Word.
- RTF: the \listtext label (Word's flat copy) is drawn by WordListLabelHtml at the paragraph's first-line position, its
  \tab = tab suffix; \leveljc places it (LibreOffice rtfdispatchvalue.cxx maps \leveljc 0/1/2 to lvlJc left/center/right
  and ignores \leveljcn); \levelfollow 0/1/2 = tab/space/nothing for labels counted at index time.
- ODT: label-alignment levels (ODF 1.2 19.880/19.881: fo:margin-left, fo:text-indent, text:list-tab-stop-position,
  text:label-followed-by) are stored per list label in the index (format 3) and indent the paragraph; the older
  label-width-and-position mode keeps the fixed per-level step.
- Line spacing: auto spacing is a multiple of the font's single line height (`WordParagraphFormatting.SingleLineHeightEm`
  = 1.15 em, Times New Roman / Arial), not of the em - LibreOffice draws 12 pt lines 13.8 pt apart at line=240; exact
  (docx lineRule exact, .doc / RTF negative dyaLine / \sl, ODF fo:line-height length) is the stated height; at least
  is `max(1.15em, value)`. Fonts with other ascent + descent (Hebrew faces) still differ - per-font metrics not modelled.

## In-document search (branch work/doc-search, 2026-10-29)

`IWordWindowedDocument.Search(query, WordDocumentSearchOptions, fromBlock, maxBlocksToScan)` returns
`WordDocumentSearchBatch(HitBlocks, NextBlock, ReachedEnd)`: one bounded step, block indexes only, nothing kept.
The caller loops from `NextBlock` until `ReachedEnd` (`NextBlock` = -1).

- **Windowed like rendering**: each batch reads its blocks through the reader's own windowed path (docx/odt: one
  contiguous read of the batch's byte range, one inflate pass; .doc: piece-table text; RTF: the block's byte range
  re-parsed from its index snapshot into a text-only sink; txt/MHTML: the block bytes). Text is fed straight into a
  KMP matcher - no XElement for body blocks, no HTML, no per-block strings kept.
- **On-demand index**: a batch past the known blocks extends the index exactly as `RenderBlocks` does (chosen over
  "search only the known blocks" so the caller needs no second loop). `ReachedEnd` only once the index is complete.
  The batch that completes the index also pays for the notes scan + sidecar save (778 ms worst batch on the
  87 MB sample); every other batch is ~10-40 ms at 500 blocks.
- **Normalisation = the app's book-view search** (KitveiHakodesh Core `SeforimDb/SeforimDbLineTextScan.cs`
  NormalizeQuery/IsIgnoredMark/FoldChar, and `vue-frontend/src/utils/hebrewTextProcessing.ts`
  SEARCH_IGNORED_MARKS / SEARCH_DOUBLE_QUOTES / SEARCH_SINGLE_QUOTES, both read 2026-10-29): U+0591-U+05C7 except
  the maqaf dropped; maqaf U+05BE = space; `" U+05F4 U+201C U+201D U+201E` = `"`; `' U+05F3 U+2018 U+2019 U+201A` and
  backtick = `'`; query trimmed; letters exact (the app is case-sensitive; `MatchCase=false` adds ToLowerInvariant).
  No final-letter folding (the app has none). MHTML blocks are HTML and walked exactly as the app walks a line's HTML.
- **What counts as text**: docx w:t only (w:delText, w:instrText not); w:tab/w:br/w:cr a space; paragraph, cell,
  w:br/w:cr, ODF line breaks, RTF \line and .doc vertical tab are separators too; w:pPr/w:rPr and mc:Fallback skipped (no match across them). ODF: text:p/text:h content, text:s = its count of
  spaces, note-citation / annotation / tracked-changes skipped. .doc: field codes hidden, results kept.
  RTF: \listtext labels skipped.
- **Notes** (`IncludeNotes`, default true): a note body holding the query makes the referencing block a hit
  (docx/doc read the note on demand; ODF/RTF notes are inline in the block).
- **Hit positions are not returned**: offsets in the plain text would not line up with the rendered HTML (list
  labels, field results, note marks), so the page highlights client-side with the app's highlighter, as the book
  view does. The `--search` check confirms on the samples' first 3000 blocks that hit blocks = blocks whose rendered
  HTML (+ notes) holds the query walked the app's way (0 differences, .doc and .docx).
- Timings (`--search <files>`, Release, 500 blocks per batch, cached index): .docx 87 MB body, 123,007 blocks:
  full scan 1.7-2.3 s (absent word), 2.3-4.2 s (word in 31,718 blocks), first hit 22-32 ms, avg batch 7-17 ms.
  .doc 40 MB, 106,442 blocks: full scan 0.4-0.7 s, first hit 6-12 ms, avg batch 1-3 ms. Cold on-demand .docx:
  open 0.2-0.75 s, full scan incl. building the index 6.5-8.4 s.

### .docx byte-level search (2026-11-03)

Profile first (sample above, Release, `DOTNET_TieredCompilation=0`, idle priority, 2 cores, best of 3 alternated, full
scan of a frequent word in batches of 500): reading the batches' bytes through the index ranges alone took 2.2 s - the
compact table rescans every span with the full `DocxBodyBlockScanner` (kinds, styles, list labels) just to find block
offsets; the raw sequential read of the same bytes is ~0.25-0.9 s (inflate). The XmlReader text pass on top: +1.8-2.5 s
(a node + a string per run). Notes: no measurable cost on the sample.

Now (`DocxSearchTextScanner`, `DocxCompactBlockTable.ReadSearchBytes`):
- **No span scan**: a batch reads its raw bytes from the view in one sequential pass, from the first block's start (an
  anchor, the previous batch's end, or - out of order - one span scan) to the end block's start (an anchor or the body
  end; otherwise on to the next anchor, the tail kept for the next batch). The scanner finds the block boundaries itself
  with the index pass's rule (every element starting at body level begins a block) and must come out at exactly the
  index's count at an anchor / the body end; any disagreement, an undefined entity or unterminated markup falls back to
  the index's per-block ranges from that block on. Read in pieces of 1000 blocks: a huge batch never holds more.
- **Text straight from the bytes**: tags skipped with vectorised searches (`>` outside quoted values), names compared
  only for the few local names that matter and resolved by namespace (root declarations + any declared in the block);
  w:t content decoded UTF-8 -> the matcher char by char (`FeedUtf8`), entities and character references decoded,
  CDATA kept, comments / PIs skipped. Same rules as before: w:t only (not w:delText, w:instrText, attributes), w:tab a
  space, w:br / w:cr / paragraph ends separators, w:pPr / w:rPr / mc:Fallback skipped. Markup inside a word (runs,
  proofErr, bookmarks, comment ranges, niqqud in its own run) adds no character, so the word still matches.
- The XmlReader pass stays as the **reference** (`DocxSearchTextMode.ReferenceXmlReader`): `--search <docx>` compares
  every block and count of 10 queries x match case on / off (Hebrew plain + pointed, two words, gershayim, maqaf,
  English, one letter, absent, mixed) over the whole sample - identical - and `DocxSearchBytesTest` checks every
  split-markup case against both the expected hits and the reference. Notes are still read as XElement (unchanged).
- Same conditions as the profile: full scan 4.0-5.3 s (reference) -> 0.67-0.98 s; first batch of 500 3 ms;
  `--search` timing: frequent 981 ms (31,718 hit blocks), absent 672 ms; cold on-demand open + full scan 2.0 s (was 3.8).
  While the on-demand index is building the batches still read per-block ranges (now with the byte scanner).
- Piece verdict (`WordSearchPiece.Judge`, shared with .odt): bytes that end exactly at an anchor / the body end must come
  out Done; Stopped inside them (an extra block) or Done on bytes that run on means the scanner's blocks are not the
  index's - the piece is redone from per-block ranges, never kept. A count mismatch / stray close tag now fails from the
  piece's FIRST block and drops the piece's hits from there (before, hits of earlier blocks of the piece were kept).

### .odt / .fodt byte-level search (2026-11-03)

`OdtSearchTextScanner` + `OdtCompactBlockTable.ReadSearchBytes`, same model as .docx: a batch reads raw bytes from an
anchor (its scan state gives the frames open there - `OdtBodyBlockScanner.SearchFramesOf`) or from where the previous
batch stopped (offset + frames kept), in pieces of 1000 blocks. Block boundaries follow the body scanner's rule (text:p /
text:h / table:table at block level: office:text, section, numbered-paragraph, index-body, list item; lists, sections,
tracked changes, sequence decls, index templates are not blocks), names between blocks classified by the root's prefixes
like the body scanner, inside blocks by namespace like the reader. Text rules = `FeedOdf` exactly: paragraph text only,
white space collapsed with the block-wide "last was a space" state, text:s c, text:tab, text:line-break, p / h
separators, note-citation / tracked-changes / office:annotation / (notes off) note-body skipped. A start not at an anchor
or resume point reads per-block ranges up to the next anchor. `FeedOdf` stays as the reference (`OdtSearchTextMode`).
- Golden (`--search <odt|fodt>`, `OdtSearchBytesTest`): 10 queries x match case on / off on the LibreOffice conversions of
  the sample (.odt 121,650 blocks, .fodt): identical blocks and counts; batches of 37 = 5000; a start mid-document.
- Release, TieredCompilation=0, idle, 2 cores, best of 3 alternated, full scan of a frequent word in batches of 500:
  .odt 12.2 s -> 0.60 s, .fodt 11.2 s -> 0.42 s; first batch of 500 after a fresh open 35.6 -> 2.5 ms (.odt), 5.7 -> 1.4 ms (.fodt).
- Not done: the on-demand (still-building) index reads per-block ranges (with the byte scanner).
- **Search = rendered text** (review, 2026-11-03): svg:title / svg:desc of a draw:frame were found though never shown.
  Both the byte scanner and `FeedOdf` now follow `OdtBlockHtmlRenderer` (`WordXmlSearchText.IsHiddenOdf`, one rule set):
  inside a paragraph only the text namespace, draw:a and draw:frame render - svg:*, shapes, draw:image (and paragraphs
  inside it), controls, objects and every other namespace (loext:...) are skipped whole, as are text:ruby-text /
  text:hidden-text / text:hidden-paragraph; of a frame only its draw:text-box, of a text box only its text:p / text:h
  (a list or table inside a text box is not rendered). Character data counts only directly in paragraph content (not
  white space between a frame's / note body's children). `OdtSearchBytesTest.RenderedTextOnly` also compares with the
  rendered HTML. Left as is: text:number (rendered only when the paragraph has no list label).
  .docx checked the same way: wp:docPr / pic:cNvPr title, descr, name and VML alt are attributes (never text); text
  boxes count once (mc:Fallback skipped) - covered by `DocxSearchBytesTest`. Open .docx gaps (not drawings): w:t inside
  w:moveFrom and inside w:ruby is searched but the renderer does not show it.
- **Drift guard on a failing piece**: a piece that fails at block k (undefined entity, unterminated markup) keeps its
  hits before k only when the scanner's numbering is verified - block k and every hit block before it start where the
  index says (`WordSearchPiece.FallBackBlock`, one span scan, failure path only); otherwise the whole piece is read
  again from per-block ranges. Same for .docx.
- **Bounded buffers**: the search read buffer (both formats) and the .docx note buffer drop back after a batch that grew
  them past 4 MB (`WordSearchPiece.NormalBufferBytes`); growth by doubling stops at that cap.

### .docx note bodies from bytes (2026-11-03)

Notes were read per search as XElement (range from the notes index, parsed whole). Now `DocxNoteSearchTextScanner`
reads the same range's raw bytes into one reused buffer (`DocxNoteXmlReader.ReadBytes`) with the reference's rules
(`FeedOoxml(XElement)`): per w:p its w:t text, whitespace-only text nodes dropped unless xml:space="preserve" (the parser
ignores whitespace), w:tab inside a run a space, w:br / w:cr / paragraph end separators. It declines - matcher rolled
back, the note parsed as before - a paragraph inside a paragraph (text box: the reference reads it twice), w:drawing /
w:pict / w:object (stripped before parsing), whitespace written only as references, malformed XML. .odt notes are inline
and already go through the body scan.
- Synthetic 100,000-footnote docx (`--search notes-heavy`): 8 queries x match case identical to the parsed path
  (blocks + counts); full scan with notes 8.5 s -> 1.0 s (body alone 0.1 s), same conditions as above.
## Bookmarks and in-document links in every format (branch work/bookmarks-all, 2026-10-29)

`IWordWindowedDocument.FindBookmark(name)` returns the block a bookmark marks; links render as
`href="#bookmark-<Uri.EscapeDataString(name)>"` (one helper: `DocxBlockHtmlRenderer.LinkTag` /
`BookmarkElementId`; field instructions parsed by `DocxBlockHtmlRenderer.HyperlinkFieldTag` for
.docx, .rtf and .doc alike). Same safety rule everywhere: only http(s), mailto, in-document anchors
and the app schemes (otzaria, zayit, kitveihakodeshapp, seforimapp) become links.

| Format | Bookmarks (source) | Index | Links |
|---|---|---|---|
| .docx | w:bookmarkStart (done earlier) | v10 | w:hyperlink, HYPERLINK, REF/PAGEREF/NOTEREF \h |
| .odt / .fodt | text:bookmark, text:bookmark-start, text:reference-mark(-start), text:name (ODF 1.3 part 3 6.2.1 / 6.2.2), captured by the body scanner | v4 | text:a xlink:href (6.1.8); text:bookmark-ref / text:reference-ref text:ref-name (7.7.5 / 7.7.6) |
| .rtf | {\*\bkmkstart name} (RTF 1.5 spec "Bookmarks"), captured by the parser | v3 | {\field{\*\fldinst HYPERLINK / REF \h}{\fldrslt}} (spec "Fields"; the HYPERLINK instruction itself is Word field syntax, not in the 1.5 spec) |
| .doc | SttbfBkmk + PlcfBkf (FibRgFcLcb97 pairs 21 / 22, LibreOffice ww8scan.cxx WW8Fib read order and WW8PLCFx_Book) - small tables read at open, CP -> block by the block index; nothing stored, no index change | v4 (unchanged) | 0x13 / 0x14 / 0x15 field marks in the text; the code collected, control characters (Word's 0x01 hyperlink-data mark) dropped |

- On-demand indexes (.docx, .odt) scan on only until the bookmark turns up (`WordDocumentBookmarkScanLookup`).
- `_GoBack` is skipped (.docx, .doc). Reference marks share the bookmark lookup (first name wins).
- Known limits (fixed 2026-11-04: .doc and .rtf now restore the fields open at a block start, see "Word fields"); .doc bookmarks in notes / headers are not targets; RTF bookmarks inside a
  footnote mark the block holding the note.
- Harness: `LinksAllFormatsTest` (odt, rtf), `DocBinaryLinksTest` (.doc written by LibreOffice from a
  small .odt; skipped without LibreOffice), in `--fidelity`.

## OpenDocument: huge styles.xml on the first open (2026-10-29)

A LibreOffice conversion of a Word file carries thousands of list styles (sample
dor-hamelaktim-shabbat.docx -> .odt: styles.xml 15.4 MB, 3,058 text:list-style x 10 levels, 5,716
style:style, 993 master pages). The first open built all of office:styles as one XML tree, built
every list style from it and inflated styles.xml again for the page setup: ~0.7-1.5 s before text.
Now `OdtPackage.ReadStyles` is one streaming pass: `OdtStylesListStyleFilterStream` cuts every
text:list-style (16.32) out of the bytes and parses it from its bytes (`OdtListStyle.Parse`; equal
level definitions parsed once), XmlReader sees the remaining 2.1 MB, the page setup is collected in
the same pass, and the on-demand open builds its rendering cascade from the filtered trees (no
serialise + parse). Allocation 134 MB -> 18 MB per open. Sample, warm, best of 10: ~700 ms -> 189 ms
to the first 40 blocks; `--fidelity` checks a synthetic 3,000-list-style file under 300 ms (104 ms)
and `--odt-first-open <path>` measures any file. All 30,580 list levels of the sample read the same
as the tree-based reading. Not done: a .fodt keeps its styles inside the body file and still captures
office:styles as a tree.

## Pictures, tables, list gaps (2026-11-03)

Sources: ECMA-376 Part 1 (DrawingML wp:inline / wp:anchor 20.4.2.x, tables 17.4.28 / 17.4.51), ODF 1.3 part 3 (20.414
table:align, 20.400 style:wrap, 20.298 style:horizontal-pos, 19.759 text:anchor-type - read in the OASIS HTML),
LibreOffice core master: writerfilter GraphicImport.cxx (wrapNone = THROUGH with every gap dropped; "when impossible to
wrap text on a side, remove the margin gap"), GraphicHelpers.cxx (wp:align left / right / center are physical),
ConversionHelper.cxx convertTableJustification (center, right = end = RIGHT), DomainMapperTableHandler.cxx (compatibilityMode
< 15: table edge = tblInd - first cell's left margin), DomainMapperTableManager.cxx (bidiVisual = RL_TB), rtftok
rtfdispatchvalue.cxx TRLEFT / PICWGOAL, rtfdispatchflag.cxx TRQ* / PNGBLIP / NONSHPPICT, rtfdocumentimpl.cxx (first cell
= cellx - trleft, goal size x scale), core/text/txttab.cxx NewTabPortion (a tab after a label goes to the next stop after
the label's END), core/text/txtfld.cxx (a bullet portion is the bullet character only).

- **The host must wrap the block HTML in `.word-document`.** Every rule of the stylesheet is scoped to it; the app page
  never had the class, so none of it applied (list labels were plain inline spans: min-width ignored = no gap; table
  borders / collapse, paragraph margins). The page now puts it on each block's text and each note body.
- **List labels**: a tab suffix keeps at least 0.25 em after a label wider than its gap (Word jumps to the next stop; the
  label's width is unknown server-side). ODF bullets drop num-prefix / num-suffix (LibreOffice writes the bullet as
  num-suffix too: labels came out doubled).
- **Pictures** (`WordPictureHtml`, shared): an `<img class="word-picture">` with a data: URI read on demand from the package
  when the block renders (never kept; .png / .jpeg / .gif / .bmp / .webp / .svg - EMF / WMF / TIFF keep a dashed box of
  the right size). Inline = in the line; square / tight / through = CSS float to wp:align's side with dist* as margins
  (none on the side without text); centred square and topAndBottom = block; wrapNone (in front / behind) = block without
  gaps. The page makes each block a formatting context (`display: flow-root`): a float never runs into other rows.
  .docx: `DocxDrawingStripper` keeps the layout as one `<fastDocViewPicture>` element (wp: / a: / r: prefixes as Word
  writes them); relationships of document / footnotes / endnotes. .odt: draw:frame + draw:image (package entry by path,
  or office:binary-data in .fodt), graphic styles now kept in the index (format 7). RTF: {\pict} / {\*\shppict} inline
  at \picwgoal x \picscalex (renderer sink only; index / search keep skipping). Not done: floating RTF \shp, VML
  w:pict images, .doc PICF pictures, posOffset relative to the page / margin (treated as from the column's left).
- **Tables**: .docx w:jc center / end, tblInd from the table's own start edge (a bidiVisual table starts at the right;
  `margin-inline-end:auto` so a table never takes the paragraphs' side), minus the cell margin before Word 2013. .odt
  table:align left / right / center / margins with fo:margin-left / -right (physical). RTF rows: 	rleft indent, 	rqc /
  	rqr, row width = last \cellx - 	rleft (index format 6: the snapshot carries them). The stylesheet's 2 pt table margin
  is gone (Word has none). .doc rows unchanged (full width).
- Checks: `--fidelity` PagePlacementFidelityTest (all formats); live `playwright/layoutCheck.mjs <file>` (label gaps, the
  column, tables, pictures, ignore-margins) and `playwright/notesLayoutCheck.mjs <file-with-notes>` (pane 25 %, notes
  column = text column).

## Re-shown windows: no needless re-inflation (2026-11-03)

Measured with `--reread <files>` (RereadCostTest: sidecar mode; 20 windows of 40 blocks from the middle read forward,
paged back, forward again; notes: a jump to the middle note window, 20 windows of 10 notes paged back, the jumped-to
window again; counters `DeflateRandomAccessStream.ProcessCheckpointResumeCount` / `ProcessBytesInflated`).
Checkpoint resumes / MB inflated, before -> after:

| | .docx (91.6 MB document.xml) | .odt (LibreOffice conversion of the same book) |
|---|---|---|
| body, first read (20 windows) | 7 / 3.26 -> 3 / 1.51 | 8 / 15.19 -> 3 / 5.77 |
| body, paged back | 18 / 5.99 -> 4 / 0.71 | 20 / 36.13 -> 0 / 0.00 |
| body, forward again | 18 / 5.99 -> 0 / 0.53 | 0 / 0.33 -> 0 / 0.00 |
| notes, jump | 1 / 0.21 -> 1 / 0.21 | 1 / 1.10 -> 1 / 1.10 |
| notes, paged back (20 windows) | 11 / 1.22 -> 3 / 0.32 | 6 / 5.18 -> 1 / 0.93 |
| notes, jumped-to window again | 0 / 0.00 -> 0 / 0.00 | 0 / 0.30 -> 0 / 0.16 |

What was wrong and what changed:

- **The docx fallback read was not retained.** When a span was already scanned and the ring held other bytes, a
  window's bytes came straight from the view and were dropped; the next window (which starts one block back, the
  neighbour read of `RenderBlocksLocked`) began outside the ring again and seeked back to a checkpoint.
- **The .odt ring restarted on any read that did not continue it exactly**, so a page back (a range just before
  the ring) re-inflated about one checkpoint span.
- Both tables now keep their bytes in one `WordIndexByteRing`: one contiguous range, extended at **either end**
  (append, prepend; full = the far end goes), restarted only by a read elsewhere. `Fill` serves a range from the
  ring and reads only the missing part from the view (up to where the ring begins), retaining it. A read or a span
  scan that starts a little past the ring's end (<= 64 KB: the next block of a window, the next span's anchor after
  bytes the previous scan had not reached) reads the gap first — the view is there anyway — so the range stays one
  piece. Bounds unchanged: .docx 512 KB, .odt 512 KB growing to 4 MB for a long read; a read longer than the ring
  keeps its last bytes.
- **SeekableDeflate: a seek back served from the inflater's history.** `DeflateRandomAccessStream` resumed from a
  checkpoint for any backward seek, even a few bytes. The inflater's output buffer always holds at least the last
  32 KB it delivered (`HistoryStart`; after a resume only from the checkpoint on, since a sealed window is sparse);
  a read at an offset in [HistoryStart, TotalOut) is now copied from there (`CopyFromHistory`) without moving the
  inflater, and the read after it continues inflating. Harness: back-seeks of 0, 1, 32 KB-1, 32 KB served with no
  resume, 32 KB+1 and before a resumed checkpoint correct either way.
- Remaining .docx body page-back resumes: 20 windows of this book are ~600 KB, more than the 512 KB ring; the
  earliest windows of the page back come from a checkpoint once. Raising the bound is a memory decision, not taken.

## Word save while we read: intermittent refusals are the environment (2026-11-03)

`--real`'s "file can be replaced (Word's save)" failed 2-5 times per run. `--coexistence-probe <file> <rounds>`
repeats copy + `File.Replace` with **nothing of ours open** (control) and with our document open: 4 of 60 and 3 of 60
refused, both with 0x80070497 (ERROR_UNABLE_TO_REMOVE_REPLACED), and the Restart Manager (RmGetList) found no holder
by the time it was asked — a transient handle on the freshly written 40 MB files from another process (a scanner or
indexer; Defender's real-time protection is off on this box, SearchIndexer runs). The exclusive-open check just
before the replace always passes, and our process was never among the holders. Not our handle. The test now names
the holders on a refusal, fails if our process is one of them, and retries (100 ms steps, up to 10 s) only while
others hold the files, as Word retries a save.

### Pictures by id, .doc pictures and rows, VML, RTF shapes (2026-11-03)

- **No bytes in the block HTML.** `<img class="word-picture" data-word-picture="ID" width height loading="lazy"
  decoding="async">` without src; `IWordWindowedDocument.ReadPicture(id)` reads that one picture (windowed). Ids:
  .docx `m:|f:|e:` + r:id (only image relationships resolve); .odt `p:` + package entry (picture extensions only) or
  `b:block:n` (office:binary-data, the block re-read); RTF `block:n` (the block re-parsed, only that picture's data
  decoded - a normal render decodes no picture data); .doc the PICF offset in the Data stream. The host maps the id to a
  URL: app dev = service `GET /fast-doc-view-picture/<doc>/<id>` (doc id = capability), hosted = RPC
  `fastDocViewPicture` -> object URL, demo = dev server `/api/documents/{id}/pictures?pictureId=`.
- **.doc**: 0x01 + sprmCFSpec + sprmCPicLocation (not sprmCFOle2 / sprmCFData) -> PICF (lcb, cbHeader, mm 0x64 / 0x66,
  dxaGoal x mx / 1000 less crops) -> OfficeArtSpContainer -> FBSE -> blip 36 bytes in (LibreOffice ww8par.cxx case 0x1,
  ww8par6.cxx Read_PicLoc, ww8graf2.cxx PicRead / WW8PicDesc, ww8struc.hxx WW8_PIC_SHADOW, msdffimp.cxx GetBLIPDirect /
  ReadCommonRecordHeader). PNG / JPEG only. Rows: sprmTJc / TJc90 & 3 (0 left, 1/3 centre, 2 right, physical), width =
  sprmTTableWidth (fts 3) else the rgdxaCenter span, indent = rgdxaCenter[0] moved by sprmTDxaLeft - dxaGapHalf, applied
  from the left (LTR, left) or right (sprmTFBiDi, right) as ww8par2.cxx LEFT_AND_WIDTH. LibreOffice-written .doc: mx
  overflows 16 bits for big upscales (wrtww8gr.cxx WritePICFHeader) - the test uses a small extent.
- **.doc floating pictures** (DocBinaryFloatingPictures, 2026-11-03): 0x08 + fSpec in the main text -> PlcSpaMom
  (FibRgFcLcb97 40) FSPA at that CP (26 bytes: spid, xaLeft..yaBottom, flags bx / by / wr / wrk / fRcaSimple) ->
  OfficeArtContent at fcDggInfo (index 50): BStore FBSEs (foDelay) and each DgContainer's SpContainer FSP spid + FOPT
  pib (260) -> blip in the Data stream, else WordDocument. Read from LibreOffice master: ww8par.cxx:3656 (case 0x8),
  ww8par.cxx:501 (streams), ww8graf.cxx:2605-2662 (Read_GrafLayer) and 2700-2745 (wr / wrk), ww8struc.hxx:903-964
  (WW8_FSPA), ww8scan.cxx:1808-1811 / 6117 / 6138, msdffimp.cxx:5967-6095 (GetCtrlData, BStore), 6215-6245 (FSP, FOPT),
  6414-6475 (GetBLIP), dffpropset.cxx:1128 (id & 0x3FFF). Id = "s" + spid. Wrap 1 = above and below, 3 = through,
  else square (wrk 1 = picture right); size from the FSPA; xaLeft kept as the offset when bx is column / margin.
  Not drawn: yaTop / page origin, wrap gaps (FOPT dxWrapDist*), header shapes (PlcSpaHdr); groups, shapes and text boxes: see .doc text boxes and shapes. Check: `--fidelity` DocBinaryFloatingPicturesTest (POI PngPicture.doc,
  FloatingPictures.doc, pictures_escher.doc). Golden dumps change only on those three files.
- **VML** (w:pict / w:object): v:imagedata r:id (o:relid), shape style width / height / position /
  mso-position-horizontal / margin-left / mso-wrap-distance-*, w10:wrap type / side, read as oox vmlshapecontext.cxx and
  vmlshape.cxx lcl_setSurround / lcl_SetAnchorType (static = inline; no wrap = through).
- **RTF \shp** with a pib picture: \shpleft..\shpbottom size, \shpwr / \shpwrk (RTF 1.5 spec + rtfdispatchvalue.cxx), posh,
  dxWrapDist* EMU (rtfsdrimport.cxx); \shprslt skipped, \shptxt read (see RTF text boxes and shapes); no \shpwr = above and below (our choice). 	rleft is now
  a physical margin-left (RTF 1.5: "leftmost edge ... with respect to the left edge of its column").
- Not done: floating .doc pictures (0x08 + FSPA + drawing group, e.g. POI PngPicture.doc), DIB / metafiles, page-relative
  offsets.


### OpenDocument drawing (2026-11-04) - supersedes the shape parts of the 2026-11-03 section below

Sources: LibreOffice core master read 2026-11-04 (cited file:line in each class), ODF 1.3 part 3 (OASIS HTML).
- **One custom-shape engine for every format** (`Html/WordCustomShapeEngine.cs`, port of svx EnhancedCustomShape2d
  SetPathSize / GetParameter / CreateSubPath / CreatePathObj): every segment command incl. G (arc angle), Q, H..K, checkClosed,
  stretch points (nXRef / nYRef - msoRoundRectangle keeps circular corners), colour data (cube, can, ... faces), fill + stroke
  objects and SortFilledObjectsToBack. Formulas (`WordCustomShapeFormula.cs`) follow EnhancedCustomShapeFunctionParser's grammar;
  OfficeArt calculation records become formulas as GetEquation does, so .doc / RTF presets now share it (expected-value change:
  the RTF round rectangle 72 x 36 has 6 pt corners, not 12).
- **ODF geometry** (`Odt/OdtEnhancedGeometry.cs`): enhanced-path, equations by name, modifiers, text-areas, view-box,
  sub-view-size, mirror, text-rotate-angle; presets by draw:type merged per property (MergeDefaultAttributes); "rectangle" is a
  plain rectangle; an unknown type without a path draws nothing. 3-D extrusion (draw:extrusion="true", `OdtExtrusion`) is drawn as LibreOffice's
  EnhancedCustomShape3d::Create3DObject builds its scene: each outline object extruded (back z 0, front z depth), the scene matrix
  (centre, shape rotation, mirroring, rotation angles about the rotation centre, skew shear for parallel), parallel or perspective
  projection (viewpoint + origin), back faces culled unless double-sided, and one flat colour per face from solveColorModel with the
  two lights, soft lights, ambient brightness, diffusion, specularity, shininess and metal (lights as 8-bit colours). The faces replace
  the outline, back to front; the box is placed by its centre without rotating or mirroring again. Needs a real 3-D renderer, so
  approximated: Gouraud / Phong (flat colour per face instead), the z-buffer (painter's order by mean depth - differs only for concave
  outlines), lit bitmap / gradient textures (front face drawn with the unlit fill), wide-line contours (centre line extruded). Fill / caps follow the shape's fill style; a point at the viewpoint's depth is skipped; projection / shade-mode read in draw: or dr3d:. Checked by `fidelity-v1-odt-extrusion.fodt` (parallel, perspective, rotated, draft, unfilled, shaded) in --fidelity and live by `vue-demo/playwright/odtExtrusionCheck.mjs`.
- **Objects** (`Odt/OdtShapeBuilder.cs`): rect (draw:corner-radius only, as LibreOffice), circle / ellipse kinds, line, polygon,
  polyline, path (open = not filled), connector (svg:d when its ends match, else `OdtConnectorRoute` = ImpCalcEdgeTrack; ends
  glued to shapes are routed around those shapes: every shape with an id - LibreOffice writes one only on a shape something refers
  to - has its snap rectangle and draw:glue-points recorded at index time (`OdtConnectedShapeTable`, index format 10; the scanner
  notes the blocks holding one, they are read once when the index is final), so a shape in another block is never read at render
  time; a shape in other coordinates (another paragraph's anchor) is moved to where svg:x1/y1 or x2/y2 put the end; a shape not
  known yet on a first open = a free end, and its block is marked provisional (`WordDocumentBlockHtml.IsProvisional`):
  the app drops those blocks when the index completes and fetches the ones on screen again (`odtConnectorsCheck.mjs --first-open`). Word formats need no routing: LibreOffice turns a .docx wps connector into a routed
  ConnectorShape only inside a drawing canvas (oox/source/shape/WpsContext.cxx:585-592, one drawing = one block) and reads the
  .doc/.rtf solver container only for PowerPoint (filter/source/msfilter/svdfppt.cxx:1603-1606; ww8par.cxx / ww8graf.cxx never read
  it) - Word and LibreOffice draw those connectors from their own stored geometry, which we already do), measure
  (`OdtMeasureShape`; the server lays it out with a 0.5 em-per-character label estimate, the app page lays the dimension line and
  label out again with the label's measured size - `data-measure` inputs, vue-frontend `fastDocViewShapeText.ts`), caption (`OdtCaptionShape`), regular-polygon (not imported
  by LibreOffice; drawn per ODF 19.123 / 19.128 / 19.211), groups; draw:transform as a full matrix (rotate / skew / scale /
  translate / matrix).
- **Paint** (`Odt/OdtShapeStyle.cs`, `Odt/OdtFillPaint.cs`): LibreOffice's pool defaults when unset (fill solid #729fcf, line
  solid #3465a4 hairline, round join); gradients linear / axial / radial / ellipsoid / square / rectangular with border, centre,
  angle units, intensities, loext gradient stops, step bands; transparency gradients as SVG masks (also over picture fills);
  hatches single / double / triple with background; bitmap fills stretch / repeat / no-repeat with ref point, offsets, row /
  column shift; dashes (XDash::CreateDotDashArray, the pool's default dash), joins (none = segments stroked apart), caps,
  markers from draw:marker paths with width / centre and the notched-back rule. Automatic steps (0) are banded like a set count:
  calculateNumberOfSteps (colour distance x 127.5) bands per stop range, as the primitive decomposition draws them
  (drawinglayer texture.cxx; the Windows VCL path) - only the Cairo renderer draws them smooth. Unitless draw:angle of
  gradients / transparency gradients: tenths of a degree before ODF 1.2 and in ODF 1.2 written by OOo / AOO or LibreOffice
  before 7 (tdf#89475, GradientStyle.cxx:115-128), decided per document from meta:generator (meta.xml / office:meta, read at
  every open - a few hundred bytes - and classified as xmlmetai.cxx setBuildId + xmlimp.cxx DocumentInfo: `Odt/OdtGeneratorVersion.cs`).
  Checked in --fidelity by `fidelity-v1-odt-angle-*.fodt` / `.odt` (OOo 3.2 / LibreOffice 7.3 / no generator).
- **Text**: text areas with padding, vertical / horizontal alignment, writing mode; Fontwork through the shared
  WordOfficeArtFontwork (align from draw:fit-to-size / textarea-horizontal-align; LibreOffice does not use text-path-mode);
  text boxes: vertical alignment, painted fills, numbered lists and tables (a paragraph holding a box table becomes
  `div.word-paragraph-block` - a table inside a `<p>` is split by the HTML parser).
- **Shapes that grow with their text** (2026-11-04): a draw:custom-shape with draw:auto-grow-height (absent = on: the item pool's
  default, svdattr.cxx:247 - no import or Writer code sets another, see AutoGrowAttributes) carries `data-autogrow*` on its text box
  (text area in its own points, size, fo:padding, fo:min-width / min-height, alignment, its matrix, mirroring, turn,
  draw:text-rotate-angle). The app page (`fastDocViewShapeText.ts`) measures the text in the unturned area and resizes the shape as
  SdrObjCustomShape::AdjustTextFrameWidthAndHeight + ImpCalculateTextFrame do (grow or shrink; direction by text alignment; width
  too for no-wrap text), in the shape's own coordinates before its transform (a rotated rectangle moves with its turned top left,
  svdoashp.cxx:2420-2428), mirroring redone about the new centre; the outline stretches (non-scaling strokes). Every frame the page
  touches (`data-shape-page`) then takes the new bounds of all its shapes and dimension lines (groups grow by their members) and a
  contour wrap is swept again from the new outlines. Other objects (rect, ellipse ...) are not text frames in LibreOffice and never
  grow. Shape text boxes are not clamped by the stylesheet's `.text-box { max-width: 100% }` (a turned area is wider than its frame).
- **Dimension lines**: the arrows are the shape's own markers (their width and length reach the layout,
  sdrmeasureprimitive2d.cxx:150-170), drawn per piece end (help lines carry none); the page redraws the pieces with their markers
  and sizes the frame from the measured label. Checks: `--fidelity` OdtShapeTextFidelityTest, live
  `vue-demo/playwright/shapeTextCheck.mjs TestOutput/fidelity-v2-odt-shape-text.fodt`.
- **Not possible / not done, with reason**: contour-mode "full" and text on both sides of a float cannot be laid out in CSS (CSS Shapes 1 s.2); see "Two-sided and contour wrap (all formats)".
- **Wrap**: style:wrap-contour = CSS shape-outside polygon from the outline (1 pt sweep) with shape-margin; page-anchored objects
  at body level are their own paragraph blocks (index format 9, which also keeps every graphic property and drawing style).
- **Search**: shapes of every kind, lists / tables / sections in text boxes, lists in shapes, body-level drawings - once.
- Checks: `--fidelity` (OdtDrawingFidelityTest, OdtConnectorRouteTest, OdtMeasureCaptionTest; LibreOffice conversion compared),
  live `vue-demo/playwright/odtDrawingCheck.mjs`.

### OpenDocument text boxes and shapes (2026-11-03)

Sources read online 2026-11-03: ODF 1.3 part 3 (OASIS HTML): 19.759 text:anchor-type, 20.400 style:wrap, 20.298
style:horizontal-pos, 20.397 / 20.398 style:vertical-pos / -rel, 19.240 fo:min-height ("overrides the height of a text box
specified by the svg:height"), 20.118 draw:fill, 20.119 draw:fill-color, 20.167 draw:stroke, 20.408 svg:stroke-color,
20.410 svg:stroke-width, 19.127 draw:corner-radius, 19.229.3 draw:type ("shall not influence the geometry of the shape").
LibreOffice core master: xmloff XMLTextFrameContext.cxx:637-661 (SizeType::MIN only for text boxes), :950-1010 (fo:min-width
/ fo:min-height read as the size); xmloff sdpropls.cxx:99-119 (draw:stroke, stroke-dash, svg:stroke-width / -color,
draw:fill, fill-color, fill-gradient-name); xmloff GradientStyle.cxx:101-104 (start / end colour); svx
EnhancedCustomShapeTypeNames.cxx:36-38, 48, 101-104 (preset names) and EnhancedCustomShapeGeometry.cxx (round-rectangle
174-216, right-arrow 684-724, left 726-753, down 761-784, up 792-819, left-right 827-868, default adjustments 46-89).

- **Text boxes** (draw:frame + draw:text-box) are placed boxes through the shared `WordPictureHtml.OpenTextBox` /
  `AppendPlacement` (one placement model with pictures and .docx boxes; `OdtBlockHtmlRenderer.Placement`). Size = svg:width /
  svg:height; fo:min-height / fo:min-width on the text box, else in the graphic style (LibreOffice 25.x writes it there) =
  a minimum, the box grows with its text. fo:padding(-*), fo:border (first side), fill: draw:fill solid / gradient, else
  fo:background-color. Its paragraphs go through the normal paragraph renderer as block spans (`text-box-paragraph`; no
  `<p>` inside a `<p>`); lists / tables in a box are still not drawn (search skips them alike).
- **Placement**: as-char = in the line (inline-block); otherwise wrap left / right / parallel / dynamic / biggest = float,
  none = above and below, run-through = block without gaps (as pictures). horizontal-pos is physical (left / right / center /
  from-left + svg:x); without one (and no one-sided wrap) the box goes to the paragraph's start side, right in a
  right-to-left paragraph - our choice, the spec gives no default. LibreOffice writes an explicit from-left for every box it
  saves, so its files are physical. vertical-pos from-top with svg:y (relation paragraph / char / line) is added to the top
  gap; page / frame relations and other vertical positions are ignored (no page in the flow view).
- **Both-sides wrap is not possible in CSS.** CSS Shapes 1, section 2 "Relation to the box model and float behavior"
  (drafts.csswg.org/css-shapes, read 2026-11-03): "this does not allow content to wrap on both sides of a float"; "Left floats
  with a shape-outside still only allow content wrapping on the right side". Fallback: a centred object with text on both
  sides (style:wrap parallel / dynamic / biggest) floats flush to its side and the text runs along the other side (user,
  2026-11-04: never top-and-bottom); see "Two-sided and contour wrap (all formats)".
- **Shapes** (draw:custom-shape, draw:rect, draw:ellipse, draw:line, draw:g) are inline SVG in a placed box
  (`text-box word-shape`, positioned; the SVG fills it): bounding box from svg:x / y / width / height, a line's svg:x1..y2,
  a group around its drawn children (each child an absolutely placed part at its offset). Fill: none / solid / gradient
  (draw:gradient start -> end colour at draw:angle as one linear gradient: axial / radial / border / steps approximated; the
  angle's orientation is NOT source-verified); hatch / bitmap = none. Stroke: draw:stroke none / solid / dash (one fixed
  dash pattern; draw:stroke-dash styles not read), svg:stroke-color, svg:stroke-width (min 0.5 pt). Unset stroke = solid
  black and unset fill = none: our choice. Geometry: draw:rect (svg:rx / draw:corner-radius), draw:ellipse, draw:line, and
  custom shapes by draw:type - LibreOffice's presets rectangle, round-rectangle, ellipse, right / left / up / down /
  left-right-arrow at their default adjustments (the calculation slots of the arrows other than right-arrow are assumed like
  mso_sptArrowCalc, NOT verified; draw:modifiers and draw:enhanced-path are not evaluated). Any other type = a dashed box of
  the right size. A shape's text:p / text:h sit on top from the box's top with its padding (draw:textarea-vertical-align not
  read); the box grows when the text does not fit. Polygon, path, connector, control, object: still not drawn.
- **Search** = rendered text (`WordXmlSearchText.IsHiddenOdf`, shared by the byte scanner and the `FeedOdf` reference): new
  contexts Shape (custom-shape / rect / ellipse / line: only its text:p / text:h) and ShapeGroup (draw:g: only drawn shapes),
  so shape text is searched exactly once; text-box text as before; hidden text (text:hidden-text, ruby text) never.
  LibreOffice 25.x saves a source text:hidden-text as plain text (shown and found there).
- Index format 8 (odt): the reduced graphic styles keep the fill / stroke / padding / border / min-size / vertical-position
  properties, and office:styles keep draw:gradient (name, colours, angle). Windowed: a shape is read from its block's XML
  only when that block renders.
- Check: `--fidelity` OdtShapesFidelityTest - a hand-written .fodt with each object type and the same document saved by
  LibreOffice headless (converted into TestOutput, never committed): placement, sizes, look, group offsets, search counts.

## RTF / ODT / FODT: no needless re-reads (2026-11-03)

Measured with `--read-counts <files>` (FileReadCountTest: sidecar mode, reopened; counters
`WordDocumentFileStream.ProcessOpenCount` / `ProcessBytesRead` / `ProcessSourceChecks`). RTF = LibreOffice
conversion of the sample book (212 MB, 123,183 blocks). Before -> after:

| | .rtf book | .fodt | .odt |
|---|---|---|---|
| first window (30 blocks) | 0.264 -> 0.074 MB | same | same |
| random jump (30 blocks) | 0.283 -> 0.094 MB | same | same |
| notes, 10 pages of 10 | 99 opens / 99 stats -> 10 / 10 | 5 / 99 -> 4 / 10 | 1 / 99 -> 1 / 10 |
| full search, batches of 500 | 422 MB (2.0x file) -> 211 MB (1.0x) | same | same |

- **Resumable RTF span rescan** (`RtfRandomAccessIndexBuilder.ResumableSpan` / `ScanThrough`, model
  `MarkupSpanRescan`): the cached span's parser pauses right after the asked-for block ends
  (`RtfParser.RequestPause` / `Continue`; each `Fill` seeks, so the stream may be shared meanwhile) and
  resumes on the next request; the block count is checked when the span completes. A window at a span's
  start reads one 64 KB buffer, not the span's 256 KB.
- **RTF search in one parse**: the span rescan feeds the matcher as it finds the bounds
  (`RtfSpanSearch`, `RtfCompactBlockTable.Search`); a sequential batch resumes the paused rescan. Text that
  the index pass would only count is decoded for the matcher and counted by its byte length
  (`RtfParser.TextByteCount`), so block values do not change. Old path kept as the reference:
  `RtfWindowedDocument.SearchByBlockReparse`; `--rtf-search-golden <files>` (RtfSearchOnePassTest) checks
  stepped rescan = whole rescan and one pass = old path (8 queries incl. Hebrew plain / pointed, English,
  mixed; match case and notes on / off; batches 500 / 37 / 1; random starts).
- **Notes pages**: RTF and ODT/FODT `RenderNotes` run all the page's block runs under one lock, one source
  check and one release. Not done: a notes-only render (the body HTML of a note's block is still built and
  dropped) - it needs renderer changes.
- Single-value lookups (`ListLabelOf`, `ListLevelOf`, `BlockKindOf`): no host loops over them
  (KitveiHakodesh Core calls none), so no batch API was added.
- Golden dumps byte-identical on .odt, .fodt, the .rtf book and the synthetic .rtf.

### Drawn shapes, groups, VML shapes (.docx, 2026-11-03)

- `DocxDrawingStripper.ShapeTree` copies, per drawing and at render time from the block's raw bytes, only the DrawingML a
  shape is drawn from, without prefixes, inside the `fastDocViewPicture` element: wpg:wgp / grpSp (grpSpPr), wps:wsp (spPr,
  style, bodyPr, an empty txbx marker), pic:pic (blip, spPr), xfrm, prstGeom / avLst, custGeom paths, fills, colours with
  their transforms, a:ln. Everything else (extLst, effects, 3-D) is dropped with its subtree. The w:txbxContent copies stay
  where they were (after the layout, in document order); a shape's n-th txbx marker = the n-th text box.
- `DocxShapeReader` -> `WordShape`s (Html/WordShapeHtml.cs, format-neutral) -> one frame placed by
  `WordPictureHtml.AppendPlacement`, one inline SVG, each shape's text in a box over its outline, group pictures by id at
  their box. One shape whose text grows (spAutoFit) keeps its text in the flow; the SVG stretches (non-scaling strokes).
- One plain rect text box (no wps:style, literal colours) keeps the old text-box rendering (CSS border / fill).
- Geometry (`DocxShapeGeometry`, 2026-11-03): every preset from ECMA-376 annex D presetShapeDefinitions.xml (LibreOffice's
  copy, embedded as a resource, parsed once on first use), the shape's avLst replacing the preset's adjust values by name;
  custGeom with avLst / gdLst guides. Every ST_GeomGuideFormula operator is evaluated with LibreOffice's meaning
  (customshapegeometry.cxx:71-90, :486-606), built-in names first (:180-372), the shape's own a:ext in EMU as w / h. Paths:
  both w and h set = that path's own space, otherwise the shape's EMU (customshapeproperties.cxx:236-244, svx
  EnhancedCustomShape2d SetPathSize :650-700); fill modes and stroke per path; moveTo / close / checkClosed as
  CreateSubPath builds the polygon (:2134-2236); arcTo angles are on the ellipse (lcl_getNormalizedCircleAngleRad
  :2051-2068), swAng clamped to +-360, a negative sweep reversed, drawn as LibreOffice's 30-degree Bezier pieces (basegfx
  createPolygonFromUnitEllipseSegment). Each a:path = one object: fill closes every subpath, even-odd (vcl skia kEvenOdd),
  darken / lighten brightness per GetColorData; stroke separately. A prst name that is no preset = a box.
- Colours (`DocxDrawingColor`): srgbClr / scrgbClr / hslClr / sysClr / schemeClr / prstClr (LibreOffice's preset table, so
  dkGray = A9A9A9, not CSS darkgray), every transform of color.cxx getColor :756-874 and alpha / alphaMod / alphaOff
  (:467-477) -> fill-opacity / stroke-opacity / stop-opacity.
- Fill: none, then the theme fill style fillRef idx names (fillStyleLst 1-999, bgFillStyleLst 1001+, past the end = the last,
  theme.cxx:41-55; phClr = the fillRef colour), then the shape's own over it (assignUsed), a:grpFill = the group's fill.
  Linear gradients span the box grown for the angle (gradienttools.cxx init1DGradientInfo), a:path circle = radial gradient
  of the box diagonal about the fillToRect centre (initEllipticalGradientInfo, fillproperties.cxx:528-551), rotWithShape="0"
  counter-rotated. Line: none, the theme a:ln lnRef idx names (all of it: width, fill, dash, cap, ends), then the shape's a:ln.
- Dashes: lclConvertPresetDash (lineproperties.cxx:60-84) dots / dashes / distance in line widths; a round or square cap is
  inside the length (0.99 moved from dash to gap, :451-462); a:custDash drawn stop by stop (d / sp in 1000ths of a percent or
  "n%"). Line ends: triangle / stealth / diamond / oval / arrow with sm / med / lg width and length, LibreOffice's marker
  outlines and sizes (lineproperties.cxx:223-367, minimum base 0.7 mm), placed and the shaft cut back as drawinglayer does
  (b2dlinegeometry.cxx:38-178, polygonprimitive2d.cxx:704-770), on every open subpath; headEnd = start, tailEnd = end.
- Groups: child at (x - chOff) * ext / chExt from the group's offset, chExt 0 = scale 1 (shape.cxx:1433-1456); rot / flipH /
  flipV per shape (transform2dcontext.cxx:48-49).
- VML (w:pict without v:imagedata): first v:rect / v:roundrect / v:oval / v:line / v:shape; fillcolor (default white),
  strokecolor (black), strokeweight (default 1 EMU = thinnest), filled / stroked, v:fill on, v:stroke dashstyle (as
  LibreOffice maps it to DrawingML presets, vmlformatting.cxx:565-597; a number list = dash / space pairs in line widths,
  W3C NOTE-VML), endcap, start / end arrows (block / classic / diamond / oval / open, :599-633), v:textbox inset (defaults
  0.1in / 0.05in, vmltextboxcontext.cxx:186-211), mso-fit-shape-to-text. v:roundrect arcsize = a fraction of half the shorter
  side, default 0.2 (W3C NOTE-VML; vmlshape.cxx:1066-1077 the same radius). Outlines are the rect / roundRect / ellipse
  presets. Placement as VML pictures.
- Layout (2026-11-04): the preset / custGeom text rectangle (a:rect) holds the shape's text, mirrored with the shape's flips
  (EnhancedCustomShape2d::GetTextRect); a vertical flip turns text 180 degrees (shape.cxx:1501-1515, svdoashp.cxx:1551-1566).
  Groups carry their rotation and flips to every child and its text box as Shape::createAndInsert does (flips toggled,
  rotation turned by the parent's orientation, centres mirrored and turned about the group centre), with MSO's width / height
  swap for children turned 45-135 / 225-315 degrees in a scaled group. wpc canvases: a group of their own size, the wpc:bg /
  wpc:whole background rectangle (WordprocessingCanvasContext.cxx). v:group: every child in the group's coordinate system
  (coordorigin / coordsize, lclGetAbsRect), nested groups, lines, pictures, text boxes in order.
- Fills: a:path rect / shape = LibreOffice's rectangular gradient (four edge regions with linear gradients, clipped); a:pattFill
  = LibreOffice's hatch for the preset (hatchmap.hxx) over the background colour; a:blipFill = an HTML picture layer between SVG
  layers, clipped to the outline, stretched by fillRect or tiled at the picture's original size (WordPictureSize: PNG pHYs,
  JPEG JFIF, BMP, else 96 per inch as graphichelper.cxx getOriginalSize).
- SmartArt (2026-11-04): drawn from its diagramDrawing part (the data part's dsp:dataModelExt relId, resolved against the
  drawing's own part as LibreOffice does) as shapes with their DrawingML text (DocxDrawingTextHtml, txXfrm box, bodyPr
  anchor); without a drawing part, an mc:AlternateContent shows its mc:Fallback picture. Pictures of the drawing part get
  "x:part#rId" ids.
- Charts: DocxChart draws c:chart from the cached values as SVG (bar / column clustered / stacked / percent, line, area,
  pie, doughnut, of-pie as a pie, scatter, radar, bubble, stock and surface as lines; 3-D kinds as 2-D), titles, axes,
  gridlines, legend, data labels; colours from the series / points or LibreOffice's chart style patterns
  (objectformatter.cxx). Layout and automatic axis scaling are our own. chartex charts (cx:chart: waterfall, treemap,
  sunburst, histogram, box and whisker, funnel, region map) show their mc:Fallback picture.
- chartex (2026-11-04, DocxChartEx): histogram (binning: size / count / Scott's rule, underflow / overflow), pareto,
  waterfall (subtotals, connectors), funnel, box and whisker (QUARTILE.EXC / .INC, 1.5 IQR whiskers, outliers, mean),
  treemap (squarified, banner / overlapping parent labels), sunburst, and region maps from a "clear" geography cache
  (Bing point-compressed rings; mercator / miller / robinson / albers). Only a region map cached as "binary" (format not
  specified, [MS-ODRAWXML] 2.24.3.29) shows its mc:Fallback picture. Search reads their document text (titles, legend,
  categories, data labels; not computed bin ranges).
- Pattern fills: the exact 8 x 8 cell of each prstPattern (GDI+ hatch bitmaps, Wine gdiplus brush.c), an SVG pattern of
  0.75 pt pixels, over the background colour; replaces LibreOffice's hatch approximation.
- WordArt (DocxWordArt): DrawingML a:prstTxWarp (wps:bodyPr) mapped to its WordArt shape type and adjust units as LibreOffice  does (presetgeometrynames.cxx, fontworkhelpers.cxx:95-148), VML v:shape types 136-175 with v:textpath; drawn by the shared  WordOfficeArtGeometry.Preset + WordOfficeArtFontwork.Apply in the run's w14:textFill / w:color (DrawingML) or the fill colour  (VML). Search: the DrawingML text box text once; the VML v:textpath string once (byte scanner, reference reader, and notes,
  where the note parser sees the stripped layout's textpathString); none inside mc:Fallback.
- Search: SmartArt shape text and chart text (title, axis titles, legend, category and data labels; not value-axis
  numbers) are fed after their block's text, once, by both the byte scanner and the reference reader (dgm:relIds /
  c:chart outside mc:Fallback); notes likewise.
- Search: shape text is w:txbxContent text, found once (mc:Fallback skipped by both scanner and reference).
- Checks: `--fidelity` DocxShapesFidelityTest and DocxDrawingFidelityTest; live `vue-demo/playwright/docxDrawingDetailsCheck.mjs TestOutput/fidelity-v1-drawing-details.docx` (every path parses, gradients resolve, markers measured at the line ends) (hand-worked values from the cited source:
  dashes, every line end, alpha, arcTo, path without w / h, every guide operator, darken / lighten, theme styles, grpFill,
  VML strokes); live `vue-demo/playwright/docxShapesCheck.mjs TestOutput/fidelity-v1-shapes.docx`
  (SVG boxes measured in the DOM, text once, no console errors).

### RTF text boxes and shapes (2026-11-03)

Sources read online 2026-11-03: "Word 97 RTF for Drawing Objects (Shapes)" (the RTF spec's shapes appendix,
techshelps.github.io/MSDN/SPECS/newrtf/word97rtffordrawingobjectsshapes.htm: \shp / \shpinst / \sp / \sn / \sv, \shptxt "must
come after all the other properties ... inside the \shpinst destination", \shprslt = the Word 6/95 copy, \shpgrp "can be
substituted for a {\shp} at any place", units: EMU, colours 24 bit, booleans 1 / 0, ShapeType table); LibreOffice core master
sw/source/writerfilter/rtftok/rtfsdrimport.cxx (defaults: fFilled on :199, white fill :360-363, black 0.75 pt line :382-384,
insets 91440 / 45720 EMU :113-124, BGR colours :459 / :476, fLine 0 = none :183-190, rotation 16.16 clockwise :245-249, group
child = relLeft.. in groupLeft.. scaled to the group's box :1036-1053, shapeType unset = rectangle :293-300, gtextUNICODE = the
shape's text :442 / :949-957); svx EnhancedCustomShapeGeometry.cxx (MSO outlines in a 21600 space, see `WordOfficeArtShapeTypes`).

- **One parse for all sinks**: `RtfParser.OpenShape` / `CloseShapeGroup` read \shp and \shpgrp in the index pass, the search and
  the renderer alike (no more `ReadShape` byte scan for the renderer only). Inside a shape \par / \sect / \cell / 
ow are
  `TextBoxParagraphEnd` - a block never ends in a shape, so block boundaries, snapshots and block values are unchanged; the
  index counts the shape as the 8-character placeholder it was. Text-box text is not the block's text (no heading text, length,
  list counting, \listtext label). \shprslt is skipped (the shape is read from \shpinst). Unknown words inside a shape (outside
  its text) are shape keywords and not audited.
- **Search**: \shptxt text (and gtextUNICODE) is fed to the matcher once, between separators (shape start / end, each box
  paragraph end), in `RtfSearchTextSink` and the one-pass `RtfSpanSearch` alike; `--rtf-search-golden` stays equal.
- **Renderer**: a lone pib picture keeps the old picture path (ids unchanged). A plain rectangle text box (shapeType unset / 1 /
  202, not rotated or flipped) = `WordPictureHtml.OpenTextBox` with the line as border, the fill, insets, fFitShapeToText = grows.
  Everything else = `WordShapeHtml` (one frame at the root placement, outlines in one SVG, each \shptxt over its shape, group
  pictures at their box); geometry by `Html/WordOfficeArtShapeTypes` (shared with .doc), adjustValue / adjust2Value honoured.
  Placement as RTF pictures. Text-box paragraphs = `text-box-paragraph` block spans.
- Check: `--fidelity` RtfShapesFidelityTest (hand-written .rtf; the .odt shapes fixture saved as .rtf by LibreOffice - its RTF
  export drops draw:g groups); live `vue-demo/playwright/rtfDocShapesCheck.mjs <file> <frames> <texts...>`.

### .doc text boxes and shapes (2026-11-03)

Sources read online 2026-11-03: LibreOffice core master sw/source/filter/ww8/ww8graf.cxx:852-945 (GetTxbxTextSttEndCp: lTxid
story / sequence, fReusable, break table, last mark dropped), ww8scan.cxx:5730-5760 (text-box story base CP), :5969-5976 (ccpTxbx),
:6138-6150 / :6208-6216 (PlcftxbxTxt index 56, PlcftxbxBkd index 75), :1866 (22-byte FTXBXS), ww8struc.hxx:967-975 (WW8_TXBXS),
ww8par.cxx:3594-3622 + ww8par.hxx:1737 (a 0x01 inside a SHAPE field result is skipped), fields.cxx:129 (field 95 = SHAPE);
include/svx/msdffdef.hxx (property ids, MSO_LineDashing, record types), include/filter/msfilter/escherex.hxx:83-99 (FSP flags),
filter/source/msfilter/msdffimp.cxx (line :903-1050 incl. dash patterns, fill :1313-1330, text insets :1474-1477, FitTextToShape
:1598, MSO_CLR :3420-3600, lTxid :5560), dffpropset.cxx (boolean groups :1146-1157, defaults table), svx
EnhancedCustomShapeGeometry.cxx (default filling / stroking tables :6155-6213); [MS-DOC] 2.8.30 PlcfTxbxBkd / 2.9.312 Tbkd (6 bytes).

- `DocBinaryFloatingPictures` now also maps spid -> its SpContainer (a group: its SpgrContainer) on the drawing walk it already
  made; `At` keeps lone pictures (ids unchanged), anything else at a 0x08 is `Drawing(cp)`: `DocBinaryDrawnShapes` reads that
  drawing's records only when its block renders (FSP type / flags, FOPT, ChildAnchor, FSPGR) into the shared `WordShape` model
  (`WordOfficeArtShapeTypes` outlines, dashes, default fill / stroke tables; `WordShapeGroups.Turn` for group rotation / flips,
  shared with RTF). Text: lTxid -> `DocBinaryTextBoxStories` -> a CP range of the text-box story, rendered by the normal
  paragraph renderer as `text-box-paragraph` spans. Plain rectangle text boxes use `WordPictureHtml.OpenTextBox`.
- **Search**: the text-box story was never searched before (the one-pass cursor reads the main text only). Now each block's
  anchored drawings feed their text-box ranges after the block text (`TextRangesIn`), in the one-pass and the block-range
  path alike - each story part once, at its shape's anchor.
- Checks: `--fidelity` DocBinaryShapesFidelityTest (the .odt shapes fixture saved as .doc by LibreOffice: every box text shown
  and found once, no placeholder; POI Bug41898.doc and LibreOffice tdf112618 / tdf76349 / shapes-line-ellipse.doc); live
  `vue-demo/playwright/rtfDocShapesCheck.mjs` on the LibreOffice .doc.
- Not done yet (next): the full MSO preset table (only the outlines in `WordOfficeArtShapeTypes` are drawn; other types are a
  box of their size), custom geometry (pVertices / pSegmentInfo), gradient / pattern / picture fills (fillType), WordArt text
  (gtextUNICODE complex property in .doc), header / footer shapes (PlcSpaHdr), text-box tables drawn as paragraphs, system
  colours (the defaults are used), RTF fillType gradients.

### OfficeArt presets and line ends for RTF and .doc (2026-11-04)

- **Every MSO preset**: `scripts/generate-officeart-shapes.py` copies LibreOffice's svx EnhancedCustomShapeGeometry.cxx tables
  (GetCustomShapeContent: 190 types - vertices, segments, calculation records, default adjust values, coordinate size) into
  the generated `Html/WordOfficeArtShapeDefinitions.cs`; `Html/WordOfficeArtGeometry.cs` evaluates them as LibreOffice does
  (EnhancedCustomShape2d.cxx GetEquation :88-330 and CreateSubPath :2078-2690, svdoashp.cxx lcl_ShapeSegmentFromBinary
  :96-165, tools poly.cxx arcs :246-330) into the shared `WordShapePath` model; adjustValue .. adjust10Value honoured. Types
  without a table (e.g. 0 = not primitive) stay a box. Regenerate:
  `python scripts/generate-officeart-shapes.py <EnhancedCustomShapeGeometry.cxx> <include/svx/msdffdef.hxx> FastDocView/Html/WordOfficeArtShapeDefinitions.cs`.
- **Line ends** (both formats): MSOLINEEND 1-5 ([MS-ODRAW] 2.4.16; 6 / 7 "MUST be ignored"; the RTF shapes page's list
  numbers the open arrow 6 and its lineDashing list skips 5 - both disagree with [MS-ODRAW] / msdffdef.hxx, which are followed),
  width / length 2.4.17 / 2.4.18, drawn with LibreOffice's binary arrow geometry (msdffimp.cxx:786-900 GetLineArrow: 2 pt
  minimum, factors 2 / 3 / 5, open arrow 3.5 / 4.5 / 6, diamond and oval centred) through the shared line-end drawing
  (`WordShape.HeadMarker` / `TailMarker`). .doc: only with fArrowheadsOK (line booleans 0x10, msdffimp.cxx:1063).
- RTF fill default now follows the OfficeArt table too (a line is not filled: LibreOffice makes it a LineShape).
- Check: `--fidelity` WordOfficeArtGeometryTest (every table finite) and RtfShapesFidelityTest outlines (ellipse, triangle,
  round rectangle, diamond, smiley, arrow, line + arrow end).
- **Colours** (both formats, 2026-11-04): `Html/WordOfficeArtColor` resolves MSO_CLR as msdffimp.cxx:3420-3640 - scheme colours
  (no palette in Word) = the property default, system colours = LibreOffice's standard StyleSettings values
  (vcl settings.cxx:497-559) or another property (0xF0..0xFF), then gray / darken / lighten / add / subtract / black-white /
  inverts; a system colour that names another system colour keeps Word's manager default white (ww8par.cxx:500-501).
- **Fills** (both formats): `Html/WordOfficeArtFill` follows msdffimp.cxx ApplyFillAttributes / ImportGradientColor /
  ApplyRectangularGradientAsBitmap - solid; shade 4 / 7 / 8 linear or axial (focus 40..60) with LibreOffice's colour swaps
  (angle >= 0, focus 0, negative focus, axial) and the box grown to the turned extent (basegfx gradienttools.cxx
  init1DGradientInfo); shade shape 6 rectangular (aspect-corrected, getRectangularGradientAlpha) and shade centre 5 as the
  bitmap LibreOffice renders (rectangles shrinking toward fillToRight / fillToBottom, fillShadeColors) - both drawn in 64
  steps; fill / back opacity as stop opacity; pattern 1 = the fill blip's DIB pixels (black = back colour); texture 2 /
  picture 3 = the fill blip by id through the shared picture-fill layer (`WordShapeHtml` FillPictureId / FillPicture; tiled at fillWidth / fillHeight for a
  texture). .doc fill pictures have ids "sf" + BStore index; RTF fill picts count among the block's pictures.
- **Own geometry**: .doc pVertices / pSegmentInfo / pFormulas (complex FOPT data read as dffpropset.cxx does, incl. the
  6-byte array header fix) and RTF pVerticies / pSegmentInfo / geoLeft..geoBottom, evaluated by `WordOfficeArtGeometry.Custom`
  (binary segment encoding msdffimp.cxx:2267-2390, binary angle-ellipse rule).
- **WordArt**: .doc gtextUNICODE / gtextFont / gtextSize (complex strings) and RTF gtext* are the shape's text, centred in
  its font and size, searched once; with the text-effect bit (.doc gtext booleans 0x4000, RTF fGtext) the outline is not drawn
  and the text takes the fill colour (LibreOffice renders Fontwork instead of the outline). Fontwork's path warping (arches,
  waves) is not drawn: the text sits straight in the box.
- Checks: RtfShapesFidelityTest blocks 8-9; DocBinaryShapesFidelityTest (the RTF fixture and a Fontwork .fodt saved as .doc
  by LibreOffice).
- **Tables inside text boxes** (both formats): .doc text-box story paragraphs in a table are grouped up to their row-end mark
  and drawn by `RenderRow` as in the body; RTF \cell / \row inside {\shptxt} are `TextBoxCellEnd` / `TextBoxRowEnd` (never a
  block end) and the renderer draws the box's own table rows; the box's \trowd does not replace the outer row definition
  (saved and restored around the text box).
- **.doc header / footer shapes**: `DocBinaryWindowedDocument.HeaderDrawings` reads PlcSpaHdr (FibRgFcLcb97 41) with the header
  text-box story (PlcfHdrtxbxTxt 58, PlcfHdrtxbxBkd 76, after the main text-box story by ccpTxbx; ccpHdrTxbx at 0x68), anchors
  counted from the header story's start (ww8par.cxx:3656 / ww8graf.cxx:2605-2620). Ready for the paged view: `AnchorCps`,
  `Drawing(cp)`, `At(cp)`; the flow view shows no headers, so they never reach the body or its search
  (DocBinaryShapesFidelityTest: a LibreOffice-saved header with a text box and an ellipse).

### RTF audit at 0 (2026-11-04)

- `--rtf --corpus` (test-data's 715 RTF files) now reports 0 unhandled control words. Words LibreOffice's tokenizer does not
  know (damaged input, other tools' words, wrongly cased ones - control words are case sensitive) are ignored as LibreOffice
  ignores them (rtftokenizer.cxx:2169-2175 "unknown keyword"); the audit tells them apart with
  `FastDocView.Tests/LibreOfficeRtfKeywords.cs` (generated by `scripts/generate-libreoffice-rtf-keywords.py` from rtftokenizer.cxx)
  and prints them separately. Real words without a visible effect here are listed in `RtfFormattingAudit.IgnoredByDesign` with
  the LibreOffice source that shows why (no dispatch case, paged-layout only, vertical text, Word 6 drawing objects).
- `\lbrN` is now rendered: the next `\line` clears left / right / both floats (rtfdispatchvalue.cxx:1779-1800 LBR -> w:br clear),
  kept in two state bits, reset by `\plain`.

### Text effects on their outlines (WordArt / Fontwork, RTF and .doc, 2026-11-04)

- `Html/WordOfficeArtFontwork.cs` lays text effects out as LibreOffice's EnhancedCustomShapeFontWork.cxx does: the outlines are
  the shape's evaluated sub-paths; odd count = single-line mode (one area per outline), even = pairs (upper, lower); lines are
  shared out over the areas in order; a line's height is the shape's height / the most lines in an area; its width is fitted
  to the outline length (SVG textLength, lengthAdjust spacingAndGlyphs). Single-line: SVG textPath turns every glyph along the
  outline with its centre on it (several lines: the outline moved along its normals). Pairs: the text runs halfway between the
  two outlines (fraction-by-length), sized to their mean gap.
- Known differences from LibreOffice, both because an HTML writer has no glyph outlines or advance widths: in pair mode the
  glyph shapes are not bent between the two curves (they follow the curve), and in single-line mode the line height is not
  multiplied by LibreOffice's fit factor (path length / text width at that height).
- The text-effect shape draws only its text, filled with its fill (gradients too); the outline is not drawn. Search unchanged
  (.doc gtextUNICODE and RTF gtext text are still fed once).
- `.docx`: VML WordArt (v:textpath on the same MSO shape types) and DrawingML prstTxWarp are not read by the .docx reader yet; the
  shared `WordOfficeArtGeometry.Preset` + `WordOfficeArtFontwork.Apply` take an MSO type, its adjust values and the string.
- **2026-11-04, the page finishes Fontwork with the browser's font metrics** (KitveiHakodesh vue-frontend
  `features/fast-doc-view/fastDocViewFontwork.ts`, directive on the block HTML): the single-outline fit factor is applied with
  measured text widths (SVG getComputedTextLength), as LibreOffice's CalculateHorizontalScalingFactor / GetFontWorkOutline do,
  then the line is placed per gtextAlign (msdffimp.cxx:4466-4478); paired outlines are bent as FitTextOutlinesToShapeOutlines
  does - the mapping is linear in fy for each fx, so the text is drawn once into a canvas and copied column by column (one
  device pixel wide) between the two curves; the bent image replaces the paint, the text stays (transparent) for selection.
  The writer marks each line with data-fontwork-* attributes (`WordShapeHtml.AppendTextOnPaths`); without the page's script the
  midline layout above remains. Live check: `vue-demo/playwright/fontworkCheck.mjs`.
- **2026-11-04, bent text keeps its paint and stays sharp**: the page bends a white coverage mask and paints a rectangle with
  the text's own fill through it (a gradient / pattern paint is kept, as Fontwork fills its outlines with the shape's fill);
  it bends again when the page zoom or the device pixel ratio changes. Fixture: `fidelity-v1-fontwork-curved.fodt` (inflate
  with a gradient, wave) saved as .doc by LibreOffice (DocBinaryShapesFidelityTest); live:
  `vue-demo/playwright/fontworkCheck.mjs <rtf fixture> <curved .doc>` (curved bending, gradient paint, pixel ratio 1 -> 2, Ctrl + =).
- **2026-11-04, the fill spans the text's own bounds**: LibreOffice puts all glyph outlines of a text effect into one path
  object with the shape's fill (EnhancedCustomShapeFontWork.cxx:1096-1130 CreateSdrObjectFromParagraphOutlines), so its
  gradient spans the text, not the shape's box. The page copies the paint over the union of the bent ink and the single-line
  text boxes (linear gradients laid out again at their angle, rectangular shades mapped, patterns keep their tile); live check
  on the inflate fixture (its gradient runs from the ink's top to its bottom).

## Word fields (all formats) (branch work/word-fields, 2026-11-04)

One engine in `Fields/` for .docx, .doc and .rtf; each format only finds the field marks:

| Format | Begin / separator / end | Flags | Open fields at a block start |
|---|---|---|---|
| .docx | w:fldChar begin / separate / end (any run, any paragraph), w:fldSimple w:instr | w:dirty, w:fldLock; w:ffData form state | open-field stacks + a `C` counters entry per block (index format 14) |
| .doc | 0x13 / 0x14 / 0x15 in the text | PlcFld per story (FibRgFcLcb97 16 / 17 / 18 / 48 / 57 / 59, ww8scan.cxx WW8Fib read order): fResultDirty 0x04, fLocked 0x10, fPrivateResult 0x20 (POI FLDAbstractType.java:34-41; LibreOffice ww8par5.cxx:1628); FFData at the code's 0x01 data location + 68 (PICF, ww8graf2.cxx:442-457; ww8par3.cxx:2127-2240) | rebuilt per block from PlcFld (`DocBinaryFields.OpenAt`: parent chain, code re-read); counters computed once on first need |
| .rtf | {\field ... {\*\fldinst} {\fldrslt}} | \fldlock, \fldpriv (\flddirty is a formatting mark in RTF 1.5 "Fields", not a recalculation); {\*\formfield} \fftype \ffres \ffdefres {\*\ffl} {\*\ffdeftext} (RTF 1.5 "Form Fields"; LibreOffice rtfdispatchvalue.cxx:1426-1455) | anchor heads carry the open fields + counters, span rescans track them per block (index format 7) |

- `WordFieldStack` decides what shows: codes never; a nested field's result inside a code is part of that code
  ({ IF { = 2+2 } = 4 ...}); a cached result once; a page-number field (PAGE, PAGEREF, NUMPAGES, SECTION, SECTIONPAGES)
  is a placeholder `<span class="field-page-placeholder">#</span>` (PAGEREF \h still links); XE / TC / TA / RD / SET /
  ASK / PRIVATE / NEXT / NEXTIF / SKIPIF / PRINT show nothing; a private result (.doc fPrivateResult, RTF \fldpriv - Word
  marks form fields so) never shows. Computed: a field with no result, a dirty one that is not locked (.docx w:dirty, .doc
  fResultDirty) and the fields Word draws from their code (EQ, MACROBUTTON, GOTOBUTTON, LISTNUM, AUTONUM*, ADVANCE,
  FORMCHECKBOX, FORMDROPDOWN, SYMBOL - their result text is empty in Word 16). Locked = never recomputed.
- Search reads exactly that: a cached result once, computed text, never codes, placeholders or hidden fields. Every
  reader runs the same stack: .docx byte scanner + reference XmlReader (notes with fields go to the parsed path), .doc one
  pass + per-block ranges, .rtf span search + per-block reparse.
- Computed (Word 16 on the dev machine is the reference: COM Fields.Add + Update; values in WordFieldsFidelityTest):
  `=` formulas (17.16.3; -2^2 = 4; the default result rounds to max(2, decimals written); table cells A1 / A1:B2 / RnCn /
  LEFT / RIGHT / ABOVE / BELOW; bookmarks; Word's error texts), IF, COMPARE (wildcards), QUOTE (an unquoted number is a
  character code), SYMBOL (symbol fonts through U+F0xx), MERGEFIELD «name», DATE / TIME (the system culture's short
  patterns), CREATEDATE / SAVEDATE / PRINTDATE and AUTHOR ... NUMWORDS / DOCPROPERTY / DOCVARIABLE / INFO / FILENAME /
  FILESIZE from the document's properties (docProps core / app / custom + settings docVars; OLE SummaryInformation +
  DocumentSummaryInformation, PIDs from POI PropertyIDMap.java; RTF {\info} \userprops \docvar), SEQ (\r \c \h \s, counted
  at index time), AUTONUM* and LISTNUM (one counter / one level set shared by all list names, as Word 16 counts), REF
  (bookmark text), form fields (checkbox ☒ / ☐, drop-down entry, text default), MACROBUTTON / GOTOBUTTON text
  (GOTOBUTTON links), EQ as MathML Core (\a \b \d \f \i \l \o \r \s \x), ADVANCE \l / \r as a gap.
- Switches: \* Arabic ArabicDash alphabetic / ALPHABETIC roman / ROMAN Ordinal OrdText CardText DollarText Hex hebrew1
  (up to 392) hebrew2 (one ת per 22 - Word's field rule, not the list rule) Upper Lower FirstCap Caps; a \* number format
  wins over \#; \# pictures (0 # x . , - + 'text', ; sections, half away from zero); \@ date pictures.
- Computed from the index too (2026-11-04): REF \n \r \w as Word 16 writes them (COM: a list "%1)" "%2)" "%3)" at 2) b) i)
  → \n "i)", \w "2)b)i)", \r "i)" both from inside 2)b) and from outside the list; a trailing period dropped, "1." → "1").
  .docx builds the full context in the list counter (each level's own lvlText with the counters; an upper level already in a
  deeper level's text, %1 in "%1.%2.", is not repeated); .rtf walks back over the blocks' list levels to the nearest
  paragraph of each upper level; .doc builds it like .docx in its list counter during the span scan (each level's
  xst with the counters; a level a deeper xst holds as a placeholder, rgbxchNums, is not repeated) - no index change, the
  span anchors already carry the counter state (fields-lists-v1.doc, Word 16 via COM, same values). \p writes
  "above" / "below", or Word 16's Hebrew "מעל" / "מתחת" when the document is right-to-left (Word follows the field's text
  language; the readers know the document's direction, not each run's language, before any text). NOTEREF (\f \p);
  STYLEREF (back from the field's block, then forward - flow view has no page for Word's page-relative search or \l; Word's
  Hebrew error when no paragraph has the style; .docx by the index's per-block style ids and w:name, .doc by each
  paragraph's istd and STD name, .rtf by each block's first-paragraph \sN from the span scans and its \stylesheet name);
  TOC (\o / \n levels, built from the index's headings inside the field's block, page numbers as placeholders, \h entries
  link "#block-<index>" - the app's FastDocViewPage jumps there like a bookmark); INDEX (XE "main:sub", \t, \e, \f; entries
  collected at index time: .docx format 15, .rtf format 8, .doc read once from PlcFld) and TOA (TA \l by \c category);
  BARCODE as POSTNET bars in SVG (Wikipedia "POSTNET" from USPS Publication 25) and \f A-E as Facing Identification Marks
  (A 110010011, B 101101101, C 110101011, D 111010111, E 101000101 - Wikipedia "Facing Identification Mark" from the USPS
  Domestic Mail Manual 708.9 / 202); ADVANCE \u / \d raise / lower the paragraph's later runs (position: relative) - \x / \y
  are page positions (paged view TODO).
- RTF 
 \w walk back at most 4096 blocks and stop at another list's paragraph (\ls per block from the span scans);
  STYLEREF looks at most 4096 blocks each way - every block read is a span rescan (no per-style table in the index).
- RTF field lookups read through a second block table whose rescans compute no fields: a lookup runs while the document's
  table is mid-rescan and must not re-enter it.
- .docx lookups cannot re-enter a scan (checked 2026-11-04): a span scan only feeds the scanner and the list labeller (no
  fields); a render reads its blocks' XML whole before any field computes, and search feeds its own buffer - a lookup's
  rescan touches only the scan chunk and the ring. Harness: REF \w to far spans, more spans than the cache holds, random
  window order and search (DocxListLookupsAcrossSpans).
- Bookmark lookups during a render never scan on: while the first open is still indexing, a bookmark not found yet is
  unknown (the field shows nothing until the index completes), not missing. .doc SEQ \s heading positions are read once.
- Still not computed: INCLUDETEXT / INCLUDEPICTURE (never fetched), USERNAME / USERINITIALS / USERADDRESS without an
  argument (Word's user settings, not the document's). DATE \h / \s: Word 16 on a Hebrew system wrote the same Gregorian
  date with \h. Number words are English (Word 16 wrote English CardText for Hebrew text too); DollarText's "and" follows
  Word's UI language (the Hebrew UI wrote "ו-"; we write "and"); error texts are Word 16's Hebrew ones.
- RTF REF without a result reads from {\*\bkmkstart} to {\*\bkmkend} (on into later blocks when the bookmark spans them).
- **Paged view TODO**: PAGE / NUMPAGES / SECTION / SECTIONPAGES / PAGEREF placeholders get their numbers when the paged
  view lays out pages.
- Harness: `--fields` (seconds; also in `--fidelity`): WordFieldsFidelityTest (switches, formulas, counters, .docx render
  + both search readers), WordFieldsOtherFormatsTest (test-data/word-fields/fields-v1.doc / .rtf saved by Word 16 with
  make-fields-fixtures.ps1, and hand-written RTF: no \fldrslt, \fldlock, \fldpriv, \formfield, counters across index
  anchors).

## Two-sided and contour wrap (all formats)

User decision 2026-11-04: text on both sides of an object (.docx wrapText bothSides / largest, VML side both / largest,
.doc FSPA wrk 0 / 3, RTF \shpwrk 0 / 3, ODF style:wrap parallel / dynamic / biggest) has NO top-and-bottom fallback.
CSS cannot put text on both sides of a float (CSS Shapes 1 s.2), so the object floats flush to its side (left / right as
placed; centred = left, our choice for the tie), its horizontal offset dropped, and the text wraps fully along the other
side. A one-sided wrap (left / right only) floats opposite its text and keeps its offset. `WordPictureLayout.TextSides` +
`WordPictureHtml.AppendPlacement` / `FloatsRight`. .doc reads the side from posh (0x038F in the 0xF122 record,
msdffimp.cxx:5199-5216; inside / outside as on an odd page). An object placed by its offset alone (no alignment) floats to the nearer edge of its section's text width (page width
minus margins and gutter, already in every index - `WordDocumentSection`, `WordDocumentPageSetup.TextWidthPointsAt`; the
renderer gets it per block as `ColumnWidthPt`): right when its centre lies right of the column's centre. A centred object's
centre is the column's own - the tie, left. Multi-column sections are not split: the flow view shows one column at the
full text width, so that is the width the reader sees. Offsets from other origins are converted to the column first (`WordPictureOffsetOrigin`, `WordPictureHtml.ColumnOffset`):
from the page edge (.docx relativeFrom page / leftMargin / insideMargin, VML page / left- / inner-margin-area, ODF
horizontal-rel page / page-start-margin, RTF \shpbxpage / posrelh 1, .doc bx 1) less the column start = left margin +
gutter; from the right margin area (rightMargin / outsideMargin, right- / outer-margin-area, page-end-margin) plus the text
width. Inside / outside are taken as an odd page. A character / frame origin is not converted (left). The column start is
stored per section next to the text width (index formats: .docx 16, RTF 9, ODT 11; .doc reads its SEP at open).
Not done: a right-to-left section's gutter on the right (w:rtlGutter / tlgutter) - the gutter is always counted on the left.

Contour wrap = CSS `shape-outside:polygon(...)` from `WordPictureContour.ShapeOutsideCss` (the outline's extreme on the
text side per 1 pt row, at most 2000 rows; gap on the text side as shape-margin; through drawn as tight - a float keeps
its text on one side). The contour:
- .docx wp:wrapTight / wp:wrapThrough wp:wrapPolygon (ECMA-376 Part 1 20.4.2.16-18), VML wrapcoords (tight / through);
  Word's 21600 x 21600 space over the object ([MS-OI29500] as quoted by LibreOffice WrapPolygonHandler.cxx:120-129).
- .doc FSPA wr 4 / 5 (ww8graf.cxx:2721-2725 bContour): pWrapPolygonVertices 899 (msdffimp.cxx:5574-5605, 8-byte int32 or
  4-byte int16 pairs), else the drawn shapes' own outline (`WordShapeHtml.Outlines`); a lone picture without one: its box.
- RTF \shpwr 4 / 5 with pWrapPolygonVertices "8;n;(x,y);..." (rtfsdrimport.cxx:856-885), else the shape's outline.
- "Through" is a contour wrap in Word, like tight plus the polygon's holes: ECMA-376 wp:wrapThrough (Part 1 20.4.2.17),
  MS-DOC FSPA wr 5 "wrap tightly, but allow holes" (as quoted by Apache POI fspa_type.xml from the Word 97-2007 binary
  spec; LibreOffice ww8graf.cxx:2721-2725 the same), RTF \shpwr 5 "Wrap through" (Word 97 RTF for Drawing Objects). All
  three are drawn as contour wrap. LibreOffice differs for RTF only: it reads \shpwr 5 as THROUGH = text over the object
  (rtfdispatchvalue.cxx:1255-1257). Holes cannot be laid out in CSS (one side per float), so through = tight here.
- ODF style:wrap-contour: draw:contour-polygon / contour-path, else a shape's own outlines. The page re-sweeps the
  contour when a shape grows with its text (vue-frontend fastDocViewShapeText.ts `contourPolygon`, same format).
LibreOffice's extra 15-twip shift of Word polygons (WrapPolygonHandler.cxx:99-114) is its own guess and not applied.
Checks: `--fidelity` WrapSidesAndContourFidelityTest (synthetic .rtf, picture-wrap-polygon.docx / .rtf, tdf128501.doc),
PagePlacementFidelityTest (centred wrapSquare floats left), OdtShapesFidelityTest / OdtDrawingFidelityTest (ODF).

## Error handling (2026-11-04)

- Open failures are classified by `WordDocumentOpenFailures.Classify` (encrypted, corrupt, online-only, not-found,
  unsupported format, access-denied); encrypted files throw `WordDocumentEncryptedException`, cloud placeholders
  `WordDocumentOnlineOnlyException`. Hosts map the class to their own translated message.
- `WordDocumentBlockRenderGuard`: a window whose render throws a content error is rendered again block by block;
  blocks that still fail become the "unsupported" placeholder, reported once per document through
  `BlockRenderFailed`. IO errors (file changed / gone), cancellation and OOM are never swallowed.
  Harness: `CorpusRegressionTest.RunOpenFailureClasses` / `RunMalformedBlock`; live: `playwright/errorsCheck.mjs`.
### Production fixes (2026-11-04)

- **.docx w:moveFrom / w:ruby: search = what is rendered.** w:moveFrom (a tracked move's source) is hidden by the renderer
  (Word's final view) and now skipped by every search path too (byte scanner, reference reader, note scanners). w:ruby used to
  render nothing; it now renders as HTML `<ruby>base<rt>guide</rt></ruby>`; search feeds the base only (w:rt skipped, as
  text:ruby-text already is in .odt). The note byte scanner leaves notes with w:moveFrom / w:rt to the parsed path.
  Check: `DocxFidelityTest.MoveFromAndRubySearch` (body + note).
- **ODT office:binary-data without '=' padding** (LibreOffice's tdf145542 .fodt files) gave a picture id whose read failed
  (404, broken image in the page): the head check decoded 16 characters, the read decoded the whole and threw. Both now decode
  as LibreOffice does (comphelper/source/misc/base64.cxx `Base64::decodeSomeChars`, read online 2026-11-04: non-alphabet
  characters skipped, complete 4-character groups decoded, an unfinished last group dropped).
  Check: `OdtDrawingFidelityTest.UnpaddedBinaryPicture`.
- **Notes pane for .rtf / .odt / .fodt**: RenderNotes = the blocks' notes verified for all three (`NotesWindowTest` now
  includes a synthetic .fodt) and live in the notes pane. Still not done: a notes-only render (the body HTML of a note's block is
  built and dropped) - a speed item, not a correctness one.
