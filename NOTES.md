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
`DocxBodyBlockScanner`, a byte-level XML state machine (no DOM) that records, for every child of
`<w:body>` (**block** = paragraph / table / content control / other):

- uncompressed start offset (end = next block's start; last = `</w:body>` offset)
- kind, `w:pStyle` id (for headings without reading the block), text length in bytes
  (for the viewer's scroll-height estimate)

plus the root element's opening tag, whose namespace declarations wrap each block so it can be
parsed in isolation (`DocxBlockXmlReader`).

The index is saved to `%LOCALAPPDATA%\KitveiHakodesh\FastDocView\<name>.<hash>.wordindex`
keyed by path, validated by file length + last write time. Re-opening a 50 000-block file from
the cached index: ~50 ms. Cold build runs at about 40 MB/s of XML with a few MB of RAM.

Measured (synthetic, 50 000 blocks, 22 MB XML): build 0.5 s, index 1.2 MB, random window of 40
blocks worst 30 ms, consecutive windows ~2 ms each.

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

- **Images** — `[image]` placeholder. Needs the rels part + a blob endpoint.
- **Fields** — `w:fldSimple`/`instrText` show their result text only.
- **Word 6/95, encrypted, MHTML-as-.doc** — rejected with a clear message (RTF, also when saved as .doc, opens as RTF).

## Coverage plan — "full flow-preview coverage" (estimate 2–3 weeks, not months)

Parsing is the cheap part; each item below is mostly the semantics → CSS table plus a test file.
Ordered by value for text-heavy Hebrew documents:

1. Footnotes/endnotes (deferred by user) — bodies from footnotes.xml, shown like commentary.
2. ~~Themes — `theme1.xml` fonts and colors~~ — done (.docx).
3. Tables, properly — `tblGrid`/`tcW` widths, borders (`tblBorders`/`tcBorders`), row heights, `tblInd`, nested tables, `.doc` rows reassembled from cell marks into `<table>`.
4. Images — `drawing`/`pict` → rels → blob endpoint (inline only; floating as inline).
5. Text boxes / shapes — `mc:AlternateContent` → `w:txbxContent` rendered inline after the paragraph.
6. Paragraph borders/shading, tab stops (`w:tabs` positions), character spacing/scale, caps/small caps, underline styles, `w:lvlRestart`, legal numbering.
7. Fields — `fldSimple`/`fldChar` result text (done) + `HYPERLINK`, `REF`, `PAGEREF` targets as anchors; `TOC` left as text.
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

**MHTML** (.mht/.mhtml, and .doc files beginning with "MIME"): `WebArchiveContainer` reads headers only
(RFC 2045 §5.1 parameters, RFC 2046 §5.1.1 boundaries) and takes the first text/html part; its body is
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
  and table borders, text boxes / shape text (`\shp` skipped whole), hyperlink targets, style
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
