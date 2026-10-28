# WordDocumentRandomAccess — design notes

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

## The .docx index (Docx/)

`DocxRandomAccessIndexBuilder` does the one pass. While inflating it feeds every chunk to
`DocxBodyBlockScanner`, a byte-level XML state machine (no DOM) that records, for every child of
`<w:body>` (**block** = paragraph / table / content control / other):

- uncompressed start offset (end = next block's start; last = `</w:body>` offset)
- kind, `w:pStyle` id (for headings without reading the block), text length in bytes
  (for the viewer's scroll-height estimate)

plus the root element's opening tag, whose namespace declarations wrap each block so it can be
parsed in isolation (`DocxBlockXmlReader`).

The index is saved to `%LOCALAPPDATA%\KitveiHakodesh\WordDocumentRandomAccess\<name>.<hash>.wordindex`
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
`w:lvlRestart`, legal numbering, `.doc` lists (LFO/LST tables).

Deliberately not done yet:

- **Images** — `[image]` placeholder. Needs the rels part + a blob endpoint.
- **Fields** — `w:fldSimple`/`instrText` show their result text only.
- **Word 6/95, encrypted, RTF-as-.doc, MHTML-as-.doc** — rejected with a clear message.

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
10. `.doc`: fonts (Sttbfffn), lists (PlcfLst/LFO → the same label counter), footnotes (PlcffndRef/Txt), tables, the remaining common sprms.

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

`WordDocumentRandomAccess.Tests -- --audit <files>` streams every text-bearing part (.docx) or
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

## Integration plan (the app)

The block API is shaped like `seforimApi.getLinesPaged` (`index` + `html`) on purpose: a
Word-document adapter can slot into the lines view's windowed scroller, with headings feeding
the TOC. The WinForms/WebView2 viewer here is only a stand-in; its `viewer.js` shows the
protocol (`requestBlocks` → `blocks`) and a minimal estimate-then-measure windowing.

Open questions carried over from `docx-as-lines` (2026-08-27): synthetic book ids vs per-doc
SQLite; footnotes as commentary links.

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
4. Numbering `suff`, `lvlJc`, `lvlRestart`, `isLgl`, `numStyleLink`, picture bullets — index-time.
5. ~~Theme colors (`themeColor` / tint / shade, `clrSchemeMapping`)~~ — done 2026-10-28.
6. Hyperlinks: docx rels + HYPERLINK fields, .doc PlcfFld; open-field state per block at index time.
7. .doc fast-saved files: per-piece Prm / Prc sprms — render-time.
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
2. **ODT/OTT** — content.xml checkpointed like document.xml; automatic styles + font faces come
   before office:body and are captured during the index pass into the sidecar; lists flattened to
   blocks with labels per XMLTextListBlockContext rules (xml:id, continue-list, continue-numbering,
   start-value); notes are inline `text:note` — their body byte ranges recorded at index time and
   rendered separately; property map from txtprmap (per-script font/size variants, text-position,
   writing-mode for RTL). Split huge tables into row blocks. (M–L)
3. **FODT** — same scanner over the plain file: offsets only, no checkpoints — the fastest format. (S after ODT)
4. **Flat OPC** — docx parts as byte ranges inside one XML; feeds the docx pipeline via a part-source abstraction. (S–M)
5. **Word 2003 XML** — close to docx element names; inline footnotes. (M)
6. **RTF** — byte-offset blocks at `\par` / `\row`, with a formatting-state snapshot per block (state carries across paragraphs). (L)
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

