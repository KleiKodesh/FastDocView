# FastDocView: unsupported and partial features

Audited 2026-11-10 against the code. Companion of [SUPPORTED-FEATURES.md](SUPPORTED-FEATURES.md) (same areas, same
format names). Status: **planned** (on the roadmap in NOTES.md / CLAUDE.md), **deferred** (wanted, not scheduled),
**won't do** (impossible in a browser, or out of scope on purpose).

How to verify: `FastDocView.Tests -- --audit <files>` lists what a file uses that the renderers ignore (the
ignored-by-design set is printed with reasons); `--corpus ../test-data` checks the error paths.

## Character formatting

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Underline words only (skip spaces) | docx, doc, rtf, odt | No CSS equivalent; drawn as a plain underline (LibreOffice does the same) | won't do |
| w14 shadow scale / skew (sx, sy, kx, ky), perspective shadows | docx | CSS text-shadow can only offset and blur; a transformed copy needs a second text layer per line, which breaks wrapping and windowing | won't do |
| w14 reflection, 3-D text (props3d / scene3d), ligatures / number forms / stylistic sets | docx | Not read | deferred |
| Kerning | docx, doc, rtf, odt | Ignored by design (audit set); the browser's own kerning applies | won't do |
| East-Asian font slot (eastAsiaTheme, east-asian fonts), eastAsianLayout, fitText | docx, doc | Model has no East-Asian slot; not needed for Hebrew books | deferred |
| Class-based formatting in web archives | mhtml | `<style>` sheets are dropped for safety; only inline styles stay | deferred |
| ODT run script split per text node | odt/fodt | ODT has no per-run script flag; an LTR run broken by a span is judged in pieces | deferred |
| Bullet whose symbol font is set only for complex script | doc | Shows no glyph (LibreOffice also reads the Latin font) | won't do |
| Blinking / animated text | all | Not drawn | won't do |

## Paragraphs

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Tab leaders and right / centre / decimal tab alignment | docx, doc, rtf, odt | Flow HTML has only `tab-size`; true tab layout needs a line layout engine | deferred |
| Keep with next, keep lines, widow / orphan control, page break before | docx, doc, rtf, odt | Pagination: belongs to the paged view | planned (paged view) |
| Frames (framePr) other than drop caps: absolute position, wrap | docx, doc, rtf | Frames flow as paragraphs; positioning is page-bound | planned (paged view) |
| RTF border spacing (\brsp), border shadow, border groups | rtf | Each paragraph is boxed alone | deferred |
| ODT drop cap split across spans | odt/fodt | Only the first text node is wrapped | deferred |
| Hebrew hyphenation | all | Chromium has no Hebrew hyphenation dictionary | won't do |
| Line numbering, snap to grid, document grid | docx, doc, rtf, odt | Page-bound / East-Asian layout | planned (paged view) |

## Lists

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Restart-limit level (\levelnorestart is read; the limit is not) | rtf | Not read | deferred |
| A level replaced inside a list override | rtf | Only start-at is read | deferred |
| Style-linked numbering | rtf | Not read | deferred |
| Separate bullet colour / size | rtf (counted labels), odt | Labels take the general label formatting | deferred |
| ApplyParagraphMarkFormatToNumbering setting | docx | Not read (LibreOffice writes it only for Word imports) | won't do |
| LibreOffice's CJK counting, Greek and Bulgarian num-format names | odt | LibreOffice's own algorithms differ from Word's; shown as decimal | deferred |

## Tables

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Floating table position (tblpPr) and text wrap around tables | docx, doc, rtf | Ignored by design (frame / table positioning) | planned (paged view) |
| Vertical cell text direction | doc, rtf, odt | Not drawn (docx w:textDirection is) | deferred |
| Cells merged down drawn as one cell | rtf, doc | An outer row is one block (windowed): a cell merged down is cells without the line between them, not one rowspan cell (a nested table gets a real rowspan) | won't do (windowed rule) |
| Repeat header row, row "can't split" | all | Page-bound | planned (paged view) |
| Row heights in .mhtml | mhtml | The page's own HTML heights are kept as written; no Word row-height rule is applied | won't do |
| Exact row heights as Word measures | all | Row heights are drawn (at least / exact), but text heights are the browser's, not Word's metrics | won't do |

## Sections and pages

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Headers and footers (and their pictures / shapes) | docx, doc, rtf, odt | Flow view has no pages; .doc header shapes are already read for the paged view | planned (paged view) |
| Page layout, page breaks drawn as pages, page size on screen | all | Flow view by design | planned (paged view, not built yet) |
| Multiple columns | docx, doc, rtf, odt | The whole text width is one column; multicol pager belongs to the paged view | planned (paged view) |
| Page-number fields (PAGE, PAGEREF, NUMPAGES, SECTION, SECTIONPAGES) | docx, doc, rtf | `#` placeholder until there are pages | planned (paged view) |
| Mirrored margins, right-to-left gutter on the right | docx, doc, rtf, odt | Width only; no page sides | planned (paged view) |
| Page borders, page background, watermarks | all | Page-bound | planned (paged view) |

## Notes

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Notes at the page bottom / end of section | all | Draft-view model on purpose: notes pane and tooltip | won't do (flow view) |
| Notes-only render (a note's block body is built and dropped) | docx, odt | Speed item, not correctness | deferred |
| Restart note numbering per page | docx, doc, rtf | Page-bound | planned (paged view) |
| Note start value and per-section restart | docx (numStart / numRestart), doc (Dop) | Not read | deferred |
| Notes inside text boxes / drawing text | odt/fodt | The drawing-text renderer skips text:note (OdtBlockHtmlRenderer.Drawing) | deferred |
| Note number format | doc | Always decimal; nfcFtnRef / nfcEdnRef not read | deferred |

## Fields

| Feature | Formats | Reason | Status |
|---|---|---|---|
| INCLUDETEXT, INCLUDEPICTURE | docx, doc, rtf | Never fetched: a document must not read other files | won't do |
| USERNAME / USERINITIALS / USERADDRESS without an argument | docx, doc, rtf | Word's user settings, not the document's | won't do |
| STYLEREF page-relative search and \l | docx, doc, rtf | Flow view has no page | planned (paged view) |
| Lookups during the first open (REF to a bookmark not indexed yet, TOC, INDEX) | docx, doc, rtf | Show nothing until the index completes, instead of Word's "not found" | won't do (by design) |
| Field computation | odt/fodt | Saved field text only | deferred |
| Fields | txt, mhtml | Format has none | n.a. |

## Links, bookmarks, TOC

| Feature | Formats | Reason | Status |
|---|---|---|---|
| TOC page numbers | docx, doc, rtf | Placeholders until pages exist | planned (paged view) |

## Pictures

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Picture cropping (a:srcRect, \piccrop, fo:clip) | docx, rtf, odt (doc: size only) | Crop is not applied: the whole picture is drawn in the frame | deferred |
| Through wrap (text over holes) | all | Holes cannot be laid out in CSS; drawn as tight | won't do |
| Page / margin-relative vertical position | all | No page | planned (paged view) |
| Picture effects (artistic effects, recolour, soft edges, 3-D rotation) | docx | Not read | deferred |

## Shapes, WordArt, text boxes

| Feature | Formats | Reason | Status |
|---|---|---|---|
| 3-D shapes (scene3d / sp3d, bevels, extrusion) | docx, doc, rtf | Not read (ODT extrusion is drawn) | deferred |
| Connector routing (glue points) | docx, doc, rtf | Drawn as their saved geometry only | won't do |
| .doc / RTF floating groups and shapes with text boxes | doc, rtf | Rendered inline after the anchor, not positioned | deferred |
| .doc / RTF text boxes positioned | doc, rtf | Inline frame | deferred |
| Blip / pattern fill inside a:path shapes | docx | Not drawn; the fill below stays | deferred |
| Shapes in web archives | mhtml | VML is not drawn | deferred |

## Charts, SmartArt, objects

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Charts in .doc / RTF | doc, rtf | No chart reader in DocBinary/, Rtf/ (docx and ODT charts are read); the saved preview picture shows where there is one | deferred |
| 3-D line, area and bar (horizontal) charts in projection; Gouraud shading; ODT light directions | docx, odt | Columns and pies are projected and shaded; these keep the flat wall-and-floor drawing, one light from Word's default direction | deferred |
| Word's automatic plot-area layout; rotated / staggered / wrapped crowded category labels | docx, odt | Measured over the 19-chart fixture: ODT texts median 0.4 pt (x) / 0.3 pt (y) from LibreOffice's PDF; .docx without a saved layout median 2.6 / 2.4 pt from Word's PDF (largest gaps where Word sizes the plot area itself) | deferred |
| chartex region map without a geography cache | docx | No map data; the saved fallback picture is shown | won't do |
| SmartArt without a saved drawing part | docx | Layout engine not implemented; empty frame of the right size | deferred |
| OLE objects as live content (embedded sheets, media; MathType / Equation 3.0 equations) | all | Only the saved preview picture is shown (Office Math and LibreOffice formulas are drawn as MathML, see the supported list) | won't do |

## Tracked changes, comments, ruby, math

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Show markup / original view toggle for tracked changes | docx, doc, rtf, odt | Only Word's final view (planned as a toggle in the coverage plan) | planned |
| Comments | docx, doc, rtf, odt | Not shown in a reading view (references skipped) | deferred |
| Equations: exact Word layout (m:eqArr alignment marks inside nested elements, m:argSz, m:grow, line breaking of long display equations, the document's math font settings m:mathPr) | docx, doc, rtf | MathML Core has no counterpart for these; Chromium lays out the MathML (the equation and its text are shown and searched) | deferred |
| Equations in .doc saved without equationXML (Word 2003 and older) | doc | No OMML in the file; the saved picture is shown | won't do |
| Ruby | doc, rtf | Not read (EQ \o fields still compute) | deferred |

## Document level

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Encrypted files | docx, doc, odt | Rejected with a clear "encrypted" error (no password support) | won't do |
| Word 6.0 / 95, Word for Windows 1.x / 2.0 | doc | Rejected with a clear message; Word 6/95 is feasible (no Microsoft spec) | deferred |
| Word 2003 XML, Flat OPC | doc, docx | Rejected / not read | deferred |
| HTML files and HTML saved as .doc | html | Deliberately out of scope (FastDocView has no HTML support) | won't do |
| Malformed XML recovery | odt/fodt | content.xml / styles.xml fail with an error naming the part (docx keeps the content before the first error) | deferred |
| Exact Word layout metrics | all | Flow view on purpose | won't do |
| Loading a whole document at once | all | Breaks the windowed paradigm | won't do |
