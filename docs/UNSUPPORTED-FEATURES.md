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
| w14 3-D text (props3d / scene3d): bevel, extrusion, contour, material, lighting, camera | docx | Word 16 draws 3-D text as a lit raster picture of the run (XPS output). With Word's front camera the extrusion is hidden behind the letters and the bevel only shades their face; CSS has no lighting or bevel for text, so the plain letters in the run colour are shown. .doc / RTF do not store it | won't do |
| w14 reflection blur, scale / skew (sx, sy, kx, ky), fade direction | docx | box-reflect mirrors a box as it is: it can fade but not blur or transform the mirror image (Word's presets blur 0.5-10 pt and keep sx 100 %, sy -100 %) | won't do |
| Kerning | docx, doc, rtf, odt | Ignored by design (audit set); the browser's own kerning applies | won't do |
| East-Asian font slot (eastAsiaTheme, east-asian fonts), eastAsianLayout, fitText | docx, doc | Model has no East-Asian slot; not needed for Hebrew books | deferred |
| Web archive sheet rules with attribute or sibling selectors, :last-child, pseudo-elements, @media | mhtml | The sheet is resolved per element into inline styles by a small checked grammar; Word's own sheets do not use these | deferred |
| Bullet whose symbol font is set only for complex script | doc | Shows no glyph (LibreOffice also reads the Latin font) | won't do |
| Blinking / animated text | all | Not drawn | won't do |

## Paragraphs

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Tab leaders and right / centre / decimal tab alignment | docx, doc, rtf, odt | Flow HTML has only `tab-size`; true tab layout needs a line layout engine | deferred |
| Keep with next, keep lines, widow / orphan control, page break before | docx, doc, rtf, odt | Pagination: belongs to the paged view | planned (paged view) |
| Frames (framePr) other than drop caps: absolute position, wrap | docx, doc, rtf | Frames flow as paragraphs; positioning is page-bound | planned (paged view) |
| Hebrew hyphenation | all | Chromium has no Hebrew hyphenation dictionary | won't do |
| Line numbering, snap to grid, document grid | docx, doc, rtf, odt | Page-bound / East-Asian layout | planned (paged view) |

## Lists

| Feature | Formats | Reason | Status |
|---|---|---|---|
| ApplyParagraphMarkFormatToNumbering setting | docx | Not read (LibreOffice writes it only for Word imports) | won't do |

## Tables

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Floating table position (tblpPr) and text wrap around tables | docx, doc, rtf | Ignored by design (frame / table positioning) | planned (paged view) |
| Stacked cell text (style:direction ttb) | odt | LibreOffice Writer does not read it on cells either (the reference print is horizontal) | won't do |
| Cells merged down drawn as one cell | rtf, doc | An outer row is one block (windowed): a cell merged down is cells without the line between them, not one rowspan cell (a nested table gets a real rowspan) | won't do (windowed rule) |
| Repeat header row, row "can't split" | all | Page-bound | planned (paged view) |
| Margins below -1 in, negative padding, very large sizes, calc() in web archive CSS | mhtml | Clamped to safe bounds (WebArchiveCssLength) so text stays in the document area; a page that relies on them lays out a little differently | won't do |
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
| Restart note numbering per page | docx, doc, rtf, odt | Page-bound: counted like each section (right on a section's first page, continues on later pages); odt shows the citation LibreOffice wrote | planned (paged view) |
| Note mark inside Fontwork text, a dimension label or a shape left out of the drawing | odt/fodt | Drawn as plain text or not drawn; the note is in the pane only | won't do |

## Fields

| Feature | Formats | Reason | Status |
|---|---|---|---|
| INCLUDETEXT, INCLUDEPICTURE | docx, doc, rtf | Never fetched: a document must not read other files | won't do |
| USERNAME / USERINITIALS / USERADDRESS without an argument | docx, doc, rtf | Word's user settings, not the document's | won't do |
| STYLEREF page-relative search and \l | docx, doc, rtf | Flow view has no page | planned (paged view) |
| Lookups during the first open (REF to a bookmark not indexed yet, TOC, INDEX) | docx, doc, rtf | Show nothing until the index completes, instead of Word's "not found" | won't do (by design) |
| Fields | txt, mhtml | Format has none | n.a. |

## Links, bookmarks, TOC

| Feature | Formats | Reason | Status |
|---|---|---|---|
| TOC page numbers | docx, doc, rtf | Placeholders until pages exist | planned (paged view) |

## Pictures

| Feature | Formats | Reason | Status |
|---|---|---|---|
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
| Half-transparent picture fill in a darken / lighten path | docx | The darkening layer is scaled by the picture's opacity (exact only for an opaque picture) | won't do |
| Shapes in web archives | mhtml | VML is not drawn | deferred |

## Charts, SmartArt, objects

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Exact chart layout for Word-made .doc / RTF charts | doc, rtf | Median 1.4 pt (x) / 1.1-1.2 pt (y) from Word's PDF on word-charts-binary-v3 (.docx of the same charts 0.95 / 0.35; other fixtures 0.6-0.7 / 0.2-0.3). Left: the scatter x labels, the saved plot / legend positions (CHFRAMEPOS: read, not applied, they measured worse), 3-D value labels up to 6 pt off (front edge geometry) | deferred |
| .doc / RTF chart colours exactly as Word shows them | doc, rtf | Office 2007+ saves a non-palette text colour as its nearest palette colour in the BIFF8 records (both RGB and index); the exact one is only in its DrawingML extension records (CrtMlFrt), not read | deferred |
| Exact 3-D line / area perspective view (Word layout) | docx, doc, rtf | Right-angled scenes match Word (median 0.2 / 0.7 pt); with perspective the old fitted camera stays (median about 3.5 / 2 pt, 12-30 pt at perspective 60+ with strong rotation). Word's camera is now known (2026-11-16 wall-corner probes: field of view = c:perspective / 2 degrees over the plot height, eye on the plot centre, distance in points), but not where Word places the scene in it: a centred fit gives labels 0.6-0.8 pt on the perspective fixtures yet regresses the deep, strongly rotated one (Line3D Deep max 22 pt vs 12), so it is not switched on | deferred |
| 3-D series axis labels (c:serAx) | docx, doc, rtf | not drawn; Word's room for them on the right is kept (2026-11-16) | planned |
| Trend line legend names in other interface languages | docx, doc, rtf | Office writes no name; Word shows one in its interface language. Hebrew and English are rendered from the reader's interface language (both sets read from Word's PDFs); any other interface language gets the English names | deferred |
| Exact chart layout for Word-made charts | docx | Text widths come from the fonts' own advance widths (Calibri, Cambria, Times New Roman, Arial); measured: Word-made charts median 0.3 pt (x) / 0.6 pt (y) from Word's PDF (right-to-left names 0.9 / 0.6 pt), LibreOffice-written .docx 1.0 / 0.9 pt, ODT vs LibreOffice 0.4 / 0.3 pt | deferred |
| chartex region map without a geography cache | docx | No map data; the saved fallback picture is shown | won't do |
| SmartArt without a saved drawing part | docx | Layout engine not implemented; empty frame of the right size | deferred |
| OLE objects as live content (embedded sheets, media; MathType / Equation 3.0 equations) | all | Only the saved preview picture is shown (Office Math and LibreOffice formulas are drawn as MathML, see the supported list) | won't do |

## Tracked changes, comments, ruby, math

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Show markup / original view toggle for tracked changes | docx, doc, rtf, odt | Only Word's final view is shown, as it prints; markup does not print | won't do (not printed) |
| Comments | docx, doc, rtf, odt | Comments do not print (references skipped) | won't do (not printed) |
| Equations: a stretched operator around content whose ink lands right at a variant boundary can take the next size (the 3-row brace in Cambria Math, 1.48 x Word's width; 29 of 30 test cases within 15%); a host without font files falls back to tables for Cambria Math and DejaVu Math TeX Gyre; heights of single-line equations read ~0.85 of Word's PDF by the measuring method | docx, doc, rtf | Ink is read from the browser's text metrics, not Word's own layout | deferred |
| Equations in .doc saved without equationXML (Word 2003 and older) | doc | No OMML in the file; the saved picture is shown | won't do |
| Search inside a ruby base | doc, rtf | Word stores the ruby as an EQ field: the base is field code, which search skips (the ruby is drawn) | deferred |

## Document level

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Encrypted files | docx, doc, odt | Rejected with a clear "encrypted" error (no password support) | won't do |
| Word 6.0 / 95, Word for Windows 1.x / 2.0 | doc | Rejected with a clear message; Word 6/95 is feasible (no Microsoft spec) | deferred |
| Word 2003 XML: embedded fonts, w:docPr note separators | docx | Not translated; separators do not change the flow view | deferred |
| Flat OPC in UTF-16 or a legacy code page | docx | Rejected with a clear message: Word writes it in UTF-8 and its parts are read in place as UTF-8 (Word 2003 XML is translated, so any encoding works) | won't do |
| HTML files and HTML saved as .doc | html | Deliberately out of scope (FastDocView has no HTML support) | won't do |
| Malformed XML recovery before office:text in a .fodt | fodt | The file's first fatal error comes before the body, so nothing is left to keep: a clear "not well-formed" error, as LibreOffice's parser stops there too | won't do |
| Exact Word layout metrics | all | Flow view on purpose | won't do |
| Loading a whole document at once | all | Breaks the windowed paradigm | won't do |
