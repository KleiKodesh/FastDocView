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
| Web archive sheet rules with attribute or sibling selectors, :last-child, pseudo-elements, @media | mhtml | The sheet is resolved per element into inline styles by a small checked grammar; Word's own sheets do not use these | deferred |
| Bullet whose symbol font is set only for complex script | doc | Shows no glyph (LibreOffice also reads the Latin font) | won't do |
| Blinking / animated text | all | Not drawn | won't do |

## Paragraphs

| Feature | Formats | Reason | Status |
|---|---|---|---|
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
| Picture effects: a preset camera's own rotation (used without a:rot), the light rig across a perspective picture (a few levels) and with both axes turned past 45 degrees (the grid's edge is used), other a:blip children (alphaBiLevel, alphaCeiling / Floor, alphaRepl, blur, clrRepl, fillOverlay, hsl, tint) | docx | Preset rotations not known here (drawn flat); Word's two-axis lighting is measured on a grid up to 45 degrees (LibreOffice draws an orthographic picture flat, unlit); Word's UI does not write the other blip children | deferred |
| Picture effects: duotone, colour change, transparency, soft edges, 3-D rotation (DrawingML-only effects) | doc, rtf | Word keeps them only in the .docx; its .doc / RTF save drops them (grey, black and white, brightness / contrast are drawn) | won't do |
| Picture colour effects on pictures inside a drawing group and on picture bullets | doc, rtf | Not read there (single inline / floating pictures are) | deferred |

## Shapes, WordArt, text boxes

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Connector routing (glue points) | docx, doc, rtf | Drawn as their saved geometry only | won't do |
| Half-transparent picture fill in a darken / lighten path | docx | The darkening layer is scaled by the picture's opacity (exact only for an opaque picture) | won't do |
| Shapes in web archives: the VML drawing itself | mhtml | Shown as a browser shows the archive: Word's HTML fallback (table cells and pictures of the shapes); the fallback is what a browser prints | won't do |

## Charts, SmartArt, objects

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Exact chart layout for Word-made .doc / RTF charts | doc, rtf | Median 1.4 pt (x) / 1.1-1.2 pt (y) from Word's PDF on word-charts-binary-v3 (.docx of the same charts 0.95 / 0.35; other fixtures 0.6-0.7 / 0.2-0.3). Left: the scatter x labels, the saved plot / legend positions (CHFRAMEPOS: read, not applied, they measured worse), 3-D value labels up to 6 pt off (front edge geometry) | deferred |
| Exact 3-D perspective scene placement (Word layout) | docx, doc, rtf, odt | Word's own placement of a perspective scene was not found, so Word's camera is placed as LibreOffice places its scene (projected box scaled and centred, VDiagram.cxx:409-423), and where Word's eye would come nearer than the box's diagonal LibreOffice's whole pipeline is used (2026-11-24). Residual vs Word (walls, word-charts-scene-grid / -scene-types): median 2.5 / 5.7 / 25 pt rms at perspective 30 / 60 / 120 (p90 13 / 24 / 46); labels on the perspective fixtures 0.2-3.7 pt median, but perspective category labels spread about 3 % narrower than Word's (max 6.6-13 pt; Word magnifies the front rows of deep scenes more, 2026-11-24 series probes: up to 34 pt at perspective 60, no rule found) and Line3D Deep 14.9 / 10.7 pt. 3-D pies not compared with Word | deferred |
| Exact right-angled 3-D scenes at rotX 25 or more | docx, doc, rtf | Series axis names are drawn (2026-11-17) and the scene is fitted with them as Word does (median 0.5 pt on word-charts-seraxis-v1; rotX up to 20 within 3 pt). At rotX 25-30 or rotY 45 Word's right-angled scene is 3-5 pt narrower than ours even without a series axis (word-charts-right-angled-v1, R 30 20 and the 2026-11-17 probes), so its names sit up to 10 pt off there. 2026-11-24 probes: the oblique is exactly sinY across / sinX up and the front fits as if the depth were 4 % larger, but that rule regresses other fixtures, so it is not applied | deferred |
| Trend line legend names in other interface languages | docx, doc, rtf | Office writes no name; Word shows one in its interface language. Hebrew and English are rendered from the reader's interface language (both sets read from Word's PDFs); any other interface language gets the English names | deferred |
| Exact chart layout for Word-made charts | docx | Text widths come from the fonts' own advance widths (Calibri, Cambria, Times New Roman, Arial); measured: Word-made charts median 0.3 pt (x) / 0.6 pt (y) from Word's PDF (right-to-left names 0.9 / 0.6 pt), LibreOffice-written .docx 1.0 / 0.9 pt, ODT vs LibreOffice 0.4 / 0.3 pt | deferred |
| chartex region map without a geography cache | docx | No map data; the saved fallback picture is shown | won't do |
| Exact SmartArt font size without a saved drawing part in rarer font cases | docx | Text is measured in the font Word uses for each character (Latin, complex-script and East Asian slots of the theme; bold / italic faces; kern table or GPOS kerning; GSUB shaping for complex scripts; East Asian line breaks and line height) and every tested layout gets Word's size (test-data/smartart-layouts). GSUB reverse chaining (Nastaliq) and the Indic shaping engine (Devanagari ... Malayalam: reordering, reph / half / below-base / pre-base forms, GPOS 'dist') are applied and match HarfBuzz on Nirmala UI. Still approximate: a style a family doesn't ship is measured in the nearest face, advances unchanged, as Word measures a synthesized bold / italic (Word 16 fixtures from scripts/make-smartart-shaping-fixtures.py: Sylfaen and Impact bold, Gabriola italic and bold, Devanagari, Kannada and Telugu get Word's size; a synthesized bold of a face whose instructions may change advances and that has no hdmx table is narrowed by 1 / 1.02, as GDI does); no Nastaliq font is installed to check one against Word; Southeast Asian / USE scripts (Khmer, Myanmar, Sinhala, Tibetan ...) are not reordered; Japanese line breaking follows Word's default kinsoku characters only; a font that is not installed or cached by Office (or no font folder, as in WASM) falls back to built-in tables for Calibri / Cambria / Times New Roman / Arial (ASCII, Arial's Hebrew), other characters at the font's average width | deferred |
| Exact SmartArt cycle node ring (other node counts) | docx | The cycle's node size follows Word's 1/2048 scale steps, but the ring's radius does not: Word's nodes are up to 0.35 pt off ours (cycle-8-nodes: Word's node is 485/4096 of the stated width where the steps give 484/4096; the radius is 0.01-0.2 % larger than one node width plus sibSp along the chord, with no rule found in five Word drawings - not the 1/2048 steps, EMU truncation or an iterative fit as LibreOffice does it, oox/source/drawingml/diagram/diagramlayoutatoms.cxx:10178-10190, which has no steps and is further from Word); matched within 4 pt | deferred |
| OLE objects as live content (embedded sheets, media; MathType / Equation 3.0 equations) | all | Only the saved preview picture is shown (Office Math and LibreOffice formulas are drawn as MathML, see the supported list) | won't do |

## Tracked changes, comments, ruby, math

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Show markup / original view toggle for tracked changes | docx, doc, rtf, odt | Only Word's final view is shown, as it prints; markup does not print | won't do (not printed) |
| Comments | docx, doc, rtf, odt | Comments do not print (references skipped) | won't do (not printed) |
| Equations: a host without font files falls back to tables for Cambria Math and DejaVu Math TeX Gyre; heights of single-line equations read ~0.85 of Word's PDF by the measuring method; matrix rows sit 1.04 x Word's pitch (1.226 em for Word's 1.18 in Cambria Math: Chromium's cell baseline alignment adds ~0.85 px a row) | docx, doc, rtf | Ink is read from the browser's text metrics, not Word's own layout | deferred |
| Equations in .doc saved without equationXML (Word 2003 and older) | doc | No OMML in the file; the saved picture is shown | won't do |

## Document level

| Feature | Formats | Reason | Status |
|---|---|---|---|
| Encrypted files | docx, doc, odt | Rejected with a clear "encrypted" error (no password support) | won't do |
| Word 6.0 / 95 drawing objects (DO records, PlcfdoaMom); Word 6 / Word 2 metafile pictures drawn (their box is shown) | doc | No OfficeArt in a Word 6 file. A metafile picture keeps its box at its size, as a WMF does in every format (a browser cannot draw one; bitmaps after the PIC are shown). No public Word 6 file with pictures or drawing objects to check against (2026-11-18: LibreOffice's and Apache POI's test data hold only the fuzz file ofz21385-1 with drawing objects and none with pictures; neither Word 16 nor LibreOffice can save Word 6/95) | deferred |
| Word 2003 XML: embedded fonts, w:docPr note separators | docx | Not translated; separators do not change the flow view | deferred |
| Flat OPC or Word 2003 XML in UTF-16 or a legacy code page | docx | Rejected with a clear message: Word writes both in UTF-8, and their blocks are read in place by byte offsets as UTF-8 | won't do |
| Word 2003 XML files over 1 GB | docx | Rejected with a clear message: note ids are made from the note's block offset (an int); Word itself opens documents up to 512 MB | won't do |
| HTML files and HTML saved as .doc | html | Deliberately out of scope (FastDocView has no HTML support) | won't do |
| Malformed XML recovery before office:text in a .fodt | fodt | The file's first fatal error comes before the body, so nothing is left to keep: a clear "not well-formed" error, as LibreOffice's parser stops there too | won't do |
| Exact Word layout metrics | all | Flow view on purpose | won't do |
| Loading a whole document at once | all | Breaks the windowed paradigm | won't do |
