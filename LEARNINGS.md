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
