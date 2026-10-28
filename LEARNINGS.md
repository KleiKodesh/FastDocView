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
