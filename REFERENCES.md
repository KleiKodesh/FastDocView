# References — in order of authority

Word's source code is closed and unavailable. These are what we verify against, highest authority
first. Every format fact implemented must be checked in one of them (CLAUDE.md: never from memory),
and research reports must say which source a fact came from. Links checked 2026-10-28
(NetFree blocks a few for us — marked; they are still the right sources).

## 1. Microsoft Open Specifications — Microsoft's own description of what Word reads and writes

Published under the Open Specification Promise; the closest thing to Word's source.
Index: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ and
https://learn.microsoft.com/en-us/openspecs/office_standards/

| Spec | Use it for | Link |
|---|---|---|
| **[MS-DOC]** Word (.doc) Binary File Format | everything .doc: FIB, piece table, FKPs, sprms, STSH, lists, notes, fields | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/ |
| **[MS-OI29500]** Office Implementation Information for ISO/IEC 29500 | **where Word deviates from / extends the .docx standard** — check first for any Word-specific .docx behaviour | https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/ |
| **[MS-OE376]** Office Implementation Information for ECMA-376 | same, against the ECMA-376 1st edition (older Word versions) | https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/ |
| **[MS-DOCX]** Word Extensions to the Office Open XML File Format | `w14:` / `w15:` / `w16*:` elements (text effects, collapsed headings, …) | https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/ |
| **[MS-ODRAWXML]** Office Drawing Extensions to Office Open XML | DrawingML extensions (text boxes, shapes, `mc:` content in drawings) | https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/ |
| **[MS-OFFSTANDLP]** Office Standards Support Overview | map of which implementation-notes doc covers what | https://learn.microsoft.com/en-us/openspecs/office_standards/ms-offstandlp/ |
| **[MS-OSHARED]** Office Common Data Types | MSONFC number formats, shared structures used by [MS-DOC] | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/ |
| **[MS-CFB]** Compound File Binary Format | the OLE container of .doc | https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/ |
| **[MS-ODRAW]** Office Drawing Binary File Format | .doc shapes / pictures (FSPA, OfficeArt) | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw/ |
| **[MS-OFFCRYPTO]** Office Document Cryptography Structure | detecting encrypted .doc / .docx (we reject them) | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/ |
| **[MS-OFFDI]** Office File Formats Overview | orientation across the Office binary formats | https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-offdi/ |

## 2. International standards

| Standard | Use it for | Link |
|---|---|---|
| **ISO/IEC 29500** (Office Open XML), Parts 1–4 — freely downloadable from ISO's ITTF | the .docx standard itself (Part 1 WordprocessingML §17, DrawingML §20; Part 2 packaging/OPC; Part 4 transitional) | https://standards.iso.org/ittf/PubliclyAvailableStandards/ |
| **ECMA-376** (same content, Ecma edition) | as above; editions 1–5 | https://ecma-international.org/publications-and-standards/standards/ecma-376/ |
| **OASIS ODF 1.3** (OpenDocument) — ODF 1.2 is ISO/IEC 26300:2015 | .odt / .ott / .fodt: schema, packaging, encryption | https://docs.oasis-open.org/office/OpenDocument/v1.3/ |
| **Microsoft RTF Specification** — 1.9.1 (2008, Word 2007) is the last; its download is NetFree-blocked here, the 1.5 text (Word 97) is readable | .rtf: syntax, `\uc` skipping, charsets, tables, lists, notes | 1.5: https://www.biblioscape.com/rtf15_spec.htm ; 1.9.1: https://www.microsoft.com/en-us/download/details.aspx?id=10725 |
| **RFC 1951** DEFLATE | the inflater (SeekableDeflate) | https://www.rfc-editor.org/rfc/rfc1951 |
| **PKWARE APPNOTE.TXT** (zip, Zip64) | zip central directory / entries *(NetFree-blocked here)* | https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT |
| **Unicode UAX #9** Bidirectional Algorithm | RTL / mixed Hebrew-Latin ordering in the HTML | https://www.unicode.org/reports/tr9/ |
| **CSS Writing Modes 4**, **CSS Text Decoration 4** (W3C) | `direction`, logical margins, `text-decoration-*`, `text-emphasis` mappings | https://www.w3.org/TR/css-writing-modes-4/ , https://www.w3.org/TR/css-text-decor-4/ |

## 3. Word itself — the ground truth for behaviour

Word 16 is installed on the dev machine. When the specs are silent or two sources disagree (or a
fact is marked "unverified"), build a minimal test document, let Word open / save / convert it
(COM automation, as KitveiHakodesh's Word→PDF path already does), and compare with our output.
The test file and the observed result go into the harness.

## 4. Open-source implementations — second-hand models of Word's behaviour

| Project | Use it for | Link |
|---|---|---|
| **LibreOffice** — `sw/source/filter/ww8` (.doc: `sprmids.hxx`, `ww8scan.cxx`, `ww8par*.cxx`), `sw/source/writerfilter` (.docx: `dmapper/DomainMapper.cxx`, `ThemeHandler.cxx`, `ooxml/model.xml`), `xmloff/source/text` (.odt) | how a mature importer interprets each construct; opcode tables (look up, never recall) | https://github.com/LibreOffice/core |
| **LibreOffice RTF** — `sw/source/writerfilter/rtftok` (`rtftokenizer.cxx` keyword table + tokenizer, `rtfdispatch*.cxx`, `rtfdocumentimpl.cxx`, `rtfcharsets.cxx`); export `sw/source/filter/ww8/rtfattributeoutput.cxx` | how each RTF control word is read: destination set, run types (`\rtlch`/`\ltrch`), charsets, list levels | https://github.com/LibreOffice/core/tree/master/sw/source/writerfilter/rtftok |
| **LibreOffice test documents** — `sw/qa/extras/*/data` | thousands of real edge-case .doc / .docx / .odt files with known expectations | https://github.com/LibreOffice/core/tree/master/sw/qa/extras/ooxmlexport/data |
| **zlib** `examples/zran.c` | the checkpoint / resume technique | https://github.com/madler/zlib/blob/develop/examples/zran.c |
| **Apache POI** (HWPF / XWPF) | a second .doc / .docx reading of the same structures | https://github.com/apache/poi |
| **Open XML SDK** (Microsoft) | schema-level .docx reading/validation only — no rendering logic | https://github.com/dotnet/Open-XML-SDK |

## 5. Readable mirrors (convenience, not authority)

- docx4java's HTML rendering of the ECMA-376 WordprocessingML reference (element pages with the
  spec text): https://webapp.docx4java.org/OnlineDemo/ecma376/WordML/index.html — quote the spec
  section, not the mirror, when it matters.
- Word 6.0 / 95 (.doc with nFib < 0xC1) has **no** Microsoft specification; only reverse-engineered
  descriptions (LibreOffice's ww8 "Ver67" paths, the UK National Archives PRONOM entry fmt/39 —
  https://www.nationalarchives.gov.uk/PRONOM/fmt/39, NetFree-blocked here). Treat as unverified.
