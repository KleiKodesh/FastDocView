# Test documents

Third-party test documents used by the corpus pass (`WordDocumentRandomAccess.Tests -- --corpus <list>`).
They are deliberately hard: edge cases, broken packages, fuzzer crash files, encrypted files, 5,000
nested tables. Every one must open or fail with a clear error — never crash the process.

| Folder | Source | Commit | License |
|---|---|---|---|
| `libreoffice/` | LibreOffice core, `sw/qa/**/data` (.docx) | 504dd9168724bd5bbf471d934f9dc0c65a5259bc | MPL 2.0 (`libreoffice/COPYING.MPL`) |
| `apache-poi/` | Apache POI, `test-data/document` (.docx/.docm/.dotx) | 07f1631dd66f647aacf2f43e3f408fe033aac4d0 | Apache 2.0 (`apache-poi/LICENSE`, `apache-poi/NOTICE`) |

Files are unmodified copies; LibreOffice keeps its source folder layout.
