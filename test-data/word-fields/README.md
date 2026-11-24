# Word field fixtures

`fields-v1.doc`, `fields-v1.rtf` and `fields-v1.docx` were saved by Microsoft Word 16 from one document built by
`make-fields-fixtures.ps1` (COM automation, one field per paragraph; the script prints Word's own results).
Made for this project; same license as the repository.

`fields-lists-v1.doc` (REF with the number switches on a three-level list) was saved by Word 16 from
`make-fields-lists-fixture.ps1`, which prints Word's own results.

`eq-layout-v1.docx`, `.doc`, `.rtf` and `.pdf` (EQ field layout: overstrike alignment, integral limits, one-sided brackets,
bracket sizes) were saved by Word 16 from `make-eq-layout-fixture.ps1` (hidden automation, one "case NAME: {EQ} end"
paragraph per construct); the PDF is Word's own drawing of them.
