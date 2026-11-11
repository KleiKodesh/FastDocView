"""Writes test-data/word-math/brackets-base.docx: brackets ( [ { | and a growing integral around 3-, 5- and 10-row
matrices / fraction stacks (one display equation per paragraph, "case <label>: ... end"), for the stretched-glyph check
against Word's PDF. make-math-brackets.ps1 has Word 16 save it with Cambria Math and with DejaVu Math TeX Gyre.
Run: python scripts/make-math-brackets-docx.py test-data/word-math/brackets-base.docx"""
import sys, zipfile
def r(t): return f'<m:r><m:t>{t}</m:t></m:r>'
def matrix(rows):
    return '<m:m>' + ''.join(f'<m:mr><m:e>{r(str(i))}</m:e><m:e>{r("x")}</m:e></m:mr>' for i in range(1, rows + 1)) + '</m:m>'
def stack(rows):
    inner = r("a")
    for _ in range(rows - 1): inner = f'<m:f><m:num>{r("1")}</m:num><m:den>{inner}</m:den></m:f>'
    return inner
cases = []
for rows in (3, 5, 10):
    for name, beg, end in (("paren", "(", ")"), ("bracket", "[", "]"), ("brace", "{", "}"), ("bar", "|", "|")):
        cases.append((f"{name}{rows}", f'<m:d><m:dPr><m:begChr m:val="{beg}"/><m:endChr m:val="{end}"/></m:dPr><m:e>{matrix(rows)}</m:e></m:d>'))
    cases.append((f"integral{rows}", f'<m:nary><m:naryPr><m:grow m:val="1"/><m:subHide m:val="1"/><m:supHide m:val="1"/></m:naryPr><m:sub/><m:sup/><m:e>{stack(rows)}</m:e></m:nary>'))
paras = ''.join(f'<w:p><w:r><w:t xml:space="preserve">case {label}: </w:t></w:r><m:oMath>{m}</m:oMath><w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>' for label, m in cases)
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{paras}<w:sectPr/></w:body></w:document>'
types = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
         '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
rels = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
with zipfile.ZipFile(sys.argv[1], "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", types); z.writestr("_rels/.rels", rels); z.writestr("word/document.xml", doc)
