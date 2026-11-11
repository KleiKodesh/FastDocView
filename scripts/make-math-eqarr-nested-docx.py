"""Writes the base of test-data/word-math/eqarr-nested-v1.docx: equation arrays with '&' inside nested elements (a delimiter,
a superscript base) and an equation array inside an equation array row. Word 16 (COM: open, SetCompatibilityMode(15),
SaveAs2 .docx and .pdf) shows a nested '&' as text and aligns only each array's own rows at their own '&' - checked in
its PDF (eqarr-nested-v1.pdf). Run: python scripts/make-math-eqarr-nested-docx.py <out.docx>"""
import sys, zipfile
def r(t): return f'<m:r><m:t>{t}</m:t></m:r>'
rows = ''.join(f'<m:e>{r(a)}<m:d><m:e>{r(b)}</m:e></m:d>{r(c)}</m:e>' for a, b, c in [('x', 'a&amp;b', '=1'), ('yy', 'c&amp;d', '=22')])
nested = f'<m:eqArr><m:e>{r("p&amp;=")}<m:eqArr><m:e>{r("q&amp;=1")}</m:e><m:e>{r("qq&amp;=2")}</m:e></m:eqArr></m:e><m:e>{r("ppp&amp;=3")}</m:e></m:eqArr>'
sup = f'<m:eqArr><m:e><m:sSup><m:e>{r("z&amp;")}</m:e><m:sup>{r("2")}</m:sup></m:sSup>{r("=5")}</m:e><m:e>{r("zz&amp;=6")}</m:e></m:eqArr>'
paras = ''.join(f'<w:p><w:r><w:t xml:space="preserve">case {l}: </w:t></w:r><m:oMathPara><m:oMath>{m}</m:oMath></m:oMathPara></w:p>' for l, m in [('delim', f'<m:eqArr>{rows}</m:eqArr>'), ('nested', nested), ('sup', sup)])
NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
doc = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{paras}<w:sectPr/></w:body></w:document>'
types = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
rels = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
with zipfile.ZipFile(sys.argv[1], 'w') as z:
    z.writestr('[Content_Types].xml', types); z.writestr('_rels/.rels', rels); z.writestr('word/document.xml', doc)
