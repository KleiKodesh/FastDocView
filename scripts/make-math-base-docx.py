"""Writes math-base.docx: one paragraph per Office Math (OMML, ECMA-376 Part 1 22.1) construct, hand-written OMML.
make-math-fixtures.ps1 opens it in Word 16 (COM), adds a footnote, a table cell and a text box with equations, and
saves Word's own .docx / .doc / .rtf / .pdf. Run: python scripts/make-math-base-docx.py test-data/word-math/math-base.docx"""
import sys, zipfile

def r(text, rpr=""):
    return f'<m:r>{("<m:rPr>" + rpr + "</m:rPr>") if rpr else ""}<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t xml:space="preserve">{text}</m:t></m:r>'

def arg(tag, *content):
    return f'<m:{tag}>{"".join(content)}</m:{tag}>'

def e(*c): return arg("e", *c)

CASES = [
    ("frac bar", f'<m:f>{arg("num", r("a"))}{arg("den", r("b"))}</m:f>'),
    ("frac skw", f'<m:f><m:fPr><m:type m:val="skw"/></m:fPr>{arg("num", r("a"))}{arg("den", r("b"))}</m:f>'),
    ("frac lin", f'<m:f><m:fPr><m:type m:val="lin"/></m:fPr>{arg("num", r("a"))}{arg("den", r("b"))}</m:f>'),
    ("frac noBar", f'<m:f><m:fPr><m:type m:val="noBar"/></m:fPr>{arg("num", r("n"))}{arg("den", r("k"))}</m:f>'),
    ("sqrt", f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>{e(r("x"))}</m:rad>'),
    ("root", f'<m:rad>{arg("deg", r("3"))}{e(r("y"))}</m:rad>'),
    ("sup", f'<m:sSup>{e(r("x"))}{arg("sup", r("2"))}</m:sSup>'),
    ("sub", f'<m:sSub>{e(r("x"))}{arg("sub", r("i"))}</m:sSub>'),
    ("subsup", f'<m:sSubSup>{e(r("x"))}{arg("sub", r("i"))}{arg("sup", r("2"))}</m:sSubSup>'),
    ("pre", f'<m:sPre>{arg("sub", r("1"))}{arg("sup", r("n"))}{e(r("Y"))}</m:sPre>'),
    ("integral", f'<m:nary><m:naryPr><m:limLoc m:val="subSup"/></m:naryPr>{arg("sub", r("0"))}{arg("sup", r("1"))}{e(r("f"), r("dx"))}</m:nary>'),
    ("sum", f'<m:nary><m:naryPr><m:chr m:val="&#x2211;"/><m:limLoc m:val="undOvr"/></m:naryPr>{arg("sub", r("i=1"))}{arg("sup", r("n"))}{e(r("i"))}</m:nary>'),
    ("product", f'<m:nary><m:naryPr><m:chr m:val="&#x220F;"/><m:limLoc m:val="subSup"/><m:supHide m:val="1"/></m:naryPr>{arg("sub", r("k"))}<m:sup/>{e(r("k"))}</m:nary>'),
    ("delim", f'<m:d><m:dPr><m:begChr m:val="["/><m:sepChr m:val=";"/><m:endChr m:val="}}"/></m:dPr>{e(r("a"))}{e(r("b"))}</m:d>'),
    ("paren", f'<m:d>{e(r("a+b"))}</m:d>'),
    ("func", f'<m:func><m:fName>{r("sin", "<m:sty m:val=\"p\"/>")}</m:fName>{e(r("θ"))}</m:func>'),
    ("limlow", f'<m:func><m:fName><m:limLow>{e(r("lim", "<m:sty m:val=\"p\"/>"))}{arg("lim", r("n→∞"))}</m:limLow></m:fName>{e(r("a"))}</m:func>'),
    ("limupp", f'<m:limUpp>{e(r("x"))}{arg("lim", r("def"))}</m:limUpp>'),
    ("acc", f'<m:acc><m:accPr><m:chr m:val="&#x0303;"/></m:accPr>{e(r("a"))}</m:acc>'),
    ("acc hat", f'<m:acc>{e(r("v"))}</m:acc>'),
    ("bar top", f'<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>{e(r("z"))}</m:bar>'),
    ("bar bot", f'<m:bar>{e(r("w"))}</m:bar>'),
    ("box", f'<m:box>{e(r("q"))}</m:box>'),
    ("borderbox", f'<m:borderBox>{e(r("E=m"), "<m:sSup>" + e(r("c")) + arg("sup", r("2")) + "</m:sSup>")}</m:borderBox>'),
    ("eqarr", f'<m:eqArr>{e(r("x&amp;=1"))}{e(r("y&amp;=22"))}</m:eqArr>'),
    ("matrix", f'<m:d><m:e><m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="2"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs></m:mPr><m:mr>{e(r("1"))}{e(r("0"))}</m:mr><m:mr>{e(r("0"))}{e(r("1"))}</m:mr></m:m></m:e></m:d>'),
    ("groupchr", f'<m:groupChr>{e(r("a+b+c"))}</m:groupChr>'),
    ("groupchr top", f'<m:groupChr><m:groupChrPr><m:chr m:val="&#x23DE;"/><m:pos m:val="top"/><m:vertJc m:val="bot"/></m:groupChrPr>{e(r("k"))}</m:groupChr>'),
    ("phant", f'{r("a")}<m:phant><m:phantPr><m:show m:val="0"/></m:phantPr>{e(r("bbb"))}</m:phant>{r("c")}'),
    ("styles", r("P", '<m:sty m:val="p"/>') + r("B", '<m:sty m:val="b"/>') + r("I", '<m:sty m:val="i"/>') + r("X", '<m:sty m:val="bi"/>')),
    ("scripts", r("R", '<m:scr m:val="double-struck"/><m:sty m:val="p"/>') + r("F", '<m:scr m:val="fraktur"/><m:sty m:val="p"/>') + r("L", '<m:scr m:val="script"/><m:sty m:val="p"/>') + r("S", '<m:scr m:val="sans-serif"/>') + r("M", '<m:scr m:val="monospace"/><m:sty m:val="p"/>')),
    ("normal text", r("if x", '<m:nor/>') + r("lit", '<m:lit/>')),
    ("colour", '<m:r><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:color w:val="FF0000"/></w:rPr><m:t>red</m:t></m:r>'),
]

def para(label, math, display):
    if display:
        body = f'<m:oMathPara><m:oMathParaPr><m:jc m:val="{display}"/></m:oMathParaPr><m:oMath>{math}</m:oMath></m:oMathPara>'
    else:
        body = f'<m:oMath>{math}</m:oMath>'
    return f'<w:p><w:r><w:t xml:space="preserve">{label}: </w:t></w:r>{body}<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>'

paras = [para(f"case {label}", m, None) for label, m in CASES]
paras.append(f'<w:p><m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>{r("left=1")}</m:oMath></m:oMathPara></w:p>')
paras.append(f'<w:p><m:oMathPara><m:oMath>{r("center=2")}</m:oMath><m:oMath>{r("second=3")}</m:oMath></m:oMathPara></w:p>')
paras.append(f'<w:p><m:oMathPara><m:oMathParaPr><m:jc m:val="right"/></m:oMathParaPr><m:oMath>{r("right=4")}</m:oMath></m:oMathPara></w:p>')
paras.append('<w:p><w:r><w:t>last paragraph</w:t></w:r></w:p>')

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
document = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{"".join(paras)}<w:sectPr/></w:body></w:document>'
types = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
         '<Default Extension="xml" ContentType="application/xml"/>'
         '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
rels = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
with zipfile.ZipFile(sys.argv[1], "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", types)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/document.xml", document)
