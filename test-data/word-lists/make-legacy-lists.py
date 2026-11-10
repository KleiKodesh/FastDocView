"""Step of make-legacy-lists-fixture.ps1: a .docx whose levels are Word 6 / 95 levels (w:legacy, ECMA-376 17.9.10) with a range of
legacySpace / legacyIndent values, under a hanging (left 1440, hanging 720) and a flat (left 1000) indent; Word 16 re-saves it."""
import zipfile,sys
out=sys.argv[1]
import os
variants=[("none",None),("s0i0","0","0"),("s0i1000","0","1000"),("s1000i0","1000","0"),("s100i1000","100","1000"),("s0i200","0","200"),("s700i0","700","0"),("s0i1500","0","1500")]
INDENTS=[("1440","720"),("1000","0")]
CT='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'
RELS='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
DRELS='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>'
W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
nums='';abs_=''
paras=''
for i,(v,(LEFT,HANG)) in enumerate([(v,ind) for ind in INDENTS for v in variants]):
    leg='' if v[1] is None else '<w:legacy w:legacy="1" w:legacySpace="%s" w:legacyIndent="%s"/>'%(v[1],v[2])
    abs_+='<w:abstractNum w:abstractNumId="%d"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%%1."/>%s<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="%s" w:hanging="%s"/></w:pPr></w:lvl></w:abstractNum>'%(i,leg,LEFT,HANG)
    nums+='<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>'%(i+1,i)
    for n in range(2):
        paras+='<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr></w:pPr><w:r><w:t>legacy %s item</w:t></w:r></w:p>'%(i+1,v[0])
z=zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED)
z.writestr('[Content_Types].xml',CT);z.writestr('_rels/.rels',RELS);z.writestr('word/_rels/document.xml.rels',DRELS)
z.writestr('word/numbering.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering %s>%s%s</w:numbering>'%(W,abs_,nums))
z.writestr('word/document.xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s<w:sectPr><w:pgMar w:left="1440" w:right="1440" w:top="1440" w:bottom="1440"/></w:sectPr></w:body></w:document>'%(W,paras))
z.close()
