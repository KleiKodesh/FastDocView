r"""list-overrides-v1: RTF list cases Word 16 prints (FastDocView RtfListOverridesFidelityTest).
  python make-list-overrides.py list-overrides-v1.docx         -> A: level 3 restarts only after level 1 (w:lvlRestart 1);
      B: a list linked to a numbering style (w:numStyleLink / w:styleLink); C: a w:lvlOverride holding a whole w:lvl ((%1), start 5);
      D: bullet / number labels with their own colour and size.
  Word 16 (hidden, DisplayAlerts 0, SaveAs2 format 6) -> list-overrides-v1.rtf;
  strip-list-labels.py list-overrides-v1.rtf list-overrides-nolisttext-v1.rtf (every label counted);
  python make-list-overrides.py --empty-linked list-overrides-nolisttext-v1.rtf list-overrides-stylelink-v1.rtf (the linked list keeps only
      its \liststyleid). Word 16 printed both .rtf files to the same text: list-overrides-v1.expected.txt (pdftotext -layout, spaces folded)."""
import zipfile,sys
if sys.argv[1] == '--empty-linked':
    B = chr(92)
    s = open(sys.argv[2], encoding='latin1').read()
    i = s.find(B + 'liststyleid'); j = s.rfind('{' + B + 'list' + B + 'listtemplateid', 0, i)
    group, out, k = s[j:i], '', 0
    while k < len(group):
        if group.startswith('{' + B + 'listlevel', k):
            depth = 0
            while True:
                c = group[k]
                if c == B: k += 2; continue
                if c == '{': depth += 1
                elif c == '}':
                    depth -= 1
                    if depth == 0: k += 1; break
                k += 1
            continue
        out += group[k]; k += 1
    open(sys.argv[3], 'w', encoding='latin1', newline='').write(s[:j] + out + s[i:])
    sys.exit(0)
W='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
def lvl(i,fmt,text,start=1,extra='',rpr=''):
    return f'<w:lvl w:ilvl="{i}"><w:start w:val="{start}"/><w:numFmt w:val="{fmt}"/>{extra}<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="{720*(i+1)}" w:hanging="360"/></w:pPr>{rpr}</w:lvl>'
def nine(fmt='decimal',special={}):
    out=''
    for i in range(9):
        out+=special.get(i) or lvl(i,fmt,''.join(f'%{k+1}.' for k in range(i+1)))
    return out
abs0=f'<w:abstractNum w:abstractNumId="0">{nine(special={2:lvl(2,"decimal","%1.%2.%3.",extra=chr(60)+"w:lvlRestart w:val=\"1\"/>")})}</w:abstractNum>'
abs1=f'<w:abstractNum w:abstractNumId="1"><w:styleLink w:val="MyList"/>{nine("upperRoman",{0:lvl(0,"upperRoman","%1)"),1:lvl(1,"lowerLetter","%2]")})}</w:abstractNum>'
abs2=f'<w:abstractNum w:abstractNumId="2"><w:numStyleLink w:val="MyList"/></w:abstractNum>'
abs3=f'<w:abstractNum w:abstractNumId="3">{nine()}</w:abstractNum>'
abs4=f'<w:abstractNum w:abstractNumId="4">{nine(special={0:lvl(0,"bullet","\u25cf",rpr="<w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/><w:color w:val=\"FF0000\"/><w:sz w:val=\"40\"/></w:rPr>"),1:lvl(1,"decimal","%2.",rpr="<w:rPr><w:color w:val=\"0000FF\"/><w:sz w:val=\"16\"/></w:rPr>")})}</w:abstractNum>'
nums='<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>'
nums+='<w:num w:numId="4"><w:abstractNumId w:val="3"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/>'+lvl(0,"lowerLetter","(%1)",start=5)+'</w:lvlOverride></w:num><w:num w:numId="5"><w:abstractNumId w:val="4"/></w:num>'
numbering=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {W}>{abs0}{abs1}{abs2}{abs3}{abs4}{nums}</w:numbering>'
styles=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="numbering" w:styleId="MyList"><w:name w:val="MyList"/><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="ListPara"><w:name w:val="ListPara"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="3"/></w:numPr></w:pPr></w:style></w:styles>'
def p(t,num=None,il=0,style=None):
    ppr=''
    if style: ppr+=f'<w:pStyle w:val="{style}"/>'
    if num: ppr+=f'<w:numPr><w:ilvl w:val="{il}"/><w:numId w:val="{num}"/></w:numPr>'
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t>{t}</w:t></w:r></w:p>'
body=p('A restart limit')
for t,il in [('a',0),('b',1),('c',2),('d',2),('e',1),('f',2),('g',0),('h',1),('i',2)]: body+=p('item '+t,1,il)
body+=p('B style linked')
for t,il in [('a',0),('b',1),('c',0)]: body+=p('item '+t,3,il)
for t in 'de': body+=p('item style '+t,None,0,'ListPara')
for t in 'fg': body+=p('item direct '+t,2,0)
body+=p('item style h',None,0,'ListPara')
body+=p('C override level')
for t in 'abc': body+=p('item '+t,4,0)
body+=p('D bullet colour size')
for t,il in [('a',0),('b',1),('c',1),('d',0)]: body+=p('item '+t,5,il)
doc=f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}<w:sectPr/></w:body></w:document>'
ct='<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>'
rels='<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
drels='<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="n" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
with zipfile.ZipFile(sys.argv[1],'w',zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml',ct);z.writestr('_rels/.rels',rels);z.writestr('word/_rels/document.xml.rels',drels)
    z.writestr('word/document.xml',doc);z.writestr('word/numbering.xml',numbering);z.writestr('word/styles.xml',styles)
