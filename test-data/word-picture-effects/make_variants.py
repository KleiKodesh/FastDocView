# Adds XML-only picture effect variants to Word's own save (effects-com-v1.docx) -> effects-xml-v1.docx
import zipfile,re,sys
src,dst=sys.argv[1],sys.argv[2]
z=zipfile.ZipFile(src); doc=z.read('word/document.xml').decode('utf8')
para=re.search(r'<w:p [^>]*>(?:(?!<w:p ).)*?<w:drawing>.*?</w:p>',doc,re.S).group(0)
V=[('duotone-blue','<a:duotone><a:prstClr val="black"/><a:srgbClr val="4F81BD"><a:tint val="45000"/><a:satMod val="400000"/></a:srgbClr></a:duotone>',''),
   ('duotone-red-yellow','<a:duotone><a:srgbClr val="C00000"/><a:srgbClr val="FFFF00"/></a:duotone>',''),
   ('alpha50','<a:alphaModFix amt="50000"/>',''),
   ('bilevel25','<a:biLevel thresh="25000"/>',''),
   ('bilevel75','<a:biLevel thresh="75000"/>',''),
   ('grayscale-lum','<a:grayscl/><a:lum bright="20000" contrast="40000"/>',''),
   ('clrchange-green-to-blue','<a:clrChange><a:clrFrom><a:srgbClr val="00C800"/></a:clrFrom><a:clrTo><a:srgbClr val="0000FF"/></a:clrTo></a:clrChange>',''),
   ('softedge20','','<a:effectLst><a:softEdge rad="254000"/></a:effectLst>'),
   ('rot3d-y40','','<a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="2400000" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>'),
   ('persp3d-x30','','<a:scene3d><a:camera prst="perspectiveRelaxed"><a:rot lat="1800000" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>')]
add=''; n=900
for label,blip,sppr in V:
    p=para.replace('>none<','>'+label+'<',1) if '>none<' in para else para
    p=re.sub(r'(<a:blip [^>]*?)(/?)>',lambda m:m.group(1)+'>'+blip+('</a:blip>' if m.group(2) else ''),p,count=1) if blip else p
    if sppr: p=p.replace('</a:prstGeom>','</a:prstGeom>'+sppr,1)
    n+=1; p=re.sub(r'(<wp:docPr id=")\d+',r'\g<1>%d'%n,p); p=re.sub(r'(<pic:cNvPr id=")\d+',r'\g<1>%d'%n,p)
    add+='<w:p><w:r><w:t>'+label+'</w:t></w:r></w:p>'+p
doc=doc.replace('<w:sectPr',add+'<w:sectPr',1)
o=zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED)
for i in z.infolist(): o.writestr(i, doc if i.filename=='word/document.xml' else z.read(i.filename))
o.close()
