# Soft edge / 3-D camera / duotone variants on Word's own save -> effects-3d-v1.docx (Word prints it to effects-3d-v1.pdf)
import zipfile,re,sys
src,dst=sys.argv[1],sys.argv[2]
z=zipfile.ZipFile(src); doc=z.read('word/document.xml').decode('utf8')
para=re.search(r'<w:p [^>]*>(?:(?!<w:p ).)*?<w:drawing>.*?</w:p>',doc,re.S).group(0)
def cam(prst,lat=0,lon=0,rev=0,fov=None):
    f=' fov="%d"'%fov if fov is not None else ''
    return '<a:scene3d><a:camera prst="%s"%s><a:rot lat="%d" lon="%d" rev="%d"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>'%(prst,f,lat*60000,lon*60000,rev*60000)
V=[('soft5',1,'','<a:effectLst><a:softEdge rad="63500"/></a:effectLst>'),
   ('soft10',1,'','<a:effectLst><a:softEdge rad="127000"/></a:effectLst>'),
   ('soft20',1,'','<a:effectLst><a:softEdge rad="254000"/></a:effectLst>'),
   ('soft10-half',0.5,'','<a:effectLst><a:softEdge rad="127000"/></a:effectLst>'),
   ('duotone-blue','1','<a:duotone><a:prstClr val="black"/><a:srgbClr val="4F81BD"><a:tint val="45000"/><a:satMod val="400000"/></a:srgbClr></a:duotone>',''),
   ('duotone-sat','1','<a:duotone><a:srgbClr val="202020"/><a:srgbClr val="C0504D"><a:satMod val="300000"/></a:srgbClr></a:duotone>',''),
   ('relaxed-lat30',1,'',cam('perspectiveRelaxed',30)),
   ('relaxed-lat30-half',0.5,'',cam('perspectiveRelaxed',30)),
   ('front-lon30',1,'',cam('perspectiveFront',0,30)),
   ('front-lat330',1,'',cam('perspectiveFront',330)),
   ('front-lon30-fov90',1,'',cam('perspectiveFront',0,30,0,5400000)),
   ('ortho-lon40',1,'',cam('orthographicFront',0,40)),
   ('sq-lat30',(2/3,1),'',cam('perspectiveFront',30)),
   ('sq-lon30',(2/3,1),'',cam('perspectiveFront',0,30)),
   ('tall-lat30',(0.5,1.5),'',cam('perspectiveFront',30)),
   ('tall-lon30',(0.5,1.5),'',cam('perspectiveFront',0,30))]
add=''; n=950
for label,scale,blip,sppr in V:
    p=para
    if blip: p=re.sub(r'(<a:blip [^>]*?)(/?)>',lambda m:m.group(1)+'>'+blip+('</a:blip>' if m.group(2) else ''),p,count=1)
    if sppr: p=p.replace('</a:prstGeom>','</a:prstGeom>'+sppr,1)
    sx,sy=scale if isinstance(scale,tuple) else (float(scale),float(scale))
    p=re.sub(r'(cx|cy)="(\d+)"',lambda m:'%s="%d"'%(m.group(1),int(int(m.group(2))*(sx if m.group(1)=='cx' else sy))),p)
    n+=1; p=re.sub(r'(<wp:docPr id=")\d+',r'\g<1>%d'%n,p); p=re.sub(r'(<pic:cNvPr id=")\d+',r'\g<1>%d'%n,p)
    add+='<w:p><w:r><w:t>'+label+'</w:t></w:r></w:p>'+p
body_start=doc.index('<w:body>')+len('<w:body>'); sect=doc.index('<w:sectPr')
doc=doc[:body_start]+add+doc[sect:]
o=zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED)
for i in z.infolist(): o.writestr(i, doc if i.filename=='word/document.xml' else z.read(i.filename))
o.close()
