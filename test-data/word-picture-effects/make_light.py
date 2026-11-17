# Soft edge / 3-D camera / duotone variants on Word's own save -> effects-light-v1.docx (Word prints it to effects-light-v1.pdf)
import zipfile,re,sys
src,dst=sys.argv[1],sys.argv[2]
z=zipfile.ZipFile(src); doc=z.read('word/document.xml').decode('utf8')
para=re.search(r'<w:p [^>]*>(?:(?!<w:p ).)*?<w:drawing>.*?</w:p>',doc,re.S).group(0)
def cam(prst,lat=0,lon=0,rev=0,fov=None):
    f=' fov="%d"'%fov if fov is not None else ''
    return '<a:scene3d><a:camera prst="%s"%s><a:rot lat="%d" lon="%d" rev="%d"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>'%(prst,f,lat*60000,lon*60000,rev*60000)
V=[('extra-0',1,'',cam('orthographicFront',0,0)),
   ('extra-lat-5',1,'',cam('orthographicFront',355,0)),
   ('extra-lat5',1,'',cam('orthographicFront',5,0)),
   ('extra-lon-5',1,'',cam('orthographicFront',0,355)),
   ('extra-lon5',1,'',cam('orthographicFront',0,5)),
   ('extra-lat30-lon30',1,'',cam('orthographicFront',30,30)),
   ('extra-lat-30-lon-30',1,'',cam('orthographicFront',330,330)),
   ('extra-lat-60',1,'',cam('orthographicFront',300,0)),
   ('extra-lat60',1,'',cam('orthographicFront',60,0)),
   ('extra-lon60',1,'',cam('orthographicFront',0,60)),
   ('extra-lon-60',1,'',cam('orthographicFront',0,300)),
   ('persp-lat-45',1,'',cam('perspectiveFront',315,0)),
   ('persp-lat-30',1,'',cam('perspectiveFront',330,0)),
   ('persp-lat-15',1,'',cam('perspectiveFront',345,0)),
   ('persp-lat15',1,'',cam('perspectiveFront',15,0)),
   ('persp-lat30',1,'',cam('perspectiveFront',30,0)),
   ('persp-lat45',1,'',cam('perspectiveFront',45,0)),
   ('persp-lon-45',1,'',cam('perspectiveFront',0,315)),
   ('persp-lon-30',1,'',cam('perspectiveFront',0,330)),
   ('persp-lon-15',1,'',cam('perspectiveFront',0,345)),
   ('persp-lon15',1,'',cam('perspectiveFront',0,15)),
   ('persp-lon30',1,'',cam('perspectiveFront',0,30)),
   ('persp-lon45',1,'',cam('perspectiveFront',0,45)),
   ('ortho-lat-45',1,'',cam('orthographicFront',315,0)),
   ('ortho-lat-30',1,'',cam('orthographicFront',330,0)),
   ('ortho-lat-15',1,'',cam('orthographicFront',345,0)),
   ('ortho-lat15',1,'',cam('orthographicFront',15,0)),
   ('ortho-lat30',1,'',cam('orthographicFront',30,0)),
   ('ortho-lat45',1,'',cam('orthographicFront',45,0)),
   ('ortho-lon-45',1,'',cam('orthographicFront',0,315)),
   ('ortho-lon-30',1,'',cam('orthographicFront',0,330)),
   ('ortho-lon-15',1,'',cam('orthographicFront',0,345)),
   ('ortho-lon15',1,'',cam('orthographicFront',0,15)),
   ('ortho-lon30',1,'',cam('orthographicFront',0,30)),
   ('ortho-lon45',1,'',cam('orthographicFront',0,45)),
   ('soft10-60x40',(1/3,1/3),'','<a:effectLst><a:softEdge rad="127000"/></a:effectLst>'),
   ('soft20-60x40',(1/3,1/3),'','<a:effectLst><a:softEdge rad="254000"/></a:effectLst>'),
   ('soft20-40x40',(2/9,1/3),'','<a:effectLst><a:softEdge rad="254000"/></a:effectLst>'),
   ('soft40-180x120',1,'','<a:effectLst><a:softEdge rad="508000"/></a:effectLst>')]
import os
sel=os.environ.get('SEL')
if sel: V=[v for v in V if sel in v[0]]
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
