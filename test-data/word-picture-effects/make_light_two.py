# Light rig on a picture turned on both axes (orthographic camera, lat x lon grid) -> effects-light-two-v1.docx; Word prints it to
# effects-light-two-v1.pdf (save-light-two.ps1). Usage: python make_light_two.py effects-light-ortho-v1.docx effects-light-two-v1.docx
import zipfile,re,sys
src,dst=sys.argv[1],sys.argv[2]
z=zipfile.ZipFile(src); doc=z.read('word/document.xml').decode('utf8')
para=re.search(r'<w:p [^>]*>(?:(?!<w:p ).)*?<w:drawing>.*?</w:p>',doc,re.S).group(0)
para=re.sub(r'<a:scene3d>.*?</a:scene3d>','',para,flags=re.S)  # the source picture's own camera
def cam(prst,lat=0,lon=0,rev=0,fov=None):
    f=' fov="%d"'%fov if fov is not None else ''
    return '<a:scene3d><a:camera prst="%s"%s><a:rot lat="%d" lon="%d" rev="%d"/></a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>'%(prst,f,lat*60000,lon*60000,rev*60000)
V=[('two-0-0',1,'',cam('orthographicFront',0,0))]
for la in (-45,-30,-15,15,30,45):
    for lo in (-45,-30,-15,15,30,45):
        V.append(('two-lat%d-lon%d'%(la,lo),1,'',cam('orthographicFront',la%360,lo%360)))
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
