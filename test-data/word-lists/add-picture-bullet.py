# Step of make-list-bullets-fixture.ps1: Word 16's ListLevel.ApplyPictureBullet does nothing through COM, so the levels whose
# bullet is "o" in Courier New get Word's own picture-bullet markup (w:numPicBullet with a VML v:imagedata, w:lvlPicBulletId,
# ECMA-376 17.9.21 / 17.9.10) and a 40 x 20 px PNG; Word then reopens the file and saves it as .docx / .doc / .rtf.
import sys, zipfile, struct, zlib, re
src, dst = sys.argv[1], sys.argv[2]
def png(w, h):
    raw = b''.join(b'\x00' + b'\x00\x80\xff' * w for _ in range(h))
    chunk = lambda t, d: struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b'')
zin = zipfile.ZipFile(src)
zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED)
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == 'word/numbering.xml':
        s = data.decode('utf8')
        pic = ('<w:numPicBullet w:numPicBulletId="0"><w:pict><v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:30pt;height:15pt" o:bullet="t">'
               '<v:imagedata r:id="rIdPictureBullet1" o:title=""/></v:shape></w:pict></w:numPicBullet>')
        s = re.sub(r'(<w:numbering [^>]*>)', lambda m: m.group(1) + pic, s, count=1)
        s = re.sub(r'<w:lvlText w:val="o"/>((?:(?!</w:lvl>).)*?)<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"[^/]*/>',
                   lambda m: '<w:lvlText w:val="o"/><w:lvlPicBulletId w:val="0"/>' + m.group(1).replace('<w:rFonts', ''), s, flags=re.S)
        s = s.replace('<w:lvlText w:val="o"/><w:lvlPicBulletId w:val="0"/>', '<w:lvlText w:val=""/><w:lvlPicBulletId w:val="0"/>')
        data = s.encode('utf8')
    elif item.filename == '[Content_Types].xml' and b'Extension="png"' not in data:
        data = data.replace(b'<Default ', b'<Default Extension="png" ContentType="image/png"/><Default ', 1)
    zout.writestr(item, data)
zout.writestr('word/_rels/numbering.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rIdPictureBullet1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>')
zout.writestr('word/media/image1.png', png(40, 20))
zout.close()
