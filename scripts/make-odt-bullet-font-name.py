"""Packs test-data/odt-bullet-fonts/bullet-font-name-v1.fodt as bullet-font-name-v1.odt by hand (LibreOffice's own save
rewrites the bullet's style:font-name as fo:font-family, so it cannot make this case): styles.xml carries the font-face-decls
and the named list style (read by the styles.xml byte filter), content.xml the automatic list style and the body.
Run: python scripts/make-odt-bullet-font-name.py"""
import os
import re
import zipfile

folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'test-data', 'odt-bullet-fonts')
flat = open(os.path.join(folder, 'bullet-font-name-v1.fodt'), encoding='utf-8').read()
namespaces = re.search(r'<office:document (xmlns[^>]*?) office:version', flat, re.S).group(1)


def section(name):
    return re.search(r'<office:%s>.*?</office:%s>' % (name, name), flat, re.S).group(0)


fonts = section('font-face-decls')
styles = ('<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles %s office:version="1.3">%s%s</office:document-styles>'
          % (namespaces, fonts, section('styles')))
content = ('<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content %s office:version="1.3">%s%s%s</office:document-content>'
           % (namespaces, fonts, section('automatic-styles'), section('body')))
manifest = ('<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">'
            '<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>'
            '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
            '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>')
with zipfile.ZipFile(os.path.join(folder, 'bullet-font-name-v1.odt'), 'w') as package:
    package.writestr(zipfile.ZipInfo('mimetype'), 'application/vnd.oasis.opendocument.text', zipfile.ZIP_STORED)
    package.writestr('META-INF/manifest.xml', manifest, zipfile.ZIP_DEFLATED)
    package.writestr('styles.xml', styles, zipfile.ZIP_DEFLATED)
    package.writestr('content.xml', content, zipfile.ZIP_DEFLATED)
