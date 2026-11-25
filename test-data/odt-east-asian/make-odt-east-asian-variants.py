# Builds odt-east-asian-variants-v1.odt: East Asian size / bold / italic (style:font-size-asian, -weight-asian, -style-asian)
# beside the Western and complex-script ones. Reference: LibreOffice's PDF of it (odt-east-asian-variants-v1.pdf,
# soffice --headless --convert-to pdf with a private profile; LibreOffice 25: P1 Han 20pt, P2 Han 18pt not bold, P3 Han 12pt
# (fo:font-size does not reach the Asian size), P4 Han 20pt bold italic in a right-to-left paragraph).
import zipfile
HAN = '漢字'          # two Han characters
HEB = 'שלום'  # a Hebrew word
content = f'''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3">
<office:font-face-decls>
<style:font-face style:name="Arial" svg:font-family="Arial" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"/>
<style:font-face style:name="MS Mincho" svg:font-family="'MS Mincho'" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"/>
</office:font-face-decls>
<office:automatic-styles>
<style:style style:name="PL" style:family="paragraph"><style:paragraph-properties style:writing-mode="lr-tb"/><style:text-properties style:font-name="Arial" fo:font-size="12pt" style:font-name-asian="MS Mincho" style:font-size-asian="12pt" style:font-name-complex="Arial" style:font-size-complex="12pt"/></style:style>
<style:style style:name="PR" style:family="paragraph" style:parent-style-name="PL"><style:paragraph-properties style:writing-mode="rl-tb"/></style:style>
<style:style style:name="T1" style:family="text"><style:text-properties style:font-size-asian="20pt" style:font-weight-asian="bold" style:font-style-asian="italic"/></style:style>
<style:style style:name="T2" style:family="text"><style:text-properties fo:font-weight="bold" style:font-weight-asian="normal" style:font-size-asian="18pt"/></style:style>
<style:style style:name="T3" style:family="text"><style:text-properties fo:font-size="16pt" fo:font-style="italic"/></style:style>
</office:automatic-styles>
<office:body><office:text>
<text:p text:style-name="PL"><text:span text:style-name="T1">P1 Ab {HAN} Z</text:span></text:p>
<text:p text:style-name="PL"><text:span text:style-name="T2">P2 Ab {HAN} Z</text:span></text:p>
<text:p text:style-name="PL"><text:span text:style-name="T3">P3 Ab {HAN} Z</text:span></text:p>
<text:p text:style-name="PR"><text:span text:style-name="T1">P4 {HEB} {HAN} Z</text:span></text:p>
</office:text></office:body></office:document-content>'''
manifest = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">
<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>'''
with zipfile.ZipFile('odt-east-asian-variants-v1.odt', 'w') as z:
    z.writestr(zipfile.ZipInfo('mimetype'), 'application/vnd.oasis.opendocument.text', zipfile.ZIP_STORED)
    z.writestr('META-INF/manifest.xml', manifest, zipfile.ZIP_DEFLATED)
    z.writestr('content.xml', content.encode('utf-8'), zipfile.ZIP_DEFLATED)
