# Builds test-data/odt-math/odt-math-v1.odt (+ .fodt, + LibreOffice's .pdf of it) with LibreOffice itself (UNO): one
# Writer formula object (Math, MathML in "Object N/content.xml"; inline in the .fodt) per paragraph.
# Run with LibreOffice's python while soffice listens:  soffice --headless --accept="pipe,name=fdvmath;urp;" (SAL_DISABLE_PRINTERLIST=1 SAL_DISABLE_DEFAULTPRINTER=1)
import sys, time, uno
from com.sun.star.beans import PropertyValue

out = sys.argv[1]          # path without extension
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(60):
    try:
        ctx = resolver.resolve("uno:pipe,name=fdvmath;urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (hidden,))
text = doc.Text
cursor = text.createTextCursor()

FORMULAS = [
    ("fraction", "a over b"),
    ("root", "sqrt{x} + nroot{3}{y}"),
    ("scripts", "x^2 + x_i + x_i^2"),
    ("sum", "sum from{i=1} to{n} i"),
    ("integral", "int from 0 to 1 f(x) dx"),
    ("brackets", "left( a + b right) cdot left[ c right]"),
    ("matrix", "left( matrix{1 # 0 ## 0 # 1} right)"),
    ("function", "sin(%theta) + lim from{n toward infinity} a_n"),
    ("styles", "bold B + ital I + color red r + \"text here\""),
    ("accents", "overline{z} + hat{v} + tilde{a} + vec{u}"),
    ("binomial", "binom{n}{k}"),
]
for name, formula in FORMULAS:
    text.insertString(cursor, "formula " + name + ": ", False)
    obj = doc.createInstance("com.sun.star.text.TextEmbeddedObject")
    obj.CLSID = "078B7ABA-54FC-457F-8551-6147e776a997"
    obj.AnchorType = uno.Enum("com.sun.star.text.TextContentAnchorType", "AS_CHARACTER")
    text.insertTextContent(cursor, obj, False)
    obj.getEmbeddedObject().Formula = formula
    text.insertString(cursor, " end", False)
    text.insertControlCharacter(cursor, 0, False)   # PARAGRAPH_BREAK
text.insertString(cursor, "last paragraph", False)

def store(path, filter_name):
    f = PropertyValue(); f.Name = "FilterName"; f.Value = filter_name
    doc.storeToURL(uno.systemPathToFileUrl(path), (f,))
store(out + ".odt", "writer8")
store(out + ".fodt", "OpenDocument Text Flat XML")
store(out + ".pdf", "writer_pdf_Export")
doc.close(True)
