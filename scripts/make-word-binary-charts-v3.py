# Makes test-data/word-charts-binary/word-charts-binary-v3.docx from the Word-made v1 .docx (scripts/make-word-binary-charts.ps1)
# by editing chart1.xml (the column chart) directly - no Excel automation: one point's data label, a linear trend line,
# fixed y error bars, the value axis as 0.0 in blue, a green legend, no category axis labels and a red title; and in chart6.xml
# (the scatter chart) custom plus-only x error bars and fixed both-way y error bars on the first series. Then
# scripts/save-word-binary-charts.ps1 lets Word save it as .doc / .rtf / PDF.
import re, sys, zipfile, os
folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'test-data', 'word-charts-binary')
source = os.path.join(folder, 'word-charts-binary-v1.docx')
target = os.path.join(folder, 'word-charts-binary-v3.docx')
with zipfile.ZipFile(source) as zin, zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == 'word/charts/chart1.xml':
            x = data.decode('utf-8')
            first = x.index('<c:ser>')
            cat = x.index('<c:cat>', first)
            extra = ('<c:dLbls><c:dLbl><c:idx val="1"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:showLegendKey val="0"/><c:showVal val="1"/>'
                     '<c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbl>'
                     '<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>'
                     '<c:trendline><c:trendlineType val="linear"/><c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline>'
                     '<c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:noEndCap val="0"/><c:val val="0.5"/></c:errBars>')
            x = x[:cat] + extra + x[cat:]
            val_ax = x.index('<c:valAx>')
            x = x[:val_ax] + x[val_ax:].replace('<c:numFmt formatCode="General" sourceLinked="1"/>', '<c:numFmt formatCode="0.0" sourceLinked="0"/>', 1)
            cat_ax = x.index('<c:catAx>')
            x = x[:cat_ax] + x[cat_ax:].replace('<c:tickLblPos val="nextTo"/>', '<c:tickLblPos val="none"/>', 1)
            title_end = x.index('</c:title>')
            head = x[:title_end]
            head = re.sub(r'<a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>',
                          '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>', head)
            head = head.replace('<a:rPr lang="en-US"/>', '<a:rPr lang="en-US"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr>')
            x = head + x[title_end:]
            val_ax = x.index('<c:valAx>')
            val_end = x.index('</c:valAx>')
            axis = x[val_ax:val_end]
            axis = re.sub(r'(<c:txPr>.*?<a:defRPr[^>]*>)<a:solidFill>.*?</a:solidFill>', r'\1<a:solidFill><a:srgbClr val="0000FF"/></a:solidFill>', axis, count=1, flags=re.S)
            x = x[:val_ax] + axis + x[val_end:]
            legend = x.index('<c:legend>')
            legend_end = x.index('</c:legend>')
            part = re.sub(r'(<c:txPr>.*?<a:defRPr[^>]*>)<a:solidFill>.*?</a:solidFill>', r'\1<a:solidFill><a:srgbClr val="00B050"/></a:solidFill>', x[legend:legend_end], count=1, flags=re.S)
            data = (x[:legend] + part + x[legend_end:]).encode('utf-8')
        if item.filename == 'word/charts/chart6.xml':
            x = data.decode('utf-8')
            first = x.index('<c:ser>')
            at = x.index('<c:xVal>', first)
            amounts = ''.join('<c:pt idx="%d"><c:v>%s</c:v></c:pt>' % (i, v) for i, v in enumerate(['0.2', '0.4', '0.6', '0.8']))
            bars = ('<c:errBars><c:errDir val="x"/><c:errBarType val="plus"/><c:errValType val="cust"/><c:noEndCap val="0"/>'
                    '<c:plus><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="4"/>' + amounts + '</c:numLit></c:plus></c:errBars>'
                    '<c:errBars><c:errDir val="y"/><c:errBarType val="both"/><c:errValType val="fixedVal"/><c:noEndCap val="0"/><c:val val="0.3"/></c:errBars>')
            data = (x[:at] + bars + x[at:]).encode('utf-8')
        zout.writestr(item, data)
print('written', os.path.basename(target))
