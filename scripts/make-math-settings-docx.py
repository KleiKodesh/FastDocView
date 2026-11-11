"""Writes test-data/word-math/math-settings-v2.docx: Word 16's math-v2.docx with settings.xml's m:mathPr changed to brkBin
repeat, brkBinSub "-+" and wrapRight, and the 40-term equation lengthened to 50 operators (y1..y10) for the timing (the settings the page finishes once the lines are known).
Run: python scripts/make-math-settings-docx.py test-data/word-math"""
import sys, zipfile, os, re
folder = sys.argv[1]
source = zipfile.ZipFile(os.path.join(folder, "math-v2.docx"))
with zipfile.ZipFile(os.path.join(folder, "math-settings-v2.docx"), "w", zipfile.ZIP_DEFLATED) as out:
    for item in source.infolist():
        data = source.read(item.filename)
        if item.filename == "word/settings.xml":
            text = data.decode("utf-8")
            text = text.replace('<m:brkBin m:val="before"/>', '<m:brkBin m:val="repeat"/>').replace('<m:brkBinSub m:val="--"/>', '<m:brkBinSub m:val="-+"/>')
            text = text.replace('<m:wrapIndent m:val="1440"/>', '<m:wrapIndent m:val="1440"/><m:wrapRight/>')
            assert 'm:val="repeat"' in text and "wrapRight" in text
            data = text.encode("utf-8")
        if item.filename == "word/document.xml":
            text = data.decode("utf-8")
            marker = "longeq</m:t></m:r>"
            at = text.index(marker) + len(marker)
            eq = text.index("</m:r>", at) + len("</m:r>")      # the "=" run
            text = text[:eq] + "<m:r><m:t>" + "+".join(f"y{i}" for i in range(1, 11)) + "+</m:t></m:r>" + text[eq:]
            data = text.encode("utf-8")
        out.writestr(item, data)
