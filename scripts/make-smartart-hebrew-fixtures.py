"""Writes scripts/smartart-hebrew-fixtures.json (ASCII, Hebrew as u-escapes) for make-smartart-layouts.ps1 -Fixtures: SmartArt
layouts with Hebrew text (with and without niqqud), mixed Hebrew / English, and theme fonts other than Calibri, so the layout
engine's font measuring is checked against Word's own drawing part. Run it, then
  powershell -File scripts/make-smartart-layouts.ps1 -Fixtures scripts/smartart-hebrew-fixtures.json
  python scripts/strip-smartart-drawing.py"""
import json
import pathlib

TORAH = 'תורה'
MISHNAH = 'משנה'
TALMUD = 'תלמוד'
SHULCHAN_ARUCH = 'שולחן ערוך'
MISHNAH_BERURAH = 'משנה ברורה'
HALACHAH = 'הלכה'
AGGADAH = 'אגדה'
RASHI = 'רש״י'
TOSAFOT = 'תוספות'
# with niqqud and cantillation-free points (combining marks)
BERESHIT = 'בְּרֵאשִׁית בָּרָא'
ELOKIM = 'אֱלֹהִים'
SHAMAYIM = 'הַשָּׁמַיִם וְאֵת הָאָרֶץ'

fixtures = [
    {'id': 'default', 'name': 'hebrew-block-list', 'texts': [TORAH, MISHNAH, TALMUD, SHULCHAN_ARUCH, MISHNAH_BERURAH]},
    {'id': 'default', 'name': 'hebrew-niqqud-block-list', 'texts': [BERESHIT, ELOKIM, SHAMAYIM]},
    {'id': 'vList2', 'name': 'hebrew-vertical-list', 'texts': [SHULCHAN_ARUCH + ' ' + HALACHAH, MISHNAH_BERURAH, AGGADAH]},
    {'id': 'process1', 'name': 'hebrew-process', 'texts': [TORAH, MISHNAH + ' ' + TALMUD, HALACHAH]},
    {'id': 'hList1', 'name': 'hebrew-horizontal-list', 'texts': [TORAH, MISHNAH, TALMUD, HALACHAH, AGGADAH]},
    {'id': 'default', 'name': 'mixed-block-list', 'texts': ['Rashi ' + RASHI, 'Tosafot ' + TOSAFOT, TALMUD + ' Bavli', 'Mishnah Berurah']},
    {'id': 'default', 'name': 'times-block-list', 'minorLatin': 'Times New Roman', 'texts': ['Alpha', 'Beta Wave', 'Gamma', 'Delta', 'Epsilon']},
    {'id': 'default', 'name': 'hebrew-david-block-list', 'minorComplex': 'David', 'texts': [TORAH, MISHNAH, TALMUD, SHULCHAN_ARUCH, MISHNAH_BERURAH]},
]
target = pathlib.Path(__file__).resolve().parent / 'smartart-hebrew-fixtures.json'
target.write_text(json.dumps(fixtures, ensure_ascii=True, indent=1) + '\n', encoding='ascii')
print('wrote', target.name, len(fixtures), 'fixtures')
