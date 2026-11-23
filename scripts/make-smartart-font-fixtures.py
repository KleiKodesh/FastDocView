"""Writes scripts/smartart-font-fixtures.json (ASCII; non-Latin text as u-escapes) for make-smartart-layouts.ps1 -Fixtures:
SmartArt block lists that check the font measuring against Word's own drawing part - bold / italic / bold italic faces, a
font that kerns only through GPOS (Bahnschrift; Hadassah Friedlaender for Hebrew), Arabic (positional forms and the lam-alef
ligature, GSUB), Japanese / Chinese text in the default theme and Korean in a theme East Asian font. Run it, then
  powershell -File scripts/make-smartart-layouts.ps1 -Fixtures scripts/smartart-font-fixtures.json
  python scripts/strip-smartart-drawing.py"""
import json
import pathlib

# one long word per node: the size is set by the word's width, so a few percent of width moves it a whole point
LATIN = ['Extraordinary', 'Internationalization', 'Characteristically']
# Arabic: al-mustashfayat, al-muhadarat, wa-al-istiqlal (positional forms, lam-alef)
ARABIC_LONG = ['\u0627\u0644\u0645\u0633\u062a\u0634\u0641\u064a\u0627\u062a', '\u0627\u0644\u0645\u062d\u0627\u0636\u0631\u0627\u062a', '\u0648\u0627\u0644\u0627\u0633\u062a\u0642\u0644\u0627\u0644']
# Hebrew letter pairs Hadassah Friedlaender kerns both ways (gimel-yod, gimel-zayin, yod-nun: 7 % of their width) in one long word each
HEBREW_KERNED = ['\u05d2\u05d9' * 6, '\u05d2\u05d6' * 5, '\u05d9\u05e0' * 6]
HEBREW_LONG = ['\u05d4\u05ea\u05e9\u05d5\u05d1\u05d5\u05ea\u05d9\u05d4\u05dd', '\u05d5\u05db\u05e9\u05d4\u05ea\u05e4\u05dc\u05dc\u05d5', '\u05d4\u05de\u05e9\u05e0\u05d9\u05d5\u05ea\u05d9\u05d4\u05dd']
SALAM = '\u0627\u0644\u0633\u0644\u0627\u0645 \u0639\u0644\u064a\u0643\u0645'
MARHABA = '\u0645\u0631\u062d\u0628\u0627'
KITAB = '\u0643\u062a\u0627\u0628'
LUGHA = '\u0627\u0644\u0644\u063a\u0629 \u0627\u0644\u0639\u0631\u0628\u064a\u0629'
TORAH = '\u05ea\u05d5\u05e8\u05d4'
MISHNAH = '\u05de\u05e9\u05e0\u05d4'
TALMUD = '\u05ea\u05dc\u05de\u05d5\u05d3'
SHULCHAN_ARUCH = '\u05e9\u05d5\u05dc\u05d7\u05df \u05e2\u05e8\u05d5\u05da'
NIHONGO = '\u65e5\u672c\u8a9e'
HIRAGANA = '\u3072\u3089\u304c\u306a'
KATAKANA_KANJI = '\u30ab\u30bf\u30ab\u30ca \u6f22\u5b57'
TOKYO = 'Tokyo \u6771\u4eac'
ZHONGWEN = '\u4e2d\u6587'
BEIJING = '\u5317\u4eac\u5e02'
HANZI = '\u6c49\u5b57\u6587\u672c'
HANGUK = '\ud55c\uad6d\uc5b4'
SEOUL = '\uc11c\uc6b8'
# several lines of East Asian text: the size is set by the line height of the East Asian font
NIHON_BUN = '\u65e5\u672c\u8a9e\u306e\u6587\u7ae0\u3092\u66f8\u304f\u7df4\u7fd2'
KANKOKU_BUN = '\ub300\ud55c\ubbfc\uad6d \uc11c\uc6b8\ud2b9\ubcc4\uc2dc \uc885\ub85c\uad6c'
CHUUGOKU_BUN = '\u4e2d\u534e\u4eba\u6c11\u5171\u548c\u56fd\u5317\u4eac\u5e02\u4e1c\u57ce\u533a'
ANNYEONG = '\uc548\ub155\ud558\uc138\uc694'

fixtures = [
    {'id': 'default', 'name': 'bold-block-list', 'bold': True, 'texts': LATIN},
    {'id': 'default', 'name': 'italic-block-list', 'italic': True, 'minorLatin': 'Times New Roman', 'texts': LATIN},
    {'id': 'default', 'name': 'bold-italic-block-list', 'bold': True, 'italic': True, 'minorLatin': 'Times New Roman', 'texts': LATIN},
    {'id': 'default', 'name': 'bold-hebrew-block-list', 'bold': True, 'texts': HEBREW_LONG},
    {'id': 'default', 'name': 'bahnschrift-kern-block-list', 'minorLatin': 'Bahnschrift', 'texts': ['AVAVAVAVAVAV', 'ToToToToToTo', 'PAPAPAPAPA']},
    {'id': 'default', 'name': 'hadassah-hebrew-block-list', 'complexFont': 'Hadassah Friedlaender', 'texts': HEBREW_KERNED},
    {'id': 'default', 'name': 'arabic-block-list', 'texts': ARABIC_LONG},
    {'id': 'default', 'name': 'arabic-words-block-list', 'texts': [SALAM, MARHABA, KITAB, LUGHA]},
    {'id': 'default', 'name': 'japanese-block-list', 'texts': [NIHONGO, HIRAGANA, KATAKANA_KANJI, TOKYO]},
    {'id': 'default', 'name': 'chinese-block-list', 'texts': [ZHONGWEN, BEIJING, HANZI]},
    {'id': 'default', 'name': 'yugothic-lines-block-list', 'minorEastAsian': 'Yu Gothic', 'texts': [NIHON_BUN, NIHON_BUN[:6], NIHON_BUN[3:]]},
    {'id': 'default', 'name': 'msgothic-lines-block-list', 'minorEastAsian': 'MS Gothic', 'texts': [NIHON_BUN, NIHON_BUN[:6], NIHON_BUN[3:]]},
    {'id': 'default', 'name': 'malgun-lines-block-list', 'minorEastAsian': 'Malgun Gothic', 'texts': [KANKOKU_BUN, KANKOKU_BUN[:4], KANKOKU_BUN[5:]]},
    {'id': 'default', 'name': 'simsun-lines-block-list', 'minorEastAsian': 'SimSun', 'texts': [CHUUGOKU_BUN, CHUUGOKU_BUN[:5], CHUUGOKU_BUN[4:]]},
    {'id': 'default', 'name': 'korean-ea-block-list', 'minorEastAsian': 'Malgun Gothic', 'texts': [HANGUK, SEOUL, ANNYEONG]},
]
target = pathlib.Path(__file__).resolve().parent / 'smartart-font-fixtures.json'
target.write_text(json.dumps(fixtures, ensure_ascii=True, indent=1) + '\n', encoding='ascii')
print('wrote', target.name, len(fixtures), 'fixtures')
