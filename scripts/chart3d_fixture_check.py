# 3-D chart fixtures vs Word: renders each fixture with the test runner's --chart-html and compares every chart's texts with
# Word's PDF (compare-chart-text-positions.py, positions relative to the chart title). Prints per chart median / max |dx|, |dy|.
# Usage: python chart3d_fixture_check.py <runner exe> <out folder> [fixture.docx ...]  (default: the 3-D Word chart fixtures)
import sys, os, re, subprocess
here = os.path.dirname(os.path.abspath(__file__)); root = os.path.join(here, '..')
exe, out = sys.argv[1], sys.argv[2]; os.makedirs(out, exist_ok=True)
B = os.path.join(root, 'test-data', 'word-charts-binary')
fixtures = sys.argv[3:] or [os.path.join(B, f) for f in ('word-charts-3d-rows-v1.docx', 'word-charts-right-angled-v1.docx', 'word-charts-seraxis-v1.docx',
                                                       'word-charts-right-angled-v1.doc', 'word-charts-seraxis-v1.doc', 'word-charts-right-angled-v1.rtf', 'word-charts-seraxis-v1.rtf',
                                                       'word-charts-binary-v3.docx', 'word-charts-binary-v3.doc', 'word-charts-binary-v1.docx')]
for fx in fixtures:
    pdf = os.path.splitext(fx)[0] + '.pdf'
    if not os.path.exists(fx) or not os.path.exists(pdf): continue
    html = os.path.join(out, os.path.basename(fx) + '.html')
    subprocess.run([exe, '--chart-html', fx, html], check=True, stdout=subprocess.DEVNULL)
    text = open(html, encoding='utf-8').read()
    titles = [m.group(1).strip() for m in re.finditer(r'<svg class="word-chart".*?<text[^>]*>([^<]*)</text>', text, re.S)]
    for t in titles:
        if not t: continue
        r = subprocess.run([sys.executable, os.path.join(here, 'compare-chart-text-positions.py'), pdf, html, t], capture_output=True, text=True).stdout
        m = re.search(r'median dx ([0-9.]+) pt  dy ([0-9.]+) pt  max dx ([0-9.]+) dy ([0-9.]+)', r)
        if m: print('%-34s %-24s med %s / %s  max %s / %s' % (os.path.basename(fx), t.encode('ascii', 'backslashreplace').decode()[:24], *m.groups()))
