# Word's band colours on the ovals of shapes-3d-bevels-oval-v1 (make-shapes-3d-bevels-oval.ps1): per preset, 48 steps of the
# 24 pt bevel inward along the outline's normal at 8 directions (0, 45, ... 315 degrees of the ellipse's parameter), each the mean
# of 3 pixels along the outline (-1, 0, 1 pt) at 8 px per pt.
#   python sample_bevels_oval.py shapes-3d-bevels-oval-v1-docx-word.pdf shapes-3d-bevels-oval-v1-word-samples.txt
import fitz, sys, math
pdf = sys.argv[1]; out = sys.argv[2]
Z = 8
pix = fitz.open(pdf)[0].get_pixmap(matrix=fitz.Matrix(Z, Z))
W = pix.width; S = pix.samples
def px(x, y):
    i = (int(y * Z) * W + int(x * Z)) * pix.n
    return S[i], S[i + 1], S[i + 2]
names = ['relaxedInset','circle','slope','cross','angle','softRound','convex','coolSlant','divot','riblet','hardEdge','artDeco']
with open(out, 'w') as f:
    for i, name in enumerate(names):
        cx = 50 + (i % 3) * 180 + 75; cy = 70 + (i // 3) * 170 + 55
        for k in range(48):
            d = (k + 0.5) / 2
            cols = []
            for j in range(8):
                t = j * math.pi / 4
                ex, ey = cx + 75 * math.cos(t), cy + 55 * math.sin(t)
                nx, ny = -math.cos(t) / 75, -math.sin(t) / 55
                l = math.hypot(nx, ny); nx /= l; ny /= l
                pts = [(ex + nx * d - ny * o, ey + ny * d + nx * o) for o in (-1, 0, 1)]
                c = [px(x, y) for x, y in pts]
                cols.append(tuple(round(sum(v[q] for v in c) / 3) for q in range(3)))
            f.write('%s %.2f %s\n' % (name, d / 24, ' '.join('%02x%02x%02x' % c for c in cols)))
