# Per-type check of Word's 3-D scene model on test-data/word-charts-scene-types (make-word-charts-scene-types.py):
# perspective probes - Word 16 camera (eye distance = plot height / 2 / tan(perspective / 4)), parallel-fit scale, free box offset,
# scene size per type (series rows: depth = series x depthPercent / categories, height = floor((categories + series) / 2) / categories
# x plot height / width; one row (clustered / stacked): depth = depthPercent / categories, height = plot height / width);
# right-angled probes - oblique view fitted free (scale, offset, depth step across / up) and compared with sin(rotY), sin(rotX).
# Usage: python word_chart_scene_types_fit.py <measured types.json>
import sys, json, math, numpy as np
sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.abspath(__file__)))
import word_chart_scene_grid_fit as G
from word_chart_camera_fit import residuals, lm, walls_for
ROWS = ('col-standard', 'area-standard', 'line-standard', 'surface', 'line-ra')
def dims(p, iw, ih, free=(1.0, 1.0)):
    cats, ser = 4, p.get('series', 1)
    rows = p['type'] in ROWS
    if rows:
        base = math.floor((cats + ser) / 2) / cats
        H = base * p['h'] / 100 if p.get('h') else base * ih / iw
        D = ser * p['depth'] / 100 / cats
    else:
        H = ih / iw if not p.get('h') else p['h'] / 100
        D = p['depth'] / 100 / cats * (H if p['type'].startswith('bar-') else 1.0)   # horizontal bars: the slot runs up the height
    return np.array([1.0, H * free[0], D * free[1]])
def perspective_fit(p, free_dims=False):
    G.dims = lambda q, iw, ih: dims(q, iw, ih)
    P = G.Probe(p); G.robust_fit(P); rms = P.rms
    if not free_dims or rms < 1: return rms, (rms, (1.0, 1.0))
    best = (rms, (1.0, 1.0))
    for fh in (0.8, 0.9, 1.0, 1.1, 1.25):
        for fd in (0.6, 0.8, 1.0, 1.25, 1.6, 2.0):
            G.dims = lambda q, iw, ih, fh=fh, fd=fd: dims(q, iw, ih, (fh, fd))
            Q = G.Probe(p); G.robust_fit(Q)
            if Q.rms < best[0]: best = (Q.rms, (fh, fd))
    return rms, best
def oblique_fit(p):
    """Right-angled view: the front face flat (width s, height s x H), each unit of depth stepped kx across and ky up. Word crops
    the scene picture of a right-angled chart to the drawn faces, so the picture's aspect says nothing about the plot: H is free."""
    x0, y0, x1, y1 = p['image']
    S = dims(p, 1.0, 1.0)
    def raw(v):
        s, tx, ty, kx, ky, H = v
        def f(u):
            u = np.array(u, float) * np.array([1.0, H, S[2]])
            return np.stack([(u[..., 0] + u[..., 2] * kx) * s + tx, (u[..., 1] + u[..., 2] * ky) * s + ty], -1)
        return f
    a, b = math.radians(p['rotX']), math.radians(p['rotY'])
    best = None
    for H0 in (0.3, 0.55, 0.8):
        st = [(x1 - x0) * 0.85, -(x1 - x0) * 0.45, -(y1 - y0) * 0.4, math.sin(b), math.sin(a), H0]
        g = lambda v: residuals(p, raw(v), sim=(1.0, (x0 + x1) / 2, (y0 + y1) / 2), thin=G.THIN, faces=walls_for(0))[0]
        try: v, r = lm(g, st, iters=150)
        except Exception: continue
        if best is None or r @ r < best[0]: best = (r @ r, v, r)
    return math.sqrt(best[0] / len(best[2]) * 2), best[1]
if __name__ == '__main__':
    D = json.load(open(sys.argv[1]))
    by = {}
    for p in D:
        if p['type'].startswith('bar-'):
            # horizontal bars: Word paints the floor's fill on the x = 0 face (under the bars' bases) and the side wall's on the
            # y = 0 face, so the measured colours name the faces the other way round
            f = p['faces']; f['floor'], f['side'] = f.get('side'), f.get('floor'); p['faces'] = {k: v for k, v in f.items() if v}
        if not p['faces']: by.setdefault(p['type'], []).append(('noface', p['name'])); continue
        if p['rAngAx']:
            rms, v = oblique_fit(p)
            s, tx, ty, kx, ky, H = v
            front = (s, s * H)
            print('%-26s rms %.2f front %.1f x %.1f  across/sinY %.3f up/sinX %s' % (p['name'], rms, front[0], front[1],
                  kx / math.sin(math.radians(p['rotY'])) if p['rotY'] else float('nan'), '%.3f' % (ky / math.sin(math.radians(p['rotX']))) if p['rotX'] else 'nan'))
            by.setdefault(p['type'], []).append(rms)
        else:
            rms, best = perspective_fit(p, free_dims=True)
            print('%-26s rms %.2f free-size best %.2f at height x%.2f depth x%.2f' % (p['name'], rms, best[0], best[1][0], best[1][1]))
            by.setdefault(p['type'], []).append(rms)
    for t, r in by.items():
        num = [x for x in r if not isinstance(x, tuple)]
        print('TYPE %-14s n %d  within 1 pt %d  median %.2f  max %.2f  no faces %d' % (t, len(num), sum(x < 1 for x in num), np.median(num) if num else -1, max(num) if num else -1, len(r) - len(num)))
