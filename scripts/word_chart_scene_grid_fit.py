# Per-probe camera fit for the 3-D scene grid: Word 16 pinhole (axis through the plot centre, eye distance = plot height / 2 /
# tan(perspective / 4)), box at the parallel scale, free box offset (x, y, z). Probe(p).fit() / robust_fit(P); rms in pt.
import sys, json, math, numpy as np
sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.abspath(__file__)))
from word_chart_camera_fit import turn, residuals, lm, walls_for
THIN = 12.0  # faces thinner than this (pt) are matched by outline distance (edge-on floor / side wall)
CORN = np.array([(x,y,z) for x in (0,1) for y in (0,1) for z in (0,1)], float)
def dims(p, iw, ih):
    cats, ser = p.get('cats',4), p.get('series',1)
    base = math.floor((cats+ser)/2)/cats
    H = base*p['h']/100 if p.get('h') else base*ih/iw
    return np.array([1.0, H, ser*p['depth']/100/cats])
def unturn(q, p):
    a, b = math.radians(p['rotX']), math.radians(p['rotY'])
    x1, y2, z2 = q
    y = y2*math.cos(a) + z2*math.sin(a); z1 = -y2*math.sin(a) + z2*math.cos(a)
    x = x1*math.cos(b) + z1*math.sin(b); zz = -x1*math.sin(b) + z1*math.cos(b)
    return np.array([x, y, -zz])
class Probe:
    def __init__(s_, p, inset=(5.4,3.3)):
        s_.p = p
        x0,y0,x1,y1 = p['image']; s_.iw, s_.ih = x1-x0, y1-y0; s_.c = ((x0+x1)/2,(y0+y1)/2)
        s_.S = dims(p, s_.iw, s_.ih)
        t = turn(CORN*s_.S - s_.S/2, p['rotX'], p['rotY'])
        bw, bh = np.ptp(t[:,0]), np.ptp(t[:,1])
        a, b = (s_.iw-2*inset[0])/bw, (s_.ih-2*inset[1])/bh
        s_.s = min(a,b); s_.lim = 'W' if a < b else 'H'
        s_.E = s_.ih/2/math.tan(math.radians(p['perspective']/4)) if p['perspective'] else float('inf')
    def cam(s_, unit, o):  # camera coords in points, o = box-centre offset in scene units
        return turn(np.array(unit,float)*s_.S - s_.S/2 + o, s_.p['rotX'], s_.p['rotY'])*s_.s
    def proj(s_, q):
        f = 1.0 if not np.isfinite(s_.E) else s_.E/(s_.E - q[...,2])
        return np.stack([q[...,0]*f, q[...,1]*f], -1)
    def raw(s_, unit, o): return s_.proj(s_.cam(unit, o))
    def fit(s_):
        f = lambda o: residuals(s_.p, lambda u: s_.raw(u, o), sim=(1.0, s_.c[0], s_.c[1]), thin=THIN, faces=walls_for(s_.p['rotY']))[0]
        o, r = lm(f, [0,0,0]); s_.o = o; s_.rms = math.sqrt((r@r)/len(r)*2); return o
    def err(s_, o):
        r = residuals(s_.p, lambda u: s_.raw(u, o), sim=(1.0, s_.c[0], s_.c[1]), thin=THIN, faces=walls_for(s_.p['rotY']))[0]
        return math.sqrt((r@r)/len(r)*2), np.abs(r).max()
def robust_fit(P):
    best = None
    for st in ([0,0,0],[0.1,0,0],[-0.1,0,0],[0,0.1,0],[0,-0.1,0],[0,0,0.3],[0,0,-0.3],[-0.1,0.1,0.2],[-0.3,0,0],[-0.5,0,0]):
        f = lambda o: residuals(P.p, lambda u: P.raw(u, o), sim=(1.0, P.c[0], P.c[1]), thin=THIN, faces=walls_for(P.p['rotY']))[0]
        try:
            o, r = lm(f, st, iters=80)
        except Exception: continue
        c = r@r
        if best is None or c < best[0]: best = (c, o, r)
    P.o = best[1]; P.rms = math.sqrt(best[0]/len(best[2])*2); return P.o
