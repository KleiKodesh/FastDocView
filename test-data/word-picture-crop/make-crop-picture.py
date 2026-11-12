# A 200 x 100 px PNG at 96 dpi (pHYs 3780 px per metre: without it LibreOffice sizes it at the screen's dpi, so its fo:clip lengths
# would depend on the machine): quadrants red / green (top), blue / yellow (bottom), so a crop shows which part is kept.
import struct, zlib, sys
w, h = 200, 100
rows = []
for y in range(h):
    row = bytearray([0])
    for x in range(w):
        c = (255,0,0) if y < 50 and x < 100 else (0,160,0) if y < 50 else (0,0,255) if x < 100 else (255,220,0)
        row += bytes(c)
    rows.append(bytes(row))
def chunk(t, d): return struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
png = b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'pHYs', struct.pack('>IIB', 3780, 3780, 1)) + chunk(b'IDAT', zlib.compress(b''.join(rows))) + chunk(b'IEND', b'')
open(sys.argv[1], 'wb').write(png)
