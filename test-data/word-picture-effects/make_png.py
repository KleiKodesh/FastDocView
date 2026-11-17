import zlib,struct,sys
w,h=240,160
rows=b''
for y in range(h):
    r=bytearray([0])
    for x in range(w):
        q=(x//60)+4*(y//80)
        cols=[(255,0,0),(0,200,0),(0,0,255),(255,255,0),(0,255,255),(255,0,255),(255,255,255),(0,0,0)]
        c=cols[q]; t=x%60/59.0
        r+=bytes(int(v*(0.4+0.6*t)) for v in c)
    rows+=bytes(r)
def ch(t,d): return struct.pack('>I',len(d))+t+d+struct.pack('>I',zlib.crc32(t+d)&0xffffffff)
open(sys.argv[1],'wb').write(b'\x89PNG\r\n\x1a\n'+ch(b'IHDR',struct.pack('>IIBBBBB',w,h,8,2,0,0,0))+ch(b'IDAT',zlib.compress(rows))+ch(b'IEND',b''))
