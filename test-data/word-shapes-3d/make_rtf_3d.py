# shapes-3d-v1.rtf (Word's own save: the 3-D styles kept, f3D off) with f3D turned on in every shape -> shapes-3d-on-v1.rtf;
# make-shapes-3d-fixture.ps1 -Step 2 prints it with Word (shapes-3d-on-v1-rtf-word.pdf)
import re
src = open('shapes-3d-v1.rtf', 'rb').read()
BS = b'\x5c'
pattern = re.escape(b'{' + BS + b'sp{' + BS + b'sn shapeType}{' + BS + b'sv ') + rb'\d+' + re.escape(b'}}')
extra = b'{' + BS + b'sp{' + BS + b'sn f3D}{' + BS + b'sv 1}}'
out = re.sub(pattern, lambda m: m.group(0) + extra, src)
open('shapes-3d-on-v1.rtf', 'wb').write(out)
print(out.count(b'sn f3D'))
