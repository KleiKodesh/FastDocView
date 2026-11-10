"""Writes list-bullets-nolisttext-v1.rtf: Word 16's list-bullets-v1.rtf with every flat label group (the listtext and pntext groups) removed, so a reader must count each label and draw it with the level's own character formatting.
Word 16 re-saving the stripped file writes the same label groups again (checked through COM)."""
import sys

BS = chr(92)
source = open(sys.argv[1] if len(sys.argv) > 1 else "list-bullets-v1.rtf", "rb").read()
starts = [b"{" + BS.encode() + b"listtext", b"{" + BS.encode() + b"pntext"]
out = bytearray()
i = 0
while i < len(source):
    if any(source.startswith(s, i) and not source[i + len(s):i + len(s) + 1].isalpha() for s in starts):
        depth = 0
        while i < len(source):
            c = source[i]
            if c == 0x5C: i += 2; continue        # an escaped character never opens or closes a group
            if c == 0x7B: depth += 1
            elif c == 0x7D:
                depth -= 1
                if depth == 0: i += 1; break
            i += 1
        continue
    out.append(source[i]); i += 1
open(sys.argv[2] if len(sys.argv) > 2 else "list-bullets-nolisttext-v1.rtf", "wb").write(out)
