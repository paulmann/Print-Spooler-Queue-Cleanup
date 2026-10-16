#!/usr/bin/env python3
"""Build res/app.ico (multi-size, BMP32 frames) and packaging/assets/* from assets/icon-1024.png.
Depends only on Pillow for resampling; the ICO container is written by hand for exact control."""
import struct
from PIL import Image
from pathlib import Path

root = Path(__file__).resolve().parent.parent
src = root / "assets" / "icon-1024.png"
ico = root / "res" / "app.ico"
store = root / "packaging" / "assets"
store.mkdir(parents=True, exist_ok=True)

img = Image.open(src).convert("RGBA")
SIZES = [16, 20, 24, 32, 40, 48, 64, 128, 256]


def bmp_frame(im: Image.Image) -> bytes:
    w = h = im.size[0]
    px = im.load()
    xor = bytearray()
    for y in range(h - 1, -1, -1):          # bottom-up
        for x in range(w):
            r, g, b, a = px[x, y]
            xor += bytes((b, g, r, a))       # BGRA
    stride = ((w + 31) // 32) * 4
    and_mask = bytearray(stride * h)         # 32bpp: mask ignored, keep zeroed
    hdr = struct.pack("<IiiHHIIiiII", 40, w, h * 2, 1, 32, 0, len(xor), 0, 0, 0, 0)
    return hdr + bytes(xor) + bytes(and_mask)


entries, blobs = [], bytearray()
offset = 6 + 16 * len(SIZES)
for s in SIZES:
    fr = bmp_frame(img.resize((s, s), Image.LANCZOS))
    entries.append(struct.pack("<BBBBHHII", s & 0xFF, s & 0xFF, 0, 0, 1, 32, len(fr), offset))
    blobs += fr
    offset += len(fr)
data = struct.pack("<HHH", 0, 1, len(SIZES)) + b"".join(entries) + bytes(blobs)
ico.write_bytes(data)
print("ico:", ico, len(data), "bytes,", len(SIZES), "frames")


def centered(canvas, icon_px, bg):
    im = Image.new("RGBA", canvas, bg + (255,))
    ic = img.resize((icon_px, icon_px), Image.LANCZOS)
    im.paste(ic, ((canvas[0] - icon_px) // 2, (canvas[1] - icon_px) // 2), ic)
    return im


BG = (37, 124, 214)
centered((44, 44), 44, BG).save(store / "Square44x44Logo.png")
centered((71, 71), 64, BG).save(store / "SmallTile71x71Logo.png")
centered((150, 150), 150, BG).save(store / "Square150x150Logo.png")
centered((310, 150), 120, BG).save(store / "Wide310x150Logo.png")
centered((310, 310), 260, BG).save(store / "LargeTile310x310Logo.png")
centered((50, 50), 50, BG).save(store / "StoreLogo.png")
sp = Image.new("RGBA", (620, 300), BG + (255,))
ic = img.resize((150, 150), Image.LANCZOS)
sp.paste(ic, (235, 75), ic)
sp.save(store / "SplashScreen620x300.png")
print("store assets:", sorted(p.name for p in store.iterdir()))
