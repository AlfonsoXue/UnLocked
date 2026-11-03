"""產生 UnLocked.ico（打開的鎖頭圖示）。不需要額外套件。

用法：python build/make_icon.py
"""
import os
import struct
import zlib

HERE = os.path.dirname(os.path.abspath(__file__))

BODY = (37, 99, 235)       # 鎖身：藍色
BODY_DARK = (29, 78, 216)  # 鎖身下緣
SHACKLE = (100, 116, 139)  # 鎖環：灰藍色
HOLE = (255, 255, 255)     # 鑰匙孔


def rounded_rect(x, y, x0, y0, x1, y1, r):
    if x < x0 or x > x1 or y < y0 or y > y1:
        return False
    cx = min(max(x, x0 + r), x1 - r)
    cy = min(max(y, y0 + r), y1 - r)
    return (x - cx) ** 2 + (y - cy) ** 2 <= r * r


def sample(x, y):
    """回傳座標 (0~1) 的 RGBA；依圖層由上到下判斷。"""
    # 鑰匙孔
    if (x - 0.5) ** 2 + (y - 0.64) ** 2 <= 0.065 ** 2 or (0.465 <= x <= 0.535 and 0.64 <= y <= 0.80):
        if rounded_rect(x, y, 0.16, 0.44, 0.84, 0.94, 0.09):
            return HOLE + (255,)
    # 鎖身
    if rounded_rect(x, y, 0.16, 0.44, 0.84, 0.94, 0.09):
        return (BODY_DARK if y > 0.86 else BODY) + (255,)
    # 打開的鎖環：上半圓環 + 右腳伸到鎖身、左腳懸空
    cx, cy, ro, ri = 0.60, 0.30, 0.235, 0.135
    d2 = (x - cx) ** 2 + (y - cy) ** 2
    if y <= cy and ri * ri <= d2 <= ro * ro:
        return SHACKLE + (255,)
    if cx + ri <= x <= cx + ro and cy <= y <= 0.47:
        return SHACKLE + (255,)
    if cx - ro <= x <= cx - ri and cy <= y <= 0.36:
        return SHACKLE + (255,)
    return (0, 0, 0, 0)


def render(size, ss=6):
    """以超取樣做反鋸齒，回傳由上到下的 RGBA 列。"""
    rows = []
    for py in range(size):
        row = []
        for px in range(size):
            acc = [0, 0, 0, 0]
            for sy in range(ss):
                for sx in range(ss):
                    r, g, b, a = sample((px + (sx + 0.5) / ss) / size, (py + (sy + 0.5) / ss) / size)
                    acc[0] += r * a
                    acc[1] += g * a
                    acc[2] += b * a
                    acc[3] += a
            n = ss * ss
            alpha = acc[3] / n
            if acc[3]:
                row.append((round(acc[0] / acc[3]), round(acc[1] / acc[3]), round(acc[2] / acc[3]), round(alpha)))
            else:
                row.append((0, 0, 0, 0))
        rows.append(row)
    return rows


def bmp_entry(size):
    rows = render(size)
    header = struct.pack("<IiiHHIIiiII", 40, size, size * 2, 1, 32, 0, 0, 0, 0, 0, 0)
    xor = b"".join(bytes((b, g, r, a)) for row in reversed(rows) for (r, g, b, a) in row)
    mask_row = ((size + 31) // 32) * 4
    return header + xor + b"\0" * (mask_row * size)


def png_entry(size):
    rows = render(size, ss=3)
    raw = b"".join(b"\0" + bytes(c for px in row for c in px) for row in rows)

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)) +
            chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b""))


def main():
    images = [(s, bmp_entry(s)) for s in (16, 24, 32, 48, 64)] + [(256, png_entry(256))]
    out = struct.pack("<HHH", 0, 1, len(images))
    offset = 6 + 16 * len(images)
    body = b""
    for size, data in images:
        dim = 0 if size >= 256 else size
        out += struct.pack("<BBBBHHII", dim, dim, 0, 0, 1, 32, len(data), offset + len(body))
        body += data
    path = os.path.join(HERE, "UnLocked.ico")
    with open(path, "wb") as f:
        f.write(out + body)
    print("wrote", path)


if __name__ == "__main__":
    main()
