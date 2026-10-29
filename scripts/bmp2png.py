#!/usr/bin/env python3
"""Converte o CANAINFO.png (que na verdade é um BMP 50x50 4bpp) em PNGs reais
nos tamanhos 16, 48 e 128, usando só a biblioteca padrão (struct + zlib)."""
import struct, zlib, os

AQUI = os.path.dirname(os.path.abspath(__file__))
ICONS = os.path.join(AQUI, "..", "extension", "icons")
SRC = os.path.join(ICONS, "CANAINFO.png")


def ler_bmp(caminho):
    data = open(caminho, "rb").read()
    off = struct.unpack_from("<I", data, 10)[0]
    w = struct.unpack_from("<i", data, 18)[0]
    h = struct.unpack_from("<i", data, 22)[0]
    bpp = struct.unpack_from("<H", data, 28)[0]
    assert bpp == 4, f"esperava 4bpp, veio {bpp}"
    pal_off = 54
    n_cores = (off - pal_off) // 4
    palette = []
    for i in range(n_cores):
        b, g, r, a = data[pal_off + i * 4: pal_off + i * 4 + 4]
        palette.append((r, g, b))
    row_bytes = ((w * 4 + 31) // 32) * 4
    pixels = [[(0, 0, 0)] * w for _ in range(h)]
    for row in range(h):
        base = off + row * row_bytes
        y = h - 1 - row
        for x in range(w):
            byte = data[base + (x // 2)]
            idx = (byte >> 4) if (x % 2 == 0) else (byte & 0x0F)
            pixels[y][x] = palette[idx]
    return w, h, pixels


def redimensionar(pixels, w, h, novo):
    out = [[(255, 255, 255)] * novo for _ in range(novo)]
    for ny in range(novo):
        sy = min(h - 1, ny * h // novo)
        for nx in range(novo):
            sx = min(w - 1, nx * w // novo)
            out[ny][nx] = pixels[sy][sx]
    return out


def gravar_png(caminho, pixels, n):
    raw = bytearray()
    for y in range(n):
        raw.append(0)
        for x in range(n):
            r, g, b = pixels[y][x]
            raw += bytes((r, g, b))

    def chunk(tipo, dados):
        c = struct.pack(">I", len(dados)) + tipo + dados
        c += struct.pack(">I", zlib.crc32(tipo + dados) & 0xFFFFFFFF)
        return c

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", n, n, 8, 2, 0, 0, 0)
    idat = zlib.compress(bytes(raw), 9)
    with open(caminho, "wb") as f:
        f.write(sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b""))


def main():
    w, h, pixels = ler_bmp(SRC)
    print(f"BMP lido: {w}x{h}")
    for tam in (16, 48, 128):
        px = pixels if tam == w else redimensionar(pixels, w, h, tam)
        dest = os.path.join(ICONS, f"icon{tam}.png")
        gravar_png(dest, px, tam)
        print(f"gerado {dest}")


if __name__ == "__main__":
    main()
