#!/usr/bin/env python3
"""Analisa o ICON_CANAINFO.png: lê via stdlib (zlib) e amostra pixels para
entender o estado da folha (transparente? branca?) e do fundo."""
import struct, zlib, os

AQUI = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(AQUI, "..", "extension", "icons", "ICON_CANAINFO.png")


def ler_png_rgba(caminho):
    data = open(caminho, "rb").read()
    assert data[:8] == b"\x89PNG\r\n\x1a\n", "nao e PNG"
    pos = 8
    width = height = bit_depth = color_type = None
    idat = bytearray()
    while pos < len(data):
        length = struct.unpack_from(">I", data, pos)[0]
        tipo = data[pos+4:pos+8]
        chunk = data[pos+8:pos+8+length]
        if tipo == b"IHDR":
            width, height, bit_depth, color_type = struct.unpack_from(">IIBB", chunk, 0)
        elif tipo == b"IDAT":
            idat += chunk
        elif tipo == b"IEND":
            break
        pos += 12 + length
    assert bit_depth == 8 and color_type == 6, f"esperava 8bit RGBA, veio depth={bit_depth} ct={color_type}"
    raw = zlib.decompress(bytes(idat))
    # desfiltra (filtros PNG por linha)
    stride = width * 4
    out = bytearray()
    prev = bytearray(stride)
    p = 0
    def paeth(a, b, c):
        pp = a + b - c
        pa, pb, pc = abs(pp-a), abs(pp-b), abs(pp-c)
        return a if (pa <= pb and pa <= pc) else (b if pb <= pc else c)
    for _y in range(height):
        f = raw[p]; p += 1
        line = bytearray(raw[p:p+stride]); p += stride
        for i in range(stride):
            a = line[i-4] if i >= 4 else 0
            b = prev[i]
            c = prev[i-4] if i >= 4 else 0
            x = line[i]
            if f == 1: line[i] = (x + a) & 0xFF
            elif f == 2: line[i] = (x + b) & 0xFF
            elif f == 3: line[i] = (x + ((a+b) >> 1)) & 0xFF
            elif f == 4: line[i] = (x + paeth(a, b, c)) & 0xFF
        out += line
        prev = line
    return width, height, out


def px(buf, w, x, y):
    i = (y*w + x) * 4
    return buf[i], buf[i+1], buf[i+2], buf[i+3]


def main():
    w, h, buf = ler_png_rgba(SRC)
    print(f"{w}x{h} RGBA")
    pontos = {
        "canto (0,0) fora do circulo": (2, 2),
        "centro (folha)": (w//2, h//2),
        "meio-cima listrado": (w//2, h//6),
        "meio-baixo solido": (w//2, int(h*0.8)),
    }
    for nome, (x, y) in pontos.items():
        print(f"  {nome}: RGBA={px(buf, w, x, y)}")
    # Conta quantos pixels sao "quase brancos" (folha) e quantos transparentes
    branco = transp = vermelho = 0
    for y in range(0, h, 3):
        for x in range(0, w, 3):
            r, g, b, a = px(buf, w, x, y)
            if a < 30: transp += 1
            elif r > 200 and g > 200 and b > 200: branco += 1
            elif r > 150 and g < 100 and b < 100: vermelho += 1
    print(f"  amostra: transparentes={transp} brancos={branco} vermelhos={vermelho}")


if __name__ == "__main__":
    main()
