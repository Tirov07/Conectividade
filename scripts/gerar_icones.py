#!/usr/bin/env python3
"""Gera os ícones da extensão a partir do ICON_CANAINFO.png.

Problema: a folha de bordo ficou TRANSPARENTE (vazada), igual ao fundo fora do
círculo. Queremos (opção B): folha BRANCA + fundo externo TRANSPARENTE.

Solução: flood fill a partir das bordas marca a transparência "externa" (fundo).
Toda transparência que NÃO é alcançável pela borda está cercada de vermelho =
é a folha -> pintamos de branco opaco. O fundo externo permanece transparente.

Depois redimensiona (média de blocos) para 16/48/128 e grava PNGs RGBA reais.
Só usa a biblioteca padrão (struct + zlib).
"""
import struct, zlib, os
from collections import deque

AQUI = os.path.dirname(os.path.abspath(__file__))
ICONS = os.path.join(AQUI, "..", "extension", "icons")
SRC = os.path.join(ICONS, "ICON_CANAINFO.png")


# ---- Leitura de PNG RGBA (stdlib) -----------------------------------------
def ler_png_rgba(caminho):
    data = open(caminho, "rb").read()
    assert data[:8] == b"\x89PNG\r\n\x1a\n"
    pos, idat = 8, bytearray()
    w = h = None
    while pos < len(data):
        length = struct.unpack_from(">I", data, pos)[0]
        tipo = data[pos+4:pos+8]
        chunk = data[pos+8:pos+8+length]
        if tipo == b"IHDR":
            w, h, depth, ct = struct.unpack_from(">IIBB", chunk, 0)
            assert depth == 8 and ct == 6, "esperava RGBA 8-bit"
        elif tipo == b"IDAT":
            idat += chunk
        elif tipo == b"IEND":
            break
        pos += 12 + length
    raw = zlib.decompress(bytes(idat))
    stride = w * 4
    out = bytearray()
    prev = bytearray(stride)
    p = 0

    def paeth(a, b, c):
        pp = a + b - c
        pa, pb, pc = abs(pp-a), abs(pp-b), abs(pp-c)
        return a if (pa <= pb and pa <= pc) else (b if pb <= pc else c)

    for _ in range(h):
        f = raw[p]; p += 1
        line = bytearray(raw[p:p+stride]); p += stride
        for i in range(stride):
            a = line[i-4] if i >= 4 else 0
            b = prev[i]
            c = prev[i-4] if i >= 4 else 0
            x = line[i]
            if f == 1: line[i] = (x + a) & 0xFF
            elif f == 2: line[i] = (x + b) & 0xFF
            elif f == 3: line[i] = (x + ((a + b) >> 1)) & 0xFF
            elif f == 4: line[i] = (x + paeth(a, b, c)) & 0xFF
        out += line
        prev = line
    return w, h, out


# ---- Pinta a folha (transparência interna) de branco ----------------------
def folha_para_branco(w, h, buf, limiar_alpha=40):
    """Flood fill a partir das bordas: marca a transparência EXTERNA.
    O que sobra de transparente (interno = folha) vira branco opaco."""
    externo = bytearray(w * h)  # 1 = fundo externo (transparente a manter)
    fila = deque()

    def transp(x, y):
        return buf[(y*w + x)*4 + 3] < limiar_alpha

    # semeia nas 4 bordas
    for x in range(w):
        for y in (0, h-1):
            if transp(x, y) and not externo[y*w + x]:
                externo[y*w + x] = 1
                fila.append((x, y))
    for y in range(h):
        for x in (0, w-1):
            if transp(x, y) and not externo[y*w + x]:
                externo[y*w + x] = 1
                fila.append((x, y))

    while fila:
        x, y = fila.popleft()
        for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
            nx, ny = x+dx, y+dy
            if 0 <= nx < w and 0 <= ny < h and not externo[ny*w + nx] and transp(nx, ny):
                externo[ny*w + nx] = 1
                fila.append((nx, ny))

    # Toda transparência NÃO-externa (cercada de vermelho = folha) -> branco.
    for y in range(h):
        for x in range(w):
            i = (y*w + x)*4
            if buf[i+3] < limiar_alpha and not externo[y*w + x]:
                buf[i] = buf[i+1] = buf[i+2] = 255
                buf[i+3] = 255
    return buf


# ---- Redimensiona (média de área) preservando alfa ------------------------
def redimensionar(w, h, buf, novo):
    out = bytearray(novo * novo * 4)
    for ny in range(novo):
        y0 = ny * h // novo
        y1 = max(y0 + 1, (ny + 1) * h // novo)
        for nx in range(novo):
            x0 = nx * w // novo
            x1 = max(x0 + 1, (nx + 1) * w // novo)
            sr = sg = sb = sa = cnt = 0
            for yy in range(y0, y1):
                for xx in range(x0, x1):
                    i = (yy*w + xx)*4
                    a = buf[i+3]
                    # premultiplica pelo alfa para a média de cor ficar correta
                    sr += buf[i]*a; sg += buf[i+1]*a; sb += buf[i+2]*a
                    sa += a; cnt += 1
            o = (ny*novo + nx)*4
            if sa > 0:
                out[o] = min(255, sr // sa)
                out[o+1] = min(255, sg // sa)
                out[o+2] = min(255, sb // sa)
                out[o+3] = sa // cnt
            else:
                out[o] = out[o+1] = out[o+2] = out[o+3] = 0
    return out


# ---- Grava PNG RGBA --------------------------------------------------------
def gravar_png(caminho, w, h, buf):
    raw = bytearray()
    stride = w * 4
    for y in range(h):
        raw.append(0)
        raw += buf[y*stride:(y+1)*stride]

    def chunk(tipo, dados):
        c = struct.pack(">I", len(dados)) + tipo + dados
        return c + struct.pack(">I", zlib.crc32(tipo + dados) & 0xFFFFFFFF)

    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0)  # RGBA
    idat = zlib.compress(bytes(raw), 9)
    with open(caminho, "wb") as f:
        f.write(sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", idat) + chunk(b"IEND", b""))


def main():
    w, h, buf = ler_png_rgba(SRC)
    print(f"original {w}x{h}")
    buf = folha_para_branco(w, h, buf)
    print("folha pintada de branco; fundo externo mantido transparente")
    for tam in (16, 48, 128):
        px = redimensionar(w, h, buf, tam)
        dest = os.path.join(ICONS, f"icon{tam}.png")
        gravar_png(dest, tam, tam, px)
        print(f"gerado {dest}")


if __name__ == "__main__":
    main()
