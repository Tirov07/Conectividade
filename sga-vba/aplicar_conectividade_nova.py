#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
aplicar_conectividade_nova.py

Gera versoes ajustadas dos modulos VBA de integracao SGA <-> Inmetro para
funcionar com a NOVA extensao Chrome (Conectividade Nova), que usa as MESMAS
pastas C:\\Sga\\Req e C:\\Sga\\Resp.

O que o ajuste faz em cada modulo .bas:

  1. Neutraliza a validacao do app antigo:
        If ofuncoes.ConectividadeIsAlive("Conectividade") = False Then
           ... (MsgBox / Add_Message_Status_Operacao + Exit Function/Sub)
        End If
     -> todo o bloco e comentado (prefixado com ') e nao bloqueia mais.

  2. Neutraliza a ativacao da janela do app antigo:
        AppActivate ("Conectividade - " & RetornaLoginInMetro)
     -> a linha e comentada.

O que NAO e tocado (de proposito):
  - AppActivate "SGa"                  (traz o SGA de volta: continua valido)
  - Shell "C:\\SGA\\CanainfoPrint.exe"  (outro app, impressao de CRLV)
  - Gravacao em Req / leitura de Resp   (a extensao usa as mesmas pastas)
  - ExisteDiretorio(CaminhoSGaRemote)   (checagem da nuvem, continua valida)

Uso:
    python aplicar_conectividade_nova.py <pasta_origem> <pasta_destino>
"""

import sys
import os
import re

MARCA = "'[CONECTIVIDADE_NOVA]"  # comentario-marcador para rastrear o que foi alterado

# Modulos que fazem a integracao com o Inmetro (os unicos que precisam de ajuste)
MODULOS_ALVO = [
    "Integracao_SGA_Inmetro_Atualizar_Dados.bas",
    "Integracao_SGA_Inmetro_Cadastrar_Veiculo.bas",
    "Integracao_SGA_Inmetro_Consulta_Selagem_Completa.bas",
    "Integracao_SGA_Inmetro_Consulta_Selagem_Completa_Lote_Placas_NEw.bas",
    "Integracao_SGA_Inmetro_Consulta_Selagem_Substuicao.bas",
    "Integracao_SGA_Inmetro_CRLV.bas",
    "Integracao_SGA_Inmetro_Download_Certificado.bas",
    "Integracao_SGA_Inmetro_Emitir_GRU.bas",
    "Integracao_SGA_Inmetro_Informar_Ensaio.bas",
    "Integracao_SGA_Inmetro_Informar_Selagem.bas",
    "Integracao_SGA_Inmetro_Informar_Substituicao.bas",
]


def eh_linha_comentada(linha):
    return linha.lstrip().startswith("'")


def comentar(linha):
    """Preserva a indentacao e prefixa a linha com aspas simples."""
    if eh_linha_comentada(linha):
        return linha  # ja comentada
    espacos = len(linha) - len(linha.lstrip(" \t"))
    return linha[:espacos] + "'" + linha[espacos:]


def processar_arquivo(texto):
    """
    Recebe o texto do .bas (lista de linhas, SEM alterar quebras) e devolve
    o texto ajustado + contadores (blocos_isalive, appactivate).
    """
    # Preserva o terminador original (os .bas exportados usam CRLF)
    usa_crlf = "\r\n" in texto
    linhas = texto.split("\r\n") if usa_crlf else texto.split("\n")

    saida = []
    i = 0
    cont_isalive = 0
    cont_appactivate = 0

    # Regex dos dois padroes
    re_isalive = re.compile(r"ConectividadeIsAlive\s*\(\s*[\"']Conectividade[\"']\s*\)\s*=\s*False\s+Then",
                            re.IGNORECASE)
    re_appactivate = re.compile(r"^\s*AppActivate\s*\(\s*[\"']Conectividade\s*-",
                                re.IGNORECASE)

    while i < len(linhas):
        linha = linhas[i]

        # --- Padrao 1: bloco If ... ConectividadeIsAlive(...) = False Then ... End If ---
        if (not eh_linha_comentada(linha)) and re_isalive.search(linha):
            # Captura o bloco ate o End If correspondente (nivel 1, sem If aninhado aqui)
            bloco = [linha]
            j = i + 1
            while j < len(linhas):
                bloco.append(linhas[j])
                if re.match(r"^\s*End\s+If\b", linhas[j], re.IGNORECASE):
                    break
                j += 1

            saida.append(MARCA + " validacao da Conectividade antiga desativada:")
            for bl in bloco:
                saida.append(comentar(bl))
            cont_isalive += 1
            i = j + 1
            continue

        # --- Padrao 2: AppActivate ("Conectividade - ...") ---
        if (not eh_linha_comentada(linha)) and re_appactivate.search(linha):
            saida.append(MARCA + " ativacao da janela do app antigo desativada:")
            saida.append(comentar(linha))
            cont_appactivate += 1
            i += 1
            continue

        saida.append(linha)
        i += 1

    terminador = "\r\n" if usa_crlf else "\n"
    return terminador.join(saida), cont_isalive, cont_appactivate


def main():
    if len(sys.argv) != 3:
        print("Uso: python aplicar_conectividade_nova.py <pasta_origem> <pasta_destino>")
        sys.exit(1)

    origem, destino = sys.argv[1], sys.argv[2]
    os.makedirs(destino, exist_ok=True)

    total_isalive = 0
    total_appactivate = 0

    print("Modulo".ljust(60), "IsAlive", "AppActivate")
    print("-" * 85)
    for nome in MODULOS_ALVO:
        caminho_in = os.path.join(origem, nome)
        if not os.path.isfile(caminho_in):
            print(f"[AVISO] nao encontrado: {nome}")
            continue
        # Os .bas do VBA sao salvos em cp1252 (Windows Latin-1)
        with open(caminho_in, "r", encoding="cp1252", errors="replace") as f:
            texto = f.read()

        novo, c1, c2 = processar_arquivo(texto)

        caminho_out = os.path.join(destino, nome)
        with open(caminho_out, "w", encoding="cp1252", errors="replace", newline="") as f:
            f.write(novo)

        total_isalive += c1
        total_appactivate += c2
        print(nome.ljust(60), str(c1).center(7), str(c2).center(11))

    print("-" * 85)
    print("TOTAL".ljust(60), str(total_isalive).center(7), str(total_appactivate).center(11))
    print(f"\nArquivos ajustados gravados em: {destino}")


if __name__ == "__main__":
    main()
