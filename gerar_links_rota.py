"""
Gerador de links do Google Maps para uma planilha de rota de entregas.

O QUE O SCRIPT FAZ
-------------------
Le uma planilha .xlsx com as colunas:
    Rota, NF, Destinatario, Endereço, N°, Bairro, Cidade, Vol, KG, Ent, CEP, Seq, Motorista

Para cada parada, na ordem da coluna "Seq":
    1. Consulta o CEP no ViaCEP (base oficial dos Correios) para saber qual
       e o endereco oficial daquele CEP.
    2. Compara o endereco oficial com o que esta escrito na planilha
       (rua e cidade), ignorando acentos, maiusculas/minusculas e
       abreviacoes (RUA/R, AVENIDA/AV etc).
    3. Marca a parada como:
         OK               -> CEP confere com o endereco da planilha
         DIVERGENTE       -> CEP existe, mas o endereco nao bate
         NAO ENCONTRADO   -> CEP invalido/inexistente ou falha ao consultar
    4. Monta um link do Google Maps para o endereco (nao depende do ViaCEP,
       entao e gerado mesmo quando o status e NAO ENCONTRADO).

O resultado e impresso no terminal, na ordem da sequencia de entrega, e
salvo em um novo arquivo "..._com_links.xlsx" com duas colunas extras:
"Status Verificacao" e "Link Maps" (com hyperlink clicavel no Excel).

USO
---
    pip install pandas openpyxl requests
    python gerar_links_rota.py "Rota_190_-_LUCAS_ERB9I42__-_QUA.xlsx"

Se nenhum arquivo for informado, o script tenta usar o primeiro .xlsx
encontrado na pasta atual.
"""

import sys
import time
import glob
import unicodedata
import difflib
from urllib.parse import quote_plus

import requests
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
''
VIACEP_URL = "https://viacep.com.br/ws/{cep}/json/"
MAPS_URL = "https://www.google.com/maps/search/?api=1&query={query}"
LIMIAR_SIMILARIDADE = 0.6  # 0 a 1 - quao parecido o texto precisa ser para "bater"

PREFIXOS_LOGRADOURO = {
    "RUA", "R", "AV", "AVENIDA", "AL", "ALAMEDA", "ESTRADA", "EST",
    "ROD", "RODOVIA", "TRAVESSA", "TV", "PRACA", "PC", "LARGO",
    "VIA", "VIELA", "LADEIRA",
}


def normalizar(texto):
    """Maiusculas, sem acento, sem pontuacao e sem prefixo de logradouro."""
    if texto is None:
        return ""
    texto = str(texto).strip().upper()
    if texto == "NAN":
        return ""
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    texto = "".join(c if c.isalnum() or c.isspace() else " " for c in texto)
    palavras = [p for p in texto.split() if p not in PREFIXOS_LOGRADOURO]
    return " ".join(palavras)


def similaridade(a, b):
    """Retorna de 0 a 1 o quanto dois textos de endereco se parecem."""
    a_norm, b_norm = normalizar(a), normalizar(b)
    if not a_norm or not b_norm:
        return 0.0
    if a_norm in b_norm or b_norm in a_norm:
        return 1.0
    return difflib.SequenceMatcher(None, a_norm, b_norm).ratio()


def consultar_cep(cep, tentativas=2):
    """Consulta o ViaCEP. Retorna dict com os dados oficiais ou None se falhar."""
    cep_limpo = "".join(ch for ch in str(cep) if ch.isdigit())
    if len(cep_limpo) != 8:
        return None

    for tentativa in range(tentativas):
        try:
            resp = requests.get(VIACEP_URL.format(cep=cep_limpo), timeout=8)
            dados = resp.json()
            return None if dados.get("erro") else dados
        except (requests.RequestException, ValueError):
            if tentativa < tentativas - 1:
                time.sleep(1)
    return None


def montar_link_maps(endereco, numero, bairro, cidade, cep):
    """Monta um link de busca do Google Maps. Nao depende de API key nem do ViaCEP."""
    partes = [
        str(p).strip() for p in [endereco, numero, bairro, cidade, "SP", cep]
        if p is not None and str(p).strip().lower() != "nan" and str(p).strip() != ""
    ]
    query = quote_plus(", ".join(partes))
    return MAPS_URL.format(query=query)


def verificar_parada(endereco, bairro, cidade, cep):
    """Consulta o CEP e devolve o status de conferencia da parada."""
    dados = consultar_cep(cep)
    if dados is None:
        return "NAO ENCONTRADO"

    logradouro_oficial = dados.get("logradouro", "")
    cidade_oficial = dados.get("localidade", "")

    if logradouro_oficial:
        return "OK" if similaridade(endereco, logradouro_oficial) >= LIMIAR_SIMILARIDADE else "DIVERGENTE"

    # Alguns CEPs sao de "area" (sem logradouro especifico) - confere so a cidade
    return "OK (CEP de area)" if similaridade(cidade, cidade_oficial) >= LIMIAR_SIMILARIDADE else "DIVERGENTE"


def processar(caminho):
    df = pd.read_excel(caminho, dtype=str)

    colunas_esperadas = {"Endereço", "N°", "Bairro", "Cidade", "CEP", "Seq", "Destinatario"}
    faltando = colunas_esperadas - set(df.columns)
    if faltando:
        raise ValueError(f"Colunas faltando na planilha: {faltando}")

    df["Seq"] = pd.to_numeric(df["Seq"], errors="coerce")
    df = df.sort_values("Seq", kind="stable").reset_index(drop=True)

    status_lista, link_lista = [], []
    print(f"Conferindo {len(df)} paradas (ordem da rota)...\n")

    for _, linha in df.iterrows():
        status = verificar_parada(linha["Endereço"], linha["Bairro"], linha["Cidade"], linha["CEP"])
        link = montar_link_maps(linha["Endereço"], linha["N°"], linha["Bairro"], linha["Cidade"], linha["CEP"])

        status_lista.append(status)
        link_lista.append(link)

        seq_str = "?" if pd.isna(linha["Seq"]) else f"{int(linha['Seq']):>2}"
        print(f"[Seq {seq_str}] {linha['Destinatario']} -- {status}")
        print(f"           {link}\n")

        time.sleep(0.3)  # nao martelar o ViaCEP

    df["Status Verificacao"] = status_lista
    df["Link Maps"] = link_lista

    saida = caminho.rsplit(".", 1)[0] + "_com_links.xlsx"
    df.to_excel(saida, index=False)

    wb = load_workbook(saida)
    ws = wb.active
    col_link = df.columns.get_loc("Link Maps") + 1
    for linha_idx in range(2, len(df) + 2):
        celula = ws.cell(row=linha_idx, column=col_link)
        if celula.value:
            celula.hyperlink = celula.value
            celula.font = Font(color="0563C1", underline="single")
    wb.save(saida)

    nao_encontradas = status_lista.count("NAO ENCONTRADO")
    divergentes = status_lista.count("DIVERGENTE")
    print("-" * 50)
    print(f"Concluido: {len(df)} paradas | {divergentes} divergente(s) | {nao_encontradas} nao encontrada(s)")
    print(f"Arquivo salvo em: {saida}")


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        arquivo = sys.argv[1]
    else:
        candidatos = glob.glob("*.xlsx")
        if not candidatos:
            print("Uso: python gerar_links_rota.py <arquivo.xlsx>")
            sys.exit(1)
        arquivo = candidatos[0]
        print(f"Nenhum arquivo informado, usando: {arquivo}\n")

    processar(arquivo)
