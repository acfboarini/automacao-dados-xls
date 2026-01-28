# leitura.py

import xlrd
import numpy as np
from config import LINHAS_MED_MES


def to_float(valor):
    """
    Converte valor para float.
    Strings vazias, espaços ou valores inválidos viram NaN.
    """
    try:
        if isinstance(valor, str):
            valor = valor.strip()
            if valor == "":
                return np.nan
        return float(valor)
    except Exception:
        return np.nan


def ler_planilha(caminho_arquivo):
    """
    Lê uma planilha anual e retorna:
    - colunas  : lista de nomes das variáveis (linha 6, colunas B → U)
    - unidades : lista de unidades das variáveis (linha 7)
    - dados    : dict {mes: {variavel: valor}}
    """

    wb = xlrd.open_workbook(caminho_arquivo)
    sheet = wb.sheet_by_name("Plan1")

    # ===== 1) Ler nomes (linha 6) e unidades (linha 7)
    nomes_raw = sheet.row_values(5, start_colx=1, end_colx=21)     # B → U
    unidades_raw = sheet.row_values(6, start_colx=1, end_colx=21)

    colunas = []
    unidades = []
    indices = []

    for idx, (nome, unidade) in enumerate(zip(nomes_raw, unidades_raw), start=1):
        nome = str(nome).strip()
        unidade = str(unidade).strip()

        # remover colunas vazias e coluna "Hora"
        if nome == "" or nome.lower() == "hora":
            continue

        colunas.append(nome)
        unidades.append(unidade)
        indices.append(idx)

    # ===== 2) Ler valores MED MES
    dados = {}

    for mes, linha in LINHAS_MED_MES.items():
        valores_linha = sheet.row_values(linha - 1)

        registro = {}
        for nome, col_idx in zip(colunas, indices):
            registro[nome] = to_float(valores_linha[col_idx])

        dados[mes] = registro

    return colunas, unidades, dados
