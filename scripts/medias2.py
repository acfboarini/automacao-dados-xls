# medias2.py

import numpy as np
import pandas as pd
from leitura import ler_planilha
from config import PADRAO_ARQUIVO, ANOS_VALIDOS


def media_dicts(*dicts):
    """
    Calcula a média ignorando NaN para cada variável.
    Retorna valores arredondados para 2 casas decimais.
    """
    resultado = {}
    chaves = set().union(*dicts)

    for chave in chaves:
        valores = [d.get(chave, np.nan) for d in dicts]
        resultado[chave] = round(np.nanmean(valores), 2)

    return resultado


with pd.ExcelWriter("output/medias2.xlsx", engine="openpyxl") as writer:

    for ano in ANOS_VALIDOS:
        colunas, unidades, dados_atual = ler_planilha(
            PADRAO_ARQUIVO.format(ano=ano)
        )
        _, _, dados_anterior = ler_planilha(
            PADRAO_ARQUIVO.format(ano=ano - 1)
        )

        linhas = {
            "dez-fev": media_dicts(
                dados_anterior["dez"],
                dados_atual["jan"],
                dados_atual["fev"]
            ),
            "jan-mar": media_dicts(
                dados_atual["jan"],
                dados_atual["fev"],
                dados_atual["mar"]
            ),
            "mai-jul": media_dicts(
                dados_atual["mai"],
                dados_atual["jun"],
                dados_atual["jul"]
            ),
            "jun-ago": media_dicts(
                dados_atual["jun"],
                dados_atual["jul"],
                dados_atual["ago"]
            ),
        }

        df = pd.DataFrame.from_dict(linhas, orient="index")

        # garantir ordem EXATA das colunas do ano
        df = df.reindex(columns=colunas)

        sheet_name = str(ano)

        # escrever dados (sem header, sem duplicar nomes)
        df.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=2,
            header=False
        )

        ws = writer.sheets[sheet_name]

        # linha 1 → nomes das variáveis
        ws.cell(row=1, column=1, value="Intervalo")
        for col_idx, nome in enumerate(colunas, start=2):
            ws.cell(row=1, column=col_idx, value=nome)

        # linha 2 → unidades
        for col_idx, unidade in enumerate(unidades, start=2):
            ws.cell(row=2, column=col_idx, value=unidade)
