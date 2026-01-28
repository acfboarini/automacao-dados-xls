import pandas as pd
import numpy as np

arquivo = "output/medias2.xlsx"
xls = pd.ExcelFile(arquivo)

# acumulador por intervalo
acumulado = {
    "dez-fev": [],
    "jan-mar": [],
    "mai-jul": [],
    "jun-ago": [],
}

nomes_variaveis = None
unidades_variaveis = None

for aba in xls.sheet_names:
    # ler tudo sem header
    df = pd.read_excel(
        arquivo,
        sheet_name=aba,
        header=None
    )

    # linha 0 -> nomes, linha 1 -> unidades
    nomes = df.iloc[0, 1:]
    unidades = df.iloc[1, 1:]

    # guardar referência (todas as abas já estão alinhadas por construção)
    if nomes_variaveis is None:
        nomes_variaveis = nomes
        unidades_variaveis = unidades

    # dados começam na linha 2
    dados = df.iloc[2:, 1:]
    dados.index = df.iloc[2:, 0]  # dez-fev, jan-mar, etc.

    for intervalo in acumulado:
        acumulado[intervalo].append(dados.loc[intervalo])

# calcular média final
resultado = {}

for intervalo, listas in acumulado.items():
    df_concat = pd.concat(listas, axis=1).T
    resultado[intervalo] = df_concat.mean(skipna=True).round(2)

df_final = pd.DataFrame.from_dict(resultado, orient="index")

# ===== escrever Excel final com nomes + unidades =====
with pd.ExcelWriter("output/medias3.xlsx", engine="openpyxl") as writer:
    df_final.to_excel(writer, sheet_name="media3", startrow=2, header=False)

    ws = writer.sheets["media3"]

    # linha 1 -> nomes
    for col_idx, nome in enumerate(nomes_variaveis, start=2):
        ws.cell(row=1, column=col_idx, value=nome)

    # linha 2 -> unidades
    for col_idx, unidade in enumerate(unidades_variaveis, start=2):
        ws.cell(row=2, column=col_idx, value=unidade)

    ws.cell(row=1, column=1, value="Intervalo")
