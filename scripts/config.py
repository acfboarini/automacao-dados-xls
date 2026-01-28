# config.py

PADRAO_ARQUIVO = "dados_brutos/jan-dez{ano}.xls"

ANOS_VALIDOS = range(1998, 2026)

# MED MES começa na linha 43 e avança de 57 em 57 linhas
LINHAS_MED_MES = {
    "jan": 43,
    "fev": 43 + 57,
    "mar": 43 + 57 * 2,
    "abr": 43 + 57 * 3,
    "mai": 43 + 57 * 4,
    "jun": 43 + 57 * 5,
    "jul": 43 + 57 * 6,
    "ago": 43 + 57 * 7,
    "set": 43 + 57 * 8,
    "out": 43 + 57 * 9,
    "nov": 43 + 57 * 10,
    "dez": 43 + 57 * 11,
}
