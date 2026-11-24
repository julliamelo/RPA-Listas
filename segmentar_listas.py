# Segmenta a lista (aba PF e aba PJ) em um arquivo Excel por unidade, versão genérica.
from pathlib import Path

import pandas as pd

# ---------- 1. CONFIGURAÇÃO: é aqui que você adapta para a sua empresa ----------
ARQUIVO_ENTRADA = Path("dados/exemplo_listas.xlsx")  # planilha com a lista completa
PASTA_SAIDA = Path("saida")                          # onde os arquivos vão ser salvos

MES_ANO = "03-2025"          # mês de referência
CAMPANHA = "CAMPANHA EXEMPLO"

COLUNA_UNIDADE = "Unidade/Código"  # coluna usada para separar os arquivos (um por valor)

# Nomes das colunas do SEU Excel. Se na sua planilha tiver outro nome, troque aqui.
COL_MES = "Mês-Ano"
COL_CAMPANHA = "Campanha"
COL_DOCUMENTO = "Documento"
COL_TELEFONE = "Telefone"
COL_EMAIL = "E-mail"
COL_PONTO = "Ponto de atendimento"

# Colunas que existem nas duas abas
COLUNAS_COMUNS = [
    "Já é cliente", COLUNA_UNIDADE, COL_PONTO, COL_CAMPANHA, "Perfil",
    COL_DOCUMENTO, COL_TELEFONE, COL_EMAIL,
]
COLUNAS_FINAIS_COMUNS = ["Município", "UF", "Status/Retorno", "Observação"]

# Cada aba tem seu tamanho de documento e suas colunas de identificação
ABAS = {
    "PF": {  # pessoa física: CPF tem 11 dígitos
        "aba": "Base PF",
        "tamanho_documento": 11,
        "colunas_nome": ["Nome"],
    },
    "PJ": {  # pessoa jurídica: CNPJ tem 14 dígitos
        "aba": "Base PJ",
        "tamanho_documento": 14,
        "colunas_nome": ["Razão social", "Nome fantasia", "Porte"],
    },
}


# ---------- 2. FUNÇÕES DE LIMPEZA ----------
def primeiro_valor(serie: pd.Series) -> pd.Series:
    # "[a@x.com;b@x.com]" -> "a@x.com" (tira os colchetes e fica com o primeiro item)
    return serie.str.strip("[]").str.split(";").str[0]


def limpar(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    df = df[df[COL_MES] == MES_ANO]
    df = df[df[COL_CAMPANHA] == CAMPANHA]
    colunas = COLUNAS_COMUNS + config["colunas_nome"] + COLUNAS_FINAIS_COMUNS
    df = df[colunas].copy()  # .copy() evita aviso do pandas ao alterar

    # devolve os zeros da esquerda que o Excel come (CPF = 11 dígitos, CNPJ = 14)
    df[COL_DOCUMENTO] = df[COL_DOCUMENTO].astype(str).str.zfill(config["tamanho_documento"])
    # ponto de atendimento: fica só com o número
    df[COL_PONTO] = df[COL_PONTO].astype(str).str.extract(r"(\d+)", expand=False)

    df[COL_EMAIL] = primeiro_valor(df[COL_EMAIL])
    df = df[df[COL_EMAIL].notna()]  # sem e-mail não entra na lista

    df[COL_TELEFONE] = primeiro_valor(df[COL_TELEFONE]).str.replace(r"\(|\)|\s", "", regex=True)
    df["Status/Retorno"] = df["Status/Retorno"].fillna("").astype(str).str.strip("[]")
    df[COLUNA_UNIDADE] = df[COLUNA_UNIDADE].astype(str).str.strip()
    return df


# ---------- 3. PROGRAMA PRINCIPAL ----------
def processar_aba(tipo: str, config: dict):
    base = pd.read_excel(ARQUIVO_ENTRADA, sheet_name=config["aba"], dtype=str)
    lista = limpar(base, config)

    # um arquivo para cada unidade encontrada na própria lista
    for unidade, grupo in lista.groupby(COLUNA_UNIDADE):
        destino = PASTA_SAIDA / f"{tipo}_{MES_ANO}_{CAMPANHA}_unidade_{unidade}.xlsx"
        try:
            grupo.to_excel(destino, index=False)
            print(f"[{tipo}] Unidade {unidade}: {len(grupo)} registros em {destino}")
        except PermissionError:
            print(f"Sem permissão para salvar {destino}. Feche o arquivo se ele estiver aberto.")


def main():
    PASTA_SAIDA.mkdir(exist_ok=True)  # cria a pasta se não existir
    for tipo, config in ABAS.items():
        processar_aba(tipo, config)


if __name__ == "__main__":
    main()
