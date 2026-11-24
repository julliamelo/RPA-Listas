# Segmenta uma lista grande em um arquivo Excel por unidade (versão genérica, adapte as colunas ao seu arquivo).
from pathlib import Path

import pandas as pd

# ---------- 1. CONFIGURAÇÃO: é aqui que você adapta para a sua empresa ----------
ARQUIVO_ENTRADA = Path("dados/exemplo_listas.xlsx")  # planilha com a lista completa
ABA = "Base"                                         # nome da aba
PASTA_SAIDA = Path("saida")                          # onde os arquivos vão ser salvos

MES_ANO = "03-2025"          # mês de referência
CAMPANHA = "CAMPANHA EXEMPLO"

COLUNA_UNIDADE = "Unidade/Código"  # coluna usada para separar os arquivos (um por valor)

# Nomes das colunas do SEU Excel. Se na sua planilha tiver outro nome, troque aqui.
COL_MES = "Mês-Ano"
COL_CAMPANHA = "Campanha"
COL_DOCUMENTO = "Documento"
COL_TIPO_DOC = "Tipo de documento"  # CPF (11 dígitos) ou CNPJ (14 dígitos)
COL_TELEFONE = "Telefone"
COL_EMAIL = "E-mail"
COL_PONTO = "Ponto de atendimento"

COLUNAS_FINAIS = [
    "Já é cliente", COLUNA_UNIDADE, COL_PONTO, COL_CAMPANHA, "Perfil",
    COL_TIPO_DOC, COL_DOCUMENTO, COL_TELEFONE, COL_EMAIL, "Nome", "Município", "UF",
    "Status/Retorno", "Observação",
]


# ---------- 2. FUNÇÕES DE LIMPEZA ----------
def primeiro_valor(serie: pd.Series) -> pd.Series:
    # "[a@x.com;b@x.com]" -> "a@x.com" (tira os colchetes e fica com o primeiro item)
    return serie.str.strip("[]").str.split(";").str[0]


def limpar(df: pd.DataFrame) -> pd.DataFrame:
    df = df[df[COL_MES] == MES_ANO]
    df = df[df[COL_CAMPANHA] == CAMPANHA]
    df = df[COLUNAS_FINAIS].copy()  # .copy() evita aviso do pandas ao alterar

    # devolve os zeros da esquerda que o Excel come: CPF fica com 11 dígitos, CNPJ com 14
    tamanho = df[COL_TIPO_DOC].map({"CPF": 11, "CNPJ": 14}).fillna(11).astype(int)
    df[COL_DOCUMENTO] = [str(d).zfill(n) for d, n in zip(df[COL_DOCUMENTO], tamanho)]
    # ponto de atendimento: fica só com o número
    df[COL_PONTO] = df[COL_PONTO].astype(str).str.extract(r"(\d+)", expand=False)

    df[COL_EMAIL] = primeiro_valor(df[COL_EMAIL])
    df = df[df[COL_EMAIL].notna()]  # sem e-mail não entra na lista

    df[COL_TELEFONE] = primeiro_valor(df[COL_TELEFONE]).str.replace(r"\(|\)|\s", "", regex=True)
    df["Status/Retorno"] = df["Status/Retorno"].fillna("").astype(str).str.strip("[]")
    df[COLUNA_UNIDADE] = df[COLUNA_UNIDADE].astype(str).str.strip()
    return df


# ---------- 3. PROGRAMA PRINCIPAL ----------
def main():
    base = pd.read_excel(ARQUIVO_ENTRADA, sheet_name=ABA, dtype=str)
    lista = limpar(base)

    PASTA_SAIDA.mkdir(exist_ok=True)  # cria a pasta se não existir

    # um arquivo para cada unidade encontrada na própria lista
    for unidade, grupo in lista.groupby(COLUNA_UNIDADE):
        destino = PASTA_SAIDA / f"{MES_ANO}_{CAMPANHA}_unidade_{unidade}.xlsx"
        try:
            grupo.to_excel(destino, index=False)
            print(f"Unidade {unidade}: {len(grupo)} registros salvos em {destino}")
        except PermissionError:
            print(f"Sem permissão para salvar {destino}. Feche o arquivo se ele estiver aberto.")


if __name__ == "__main__":
    main()
