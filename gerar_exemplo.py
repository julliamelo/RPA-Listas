# Cria uma planilha de exemplo com duas abas (PF e PJ) e dados 100% FALSOS.
import random
from pathlib import Path

import pandas as pd

random.seed(1)  # mesma planilha toda vez

PRIMEIROS_NOMES = ["Carla", "João", "Maria", "Pedro", "Ana", "Lucas", "Beatriz", "Rafael"]
TIPOS_EMPRESA = ["Padaria", "Mercado", "Oficina", "Farmácia", "Papelaria", "Escola"]


def linha_comum(i: int) -> dict:
    return {
        "Mês-Ano": random.choice(["03-2025", "03-2025", "02-2025"]),
        "Campanha": random.choice(["CAMPANHA EXEMPLO", "CAMPANHA EXEMPLO", "OUTRA CAMPANHA"]),
        "Já é cliente": random.choice(["Sim", "Não"]),
        "Unidade/Código": f"{random.randint(1, 5):03d}",
        "Ponto de atendimento": f"PA {random.randint(1, 20):02d}",
        "Perfil": random.choice(["Perfil A", "Perfil B", "Perfil C"]),
        # número pequeno, sem zeros: o script devolve os zeros (…0021), como o Excel real come
        "Documento": i,
        "Telefone": f"[(00) 00000-{i:04d};(00) 00001-{i:04d}]",
        "E-mail": f"[contato{i:04d}@exemplo.com;outro{i:04d}@exemplo.com]" if i % 15 else None,
        "Município": random.choice(["Cidade Um", "Cidade Dois", "Cidade Três"]),
        "UF": random.choice(["AA", "BB", "CC"]),
        "Status/Retorno": random.choice(["[Contato feito]", "[Sem retorno]", None]),
        "Observação": None,
        # colunas extras (o script ignora, como na planilha real)
        "Endereço": f"Rua Exemplo, {i}",
        "Bairro": "Bairro Exemplo",
        "CEP": "00000-000",
        "Data de cadastro": "01/01/2025",
        "Produto de interesse": random.choice(["Produto A", "Produto B", "Produto C"]),
        "Score de propensão": random.randint(0, 100),
        "Canal preferido": random.choice(["Telefone", "E-mail", "Presencial"]),
    }


pf = []
for i in range(1, 101):
    linha = linha_comum(i)
    linha["Nome"] = f"{random.choice(PRIMEIROS_NOMES)} Exemplo {i:04d}"
    pf.append(linha)

pj = []
for i in range(1, 61):
    linha = linha_comum(i)
    tipo = random.choice(TIPOS_EMPRESA)
    linha["Razão social"] = f"{tipo} Exemplo {i:04d} Ltda"
    linha["Nome fantasia"] = f"{tipo} Exemplo {i:04d}"
    linha["Porte"] = random.choice(["MEI", "ME", "EPP"])
    pj.append(linha)

Path("dados").mkdir(exist_ok=True)
with pd.ExcelWriter("dados/exemplo_listas.xlsx") as arquivo:
    pd.DataFrame(pf).to_excel(arquivo, sheet_name="Base PF", index=False)
    pd.DataFrame(pj).to_excel(arquivo, sheet_name="Base PJ", index=False)
print("Criado: dados/exemplo_listas.xlsx (abas Base PF e Base PJ)")
