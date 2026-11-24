# Cria uma planilha de exemplo com dados FALSOS para testar o segmentar_listas.py.
import random
from pathlib import Path

import pandas as pd

random.seed(1)  # mesma planilha toda vez

linhas = []
for i in range(1, 121):
    unidade = f"{random.randint(1, 5):03d}"
    linhas.append({
        "Mês-Ano": random.choice(["03-2025", "03-2025", "02-2025"]),
        "Campanha": random.choice(["CAMPANHA EXEMPLO", "CAMPANHA EXEMPLO", "OUTRA CAMPANHA"]),
        "Já é cliente": random.choice(["Sim", "Não"]),
        "Unidade/Código": unidade,
        "Ponto de atendimento": f"PA {random.randint(1, 20):02d}",
        "Perfil": random.choice(["Perfil A", "Perfil B", "Perfil C"]),
        "Documento": 1000000 + i,  # número inteiro: perde o zero da esquerda, como no Excel real
        "Telefone": f"[(11) 90000-{i:04d};(11) 90001-{i:04d}]",
        "E-mail": f"[pessoa{i}@exemplo.com;outro{i}@exemplo.com]" if i % 15 else None,
        "Nome": f"Pessoa Exemplo {i:03d}",
        "Município": random.choice(["Cidade Um", "Cidade Dois", "Cidade Três"]),
        "UF": random.choice(["SP", "MG", "RJ"]),
        "Status/Retorno": random.choice(["[Contato feito]", "[Sem retorno]", None]),
        "Observação": None,
    })

Path("dados").mkdir(exist_ok=True)
pd.DataFrame(linhas).to_excel("dados/exemplo_listas.xlsx", sheet_name="Base", index=False)
print("Criado: dados/exemplo_listas.xlsx")
