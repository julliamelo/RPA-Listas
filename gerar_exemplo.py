# Cria uma planilha de exemplo com dados 100% FALSOS para testar o segmentar_listas.py.
import random
from pathlib import Path

import pandas as pd

random.seed(1)  # mesma planilha toda vez

linhas = []
for i in range(1, 121):
    eh_pj = i % 4 == 0  # um a cada quatro é empresa
    linhas.append({
        # --- colunas que o script usa ---
        "Mês-Ano": random.choice(["03-2025", "03-2025", "02-2025"]),
        "Campanha": random.choice(["CAMPANHA EXEMPLO", "CAMPANHA EXEMPLO", "OUTRA CAMPANHA"]),
        "Já é cliente": random.choice(["Sim", "Não"]),
        "Unidade/Código": f"{random.randint(1, 5):03d}",
        "Ponto de atendimento": f"PA {random.randint(1, 20):02d}",
        "Perfil": random.choice(["Perfil A", "Perfil B", "Perfil C"]),
        "Tipo de documento": "CNPJ" if eh_pj else "CPF",
        # número pequeno e sem zeros: o script completa com zeros (…0021), como o Excel real faz
        "Documento": i,
        "Telefone": f"[(00) 00000-{i:04d};(00) 00001-{i:04d}]",
        "E-mail": f"[pessoa{i:04d}@exemplo.com;outro{i:04d}@exemplo.com]" if i % 15 else None,
        "Nome": f"Empresa Exemplo {i:04d}" if eh_pj else f"Pessoa Exemplo {i:04d}",
        "Município": random.choice(["Cidade Um", "Cidade Dois", "Cidade Três"]),
        "UF": random.choice(["AA", "BB", "CC"]),
        "Status/Retorno": random.choice(["[Contato feito]", "[Sem retorno]", None]),
        "Observação": None,
        # --- colunas extras (o script ignora, como na planilha real) ---
        "Endereço": f"Rua Exemplo, {i}",
        "Bairro": "Bairro Exemplo",
        "CEP": "00000-000",
        "Data de cadastro": "01/01/2025",
        "Produto de interesse": random.choice(["Produto A", "Produto B", "Produto C"]),
        "Score de propensão": random.randint(0, 100),
        "Canal preferido": random.choice(["Telefone", "E-mail", "Presencial"]),
    })

Path("dados").mkdir(exist_ok=True)
pd.DataFrame(linhas).to_excel("dados/exemplo_listas.xlsx", sheet_name="Base", index=False)
print("Criado: dados/exemplo_listas.xlsx")
