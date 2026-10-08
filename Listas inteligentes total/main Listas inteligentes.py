import pandas as pd

# Defina o caminho do arquivo Excel usando barras invertidas duplas ou o prefixo r
dados = r"C:\Users\usuario\Desktop\Pasta\LISTAS INTELIGENTES - CAMPANHAS\Listas Inteligentes - Cooperativa - 11.03.2025.xlsx"
 # Caminho do seu arquivo Excel

# Carregar o arquivo Excel
pfn = pd.read_excel(dados, sheet_name="Base PF-2025")

# Defina os outros parâmetros
mes_campanhapf = "03-2025"
campanhapf = "SEGURO DE VIDA DIA DA MULHER"

# Filtrando os dados
pfn = pfn[pfn["Mês-Ano"] == mes_campanhapf]
pfn = pfn[pfn["Campanha"] == campanhapf]

# Selecionando as colunas necessárias
pfn_final = pfn[[
    "É Cliente",
    "Cooperativa/Código",
    "PA",
    "Campanha",
    "Perfil",
    "Informações pessoais - CPF",
    "Lista Sistema - Telefone Celular",
    "Lista Sistema - E-mail",
    "Informações pessoais - Nome",
    "Localização e contato - Município",
    "Localização e contato - UF"
]]

# Tratamento dos dados
pfn_final["Informações pessoais - CPF"] = pfn_final["Informações pessoais - CPF"].astype(str).str.rjust(11, fillchar='0')
pfn_final["PA"] = pfn_final["PA"].str.extract(pat = '(\d+)')
pfn_final["Lista Sistema - E-mail"] = pfn_final["Lista Sistema - E-mail"].str.strip("[]")
pfn_final["Lista Sistema - E-mail"] = pfn_final["Lista Sistema - E-mail"].str.split(";").str[0]
pfn_final = pfn_final[pfn_final["Lista Sistema - E-mail"].notna()]
pfn_final["Lista Sistema - Telefone Celular"] = pfn_final["Lista Sistema - Telefone Celular"].str.strip("[]")
pfn_final["Lista Sistema - Telefone Celular"] = pfn_final["Lista Sistema - Telefone Celular"].str.split(";").str[0]
pfn_final["Lista Sistema - Telefone Celular"] = pfn_final["Lista Sistema - Telefone Celular"].str.replace(r'\(|\)|\s', '')

# Renomeando as colunas
pfn_final.columns = ["Você já é Associado?", "Coop_Numero", "Ponto de atendimento", "Campanha", "Perfil Base", "CPF ou CNPJ", "Telefone", "E-mail", "Nome", "Cidade", "UF"]



# Caminho do arquivo para salvar no diretório de Downloads
caminho_destino = r"C:\Users\usuario\Downloads\pf_"  + campanhapf + ".xlsx"

# Salvar o arquivo Excel
pfn_final.to_excel(caminho_destino, index=False)

# Informar que o arquivo foi salvo com sucesso
print(f"Arquivo salvo com sucesso em: {caminho_destino}")

        

# Opcional: se precisar baixar o arquivo diretamente (no ambiente local, não funcionará como no Google Colab)
# Em um ambiente local, você não pode usar files.download(), mas pode mover ou acessar o arquivo no diretório onde foi salvo.
