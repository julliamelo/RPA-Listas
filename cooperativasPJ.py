import pandas as pd
import os

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


# Verifique se a coluna existe
coluna_a_filtrar = 'Cooperativa/Código'  # Certifique-se de usar o nome correto

# Lista de opções para filtrar a coluna 'Cooperativa/Código'
opcoes_a_filtrar = ['001', '002', '003', '004', '005', '006', '007', '008', '009', '010', 
    '011', '012', '013', '014', '015', '016', '017', '018', '019', '020', 
    '021', '022', '023']

# Limpeza dos dados da coluna
if coluna_a_filtrar in pfn_final.columns:
    pfn_final[coluna_a_filtrar] = pfn_final[coluna_a_filtrar].astype(str).str.strip()

    # Filtrar e gerar um arquivo para cada valor na lista de opções
    for coop_codigo in opcoes_a_filtrar:
        pfn_filtrado = pfn_final[pfn_final[coluna_a_filtrar] == coop_codigo]

        # Verificar se o DataFrame filtrado contém dados
        if not pfn_filtrado.empty:
            # Definir o nome do arquivo com base na cooperativa
            caminho_destino = r"C:\Users\usuario\Downloads\PF_" + mes_campanhapf + "-" + campanhapf + "_Cooperativa_" + coop_codigo + ".xlsx"
            
            # Verificar se o diretório de destino existe
            diretorio_destino = os.path.dirname(caminho_destino)
            if not os.path.exists(diretorio_destino):
                os.makedirs(diretorio_destino)  # Cria o diretório se não existir
                print(f"O diretório {diretorio_destino} foi criado.")
                
            try:
                # Salvar o DataFrame filtrado em um novo arquivo Excel
                pfn_filtrado.to_excel(caminho_destino, index=False)
                print(f"Arquivo para a Cooperativa {coop_codigo} salvo com sucesso em: {caminho_destino}")
            except PermissionError:
                print(f"Erro de permissão ao salvar o arquivo. Tente outro diretório.")
            except Exception as e:
                print(f"Erro ao salvar o arquivo: {e}")
        else:
            print(f"Nenhum dado encontrado para a Cooperativa {coop_codigo}.")
else:
    print(f"Coluna '{coluna_a_filtrar}' não encontrada no DataFrame.")
