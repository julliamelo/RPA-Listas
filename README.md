# Segmentação automática de listas **(RPA com Python)**

*Script em Python que lê uma lista grande, limpa os dados e separa em **um arquivo Excel por unidade**. Substitui o trabalho manual de abrir a planilha, filtrar, copiar e colar para cada área.*

Esta versão é **genérica**: usa dados de exemplo (todos falsos) e pode ser adaptada para qualquer empresa trocando só a configuração e os nomes das colunas.

## O problema

Todo mês chega uma base consolidada de contatos para campanhas comerciais. Para usar, era preciso abrir o arquivo, remover colunas que não servem, filtrar por mês e campanha e separar os registros de cada unidade. Era repetitivo e fácil de errar.

## O que o script faz

1. **Lê** a planilha principal, que tem duas abas: **Base PF** (pessoa física, CPF) e **Base PJ** (pessoa jurídica, CNPJ).
2. **Filtra** pelo mês e pela campanha escolhidos.
3. **Limpa** os dados: CPF (11 dígitos) e CNPJ (14 dígitos) com os zeros à esquerda de volta, só o primeiro e-mail e telefone de cada contato, ponto de atendimento só com o número.
4. **Descarta** registros sem e-mail.
5. **Separa** em um arquivo Excel por unidade, na pasta `saida/`, um conjunto para PF e outro para PJ (cada aba com as suas colunas).

## Como testar

```bash
pip install -r requirements.txt
python gerar_exemplo.py        # cria dados/exemplo_listas.xlsx com dados falsos
python segmentar_listas.py     # gera um arquivo por unidade em saida/
```

## Como adaptar para os seus dados

Abra `segmentar_listas.py` e mude o bloco **CONFIGURAÇÃO** no começo do arquivo: caminho da planilha, nome da aba, mês, campanha e nomes das colunas. O resto do código não precisa mudar.

## Tecnologias

- Python
- pandas (leitura, limpeza e filtragem)
- openpyxl (gravação dos arquivos Excel)

## Resultado

Preparação das listas mensais sem trabalho manual, com o mesmo padrão em todos os arquivos.

---
Mais projetos no [meu perfil](https://github.com/julliamelo).

## Direitos

© 2026 Júllia Melo. Todos os direitos reservados. Veja o [LICENSE](LICENSE).
