Automação de Tratamento e Segmentação de Campanhas (RPA)
Script em Python desenvolvido para automatizar o processamento, limpeza e segmentação de bases mensais de prospecção comercial, substituindo o trabalho manual de filtragem e cópia de dados entre dezenas de unidades.

🎯 O Problema
Mensalmente, é recebida uma base consolidada contendo leads propensos a diferentes produtos financeiros (ex.: consignado, investimentos, entre outros), acumulada desde 2023. Anteriormente, era necessário abrir o arquivo manualmente, remover colunas irrelevantes, filtrar por período e produto, e separar individualmente os dados para cada uma das 31 cooperativas — um processo repetitivo e sujeito a falhas humanas.

⚙️ Como o Código Funciona
O robô executa um fluxo completo de ETL (Extração, Transformação e Carga) e formatação:

Leitura e Higienização: Carrega a planilha principal e descarta colunas desnecessárias, mantendo apenas os campos essenciais para a 
abordagem comercial (data, campanha, cooperativa, nome, telefone, endereço, etc.).

Filtragem Dinâmica: Segmenta os registros de acordo com o mês de referência, a campanha específica e a cooperativa desejada.

Agrupamento por Unidade: Consolida os dados de cada cooperativa (entre as 31 disponíveis) de forma isolada e estruturada, eliminando a
necessidade de recortar e colar registros manualmente.

Formatação Executiva (Excel): Gera o arquivo final pronto para uso imediato pelas equipes comerciais, já aplicando identidade visual 
(cabeçalhos estilizados em verde) e filtros nativos ativados no Excel.

🛠️ Tecnologias Utilizadas
Python

Pandas (manipulação, limpeza e filtragem de grandes volumes de dados)

OpenPyXL (estilização das planilhas e ativação automática de filtros no Excel)

📈 Resultados Obtidos
Eliminação de trabalho manual: Redução drástica no tempo de preparação das listas mensais para 31 unidades.

Padronização: Todas as entregas seguem exatamente o mesmo padrão visual e estrutural, prontas para filtragem rápida pelos usuários finais.

Escalabilidade: Capacidade de processar bases históricas contínuas (desde 2023) sem travamentos.
