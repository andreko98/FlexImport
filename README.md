FlexImport é uma aplicação Windows Forms desenvolvida em C# (.NET 8) para facilitar a importação, conversão e padronização de planilhas de produtos exportadas do Bling para um formato compatível com a Loja Integrada.

Funcionalidades
•	Importação de CSV: Leitura de arquivos CSV exportados do Bling, com suporte a diferentes formatos numéricos brasileiros.
•	Conversão e Padronização: Geração de planilhas no formato esperado por marketplaces, com tratamento de categorias, subcategorias, variações e atributos de produtos.
•	Geração de SKUs únicos: Criação automática de SKUs para produtos e variações, evitando duplicidades.
•	Exportação para Excel: Conversão dos dados processados para arquivos Excel (.xlsx) prontos para upload.
•	Interface intuitiva: Interface gráfica simples para seleção de arquivos, categorias, subcategorias e gênero dos produtos.

Tecnologias Utilizadas
•	C# 12 / .NET 8
•	Windows Forms
•	CsvHelper para leitura e conversão de CSV
•	EPPlus para geração de arquivos Excel

Como usar
1.	Exporte sua planilha de produtos do Bling em formato CSV.
2.	Abra o FlexImport, selecione o arquivo CSV e defina o local de salvamento.
3.	Escolha a categoria, subcategoria e gênero dos produtos.
4.	Clique em "Converter" para gerar a planilha padronizada.
5.	O arquivo Excel será salvo no local escolhido, pronto para ser importado na Loja Integrada.
   
Observações
•	O sistema trata automaticamente números no formato brasileiro (ex: 1.234,56), convertendo corretamente para os tipos numéricos do sistema.
•	O projeto pode ser facilmente adaptado para outros ERPs ou formatos de planilha.
