Sorteador de Desempate (Python + Tkinter)
Este projeto é uma ferramenta com interface gráfica (GUI) desenvolvida em Python para automatizar o processo de desempate em listas de candidatos ou itens organizados por grupos. Ele permite realizar sorteios aleatórios ou baseados em "sementes" (seeds) específicas para garantir a auditabilidade e reprodutibilidade do resultado.
🚀 Funcionalidades
•	Importação de Excel: Lê arquivos .xlsx ou .xls contendo uma coluna chamada GRUPO.
•	Três Métodos de Sorteio:
1.	Aleatório: Usa o tempo atual do sistema como semente.
2.	Semente Numérica: O usuário define um número manual para gerar o resultado.
3.	Semente de Hash: Transforma qualquer texto/código (SHA-256) em uma semente numérica.
•	Processamento por Grupo: O sorteio é realizado de forma independente dentro de cada grupo.
•	Exportação Formatada: Gera um novo arquivo Excel com:
o	Coluna ORDEM DESEMPATE preenchida.
o	Ordenação automática por Grupo e Ordem.
o	Espaçamento visual (linha em branco) entre grupos diferentes.
o	Ajuste automático da largura das colunas.
________________________________________
🛠️ Tecnologias Utilizadas
•	Python 3.x
•	Pandas & NumPy: Para manipulação de dados e lógica matemática do sorteio.
•	Tkinter: Para a interface gráfica e diálogos de arquivos.
•	XlsxWriter: Para a criação e formatação avançada do arquivo Excel de saída.
•	Hashlib: Para a geração de sementes a partir de strings (Hash).
________________________________________
📋 Pré-requisitos
Antes de rodar o script, você precisará instalar as bibliotecas necessárias:
Bash
pip install pandas numpy xlsxwriter openpyxl
________________________________________
📖 Como Usar
1.	Prepare sua planilha: O arquivo de entrada deve ter, obrigatoriamente, uma coluna com o cabeçalho GRUPO.
2.	Execute o script:
Bash
python nome_do_seu_arquivo.py
3.	Selecione o arquivo: Uma janela se abrirá para você escolher a planilha Excel.
4.	Escolha o método:
o	Se precisar que o sorteio seja auditável (repetível), escolha Semente Numérica ou Hash e anote o valor utilizado.
5.	Confira o resultado: O sistema salvará um arquivo chamado RESULTADO_DESEMPATE.xlsx na mesma pasta do arquivo original.
________________________________________
🔍 Detalhes Técnicos (Auditabilidade)
O código utiliza o gerador de números pseudo-aleatórios do NumPy (np.random.seed). Isso significa que:
Se duas pessoas utilizarem a mesma planilha e a mesma semente (número ou hash), o resultado da ordem de desempate será exatamente o mesmo, garantindo total transparência ao processo.
________________________________________
⚖️ Licença
Este projeto está sob a licença GNU General Public License v3.0.

