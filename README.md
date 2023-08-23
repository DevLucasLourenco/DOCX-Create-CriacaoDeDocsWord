## Geração de Documentos a partir de Dados de Planilha Excel

Este repositório contém um script em Python que lê dados de uma planilha Excel, realiza algumas tratativas e gera documentos DOCX personalizados com base nos dados lidos.

### Bibliotecas Utilizadas

- `import copy`: Usada para criar cópias independentes de objetos.
- `import pandas as pd`: Utilizada para trabalhar com dados em formato de DataFrame.
- `import docx`: Biblioteca para criar e manipular documentos Word (DOCX).
- `from pathlib import Path`: Usada para lidar com caminhos de arquivo de maneira eficiente.
- `from docx.shared import Inches, Pt`: Utilizada para definir tamanhos de fonte e outras medidas.
- `from docx.enum.text import WD_PARAGRAPH_ALIGNMENT`: Usada para alinhar parágrafos.
- `from docx.oxml.ns import qn`: Usada para lidar com namespaces XML.
- `import time`: Utilizada para trabalhar com timestamps.

### Funções Principais

1. `tratativa_data(data)`: Função para formatar uma data no formato "dia/mês/ano".
2. `tratativa_dataExtenso(data)`: Função para converter uma data no formato "dia/mês/ano" para o formato "dia de mês de ano".
3. `CriarDocumento(doc_base, nome, dados)`: Função para criar um documento DOCX personalizado com base em um documento base (`doc_base`) e um conjunto de dados (`dados`).
4. `run()`: Função principal que carrega os dados da planilha Excel, faz tratativas e gera os documentos.

### Execução do Script

1. Carrega um documento base DOCX (`base.docx`) uma vez para ser usado como modelo.
2. Itera sobre cada linha da planilha Excel (`base.xlsx`).
3. Para cada linha, coleta os dados relevantes.
4. Cria um dicionário de substituição com os dados coletados.
5. Chama a função `CriarDocumento()` para gerar um novo documento personalizado com os dados.
6. Os documentos personalizados são salvos na pasta `docs feitos/`.

### Observações

- O código utiliza tratativas de data para formatar datas no formato desejado.
- Os documentos finais são gerados substituindo as tags especiais (como `==NOME==`) pelos valores correspondentes.
- O alinhamento do texto nos documentos é ajustado para centralizado.
- O texto após a tag `==DEPOISDEZDIA==` em cada parágrafo é substituído pelo valor correspondente.
- O nome do documento final é gerado com base na data de desligamento e no nome da pessoa.

### Como Usar

1. Certifique-se de ter as bibliotecas necessárias instaladas (`pandas`, `docx`, etc.).
2. Coloque o documento base (`base.docx`) no mesmo diretório do script.
3. Coloque a planilha Excel (`base.xlsx`) no mesmo diretório ou ajuste o caminho conforme necessário.
4. Execute o script. Os documentos gerados serão salvos na pasta `docs feitos/`.

**Nota:** Este é apenas um resumo explicativo do código. Certifique-se de ler e entender o código completo antes de usá-lo em um ambiente de produção.
