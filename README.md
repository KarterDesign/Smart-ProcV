
# smart-procv.py

## Descrição

Este script `smart-procv.py` tem como objetivo carregar dados de duas planilhas Excel, identificar e relacionar nomes que possuem colunas vazias em uma das planilhas, utilizando uma lógica de correspondência baseada em similaridade de strings. O resultado é salvo em uma nova planilha Excel.

## Funcionalidades

-   Carregar dados de duas planilhas a partir de um arquivo Excel.
-   Verificar e remover valores não-string de uma coluna específica.
-   Filtrar nomes em uma planilha que possuem colunas vazias.
-   Utilizar a biblioteca `difflib` para encontrar correspondências aproximadas entre os nomes das duas planilhas.
-   Salvar o resultado das correspondências em uma nova planilha Excel.

## Pré-requisitos

-   Python 3.x
-   Bibliotecas Python: `pandas`, `difflib`, `openpyxl`

## Instalação

1.  Clone este repositório para sua máquina local.
    
    bash
    
    Copiar código
    
    `git clone https://github.com/seu-usuario/seu-repositorio.git` 
    
2.  Navegue até o diretório do projeto.
    
    bash
    
    Copiar código
    
    `cd seu-repositorio` 
    
3.  Instale as dependências necessárias.
    
    bash
    
    Copiar código
    
    `pip install pandas openpyxl` 
    

## Uso

1.  Certifique-se de que o arquivo `procv.xlsx` esteja no mesmo diretório que o script `smart-procv.py`.
    
2.  Execute o script.
    
    bash
    
    Copiar código
    
    `python smart-procv.py` 
    
3.  O script irá gerar um arquivo `resultado_relacionado.xlsx` com os resultados das correspondências.
    

## Código

python

Copiar código

`import pandas as pd
from difflib import get_close_matches

## Carregar os dados das duas planilhas
sheet1_data = pd.read_excel('procv.xlsx', sheet_name='Sheet1')
sheet2_data = pd.read_excel('procv.xlsx', sheet_name='Sheet2')

## Verificar e remover possíveis valores não-string da coluna 'Cliente' em 'Sheet2'
sheet2_data = sheet2_data[sheet2_data['Cliente'].apply(lambda x: isinstance(x, str))]

## Filtrar os nomes da "Sheet1" que estão com a coluna "Proposta" vazia
missing_proposals = sheet1_data[sheet1_data['Proposta'].isna()]

## Relacionar os nomes
matched_names = []
for name in missing_proposals['NomeUP']:
    # Buscar o nome mais próximo na "Sheet2"
    close_match = get_close_matches(name, sheet2_data['Cliente'], n=1, cutoff=0.7)
    if close_match:
        matched_name = close_match[0]
        proposal = sheet2_data[sheet2_data['Cliente'] == matched_name]['Proposta'].values[0]
        matched_names.append({'NomeUP': name, 'Cliente Matched': matched_name, 'Proposta': proposal})
    else:
        matched_names.append({'NomeUP': name, 'Cliente Matched': 'Not Found', 'Proposta': 'Not Found'})

## Criar um DataFrame com os resultados
matched_df = pd.DataFrame(matched_names)

## Salvar os resultados em um novo arquivo Excel
matched_df.to_excel('resultado_relacionado.xlsx', index=False)` 
