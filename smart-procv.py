import pandas as pd
from difflib import get_close_matches

# Carregar os dados das duas planilhas
sheet1_data = pd.read_excel('procv.xlsx', sheet_name='Sheet1')
sheet2_data = pd.read_excel('procv.xlsx', sheet_name='Sheet2')

# Verificar e remover possíveis valores não-string da coluna 'Cliente' em 'Sheet2'
sheet2_data = sheet2_data[sheet2_data['Cliente'].apply(lambda x: isinstance(x, str))]

# Filtrar os nomes da "Sheet1" que estão com a coluna "Proposta" vazia
missing_proposals = sheet1_data[sheet1_data['Proposta'].isna()]

# Relacionar os nomes
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

# Criar um DataFrame com os resultados
matched_df = pd.DataFrame(matched_names)

# Salvar os resultados em um novo arquivo Excel
matched_df.to_excel('resultado_relacionado.xlsx', index=False)
