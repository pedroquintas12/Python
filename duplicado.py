import pandas as pd

# Carregar o arquivo Excel
arquivo_excel = r'C:\Users\pedro\OneDrive\Documentos\GERENCIA DE MOTO.xlsm'
df = pd.read_excel(arquivo_excel)

# Verificar valores duplicados em uma coluna específica (exemplo: 'nome_da_coluna')
coluna = 'PLACA'
duplicados = df[df.duplicated(coluna)]

# Exibir as linhas que têm valores duplicados
if not duplicados.empty:
    print("Valores duplicados encontrados:")
    print(duplicados)
else:
    print("Não há valores duplicados na coluna.")
