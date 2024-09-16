import pandas as pd
import mysql.connector

# Conectando ao banco de dados MySQL
db_connection = mysql.connector.connect(
    host="26.191.28.12",
    user="pedro",
    password="123456",
    database="apidistribuicao"
)

# Função para verificar dados que faltam
def verificar_dados_faltantes(excel_path, sheet_name, column_name, table_name, column_name_db):
    # Lendo os dados do arquivo Excel
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    
    # Pegando os dados da coluna específica
    dados_excel = df[column_name].dropna().unique().tolist()
    
    # Criando um cursor para executar a consulta no MySQL
    cursor = db_connection.cursor()
    
    # Verificando quais dados do Excel não estão na base de dados
    dados_faltantes = []
    for dado in dados_excel:
        query = f"SELECT COUNT(*) FROM {table_name} WHERE {column_name_db} = %s"
        cursor.execute(query, (dado,))
        result = cursor.fetchone()
        if result[0] == 0:
            dados_faltantes.append(dado)
    
    cursor.close()
    return dados_faltantes

# Especificando os parâmetros
excel_path = 'C:\\Users\\pedro\\Downloads\\impressao.xls'
sheet_name = 'Worksheet'
column_name = 'NUMERO PROCESSO'
table_name = 'processo'
column_name_db = 'numero_processo'

# Verificando os dados faltantes
dados_faltantes = verificar_dados_faltantes(excel_path, sheet_name, column_name, table_name, column_name_db)

if dados_faltantes is not None:
    # Exibindo os dados que estão faltando na base de dados
    print("Dados faltantes na base de dados:", dados_faltantes)
else:
    print("todos os dados cadastrados!")

# Fechando a conexão com o banco de dados
db_connection.close()
