import mysql.connector

db_config = {
    'host': '26.191.28.12',
    'port': '3306',
    'user': 'pedro',
    'password': '123456',
    'database': 'apidistribuicao'
}

def get_db_connection():
    return mysql.connector.connect(**db_config)
