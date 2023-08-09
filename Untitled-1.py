import openpyxl 
import xlrd

from datetime import datetime, timedelta

data_do_dia_anterior = datetime.now() - timedelta(days=1)

data_formatada = data_do_dia_anterior.strftime('%Y%m%d')

nomeDoArquivo= (data_formatada + " - " + "impressao.xls")

# caminho_arquivo = "C:\Users\pedro\OneDrive - LIG CONTATO DIÁRIO FORENSE\DISTRIBUIÇÃO\DISTRIBUIÇÕES\\" + nomeDoArquivo


caminho_arquivo = r"C:\\Users\\pedro\\OneDrive - LIG CONTATO DIÁRIO FORENSE\DISTRIBUIÇÃO\DISTRIBUIÇÕES\\" + nomeDoArquivo

print("Caminho do arquivo:", caminho_arquivo)

try:
    workbook = xlrd.open_workbook(caminho_arquivo)
    print("Arquivo aberto com sucesso.")
except FileNotFoundError:
    print("Arquivo não encontrado.")

workbook = xlrd.open_workbook(caminho_arquivo)

sheet = workbook.active 

last_row = sheet.max_row

dados_linhas = []

for row in range (1,last_row +1):
    cell_value = sheet.cell(row=row ,column =1).value
    if cell_value is not None:
        dados_linhas.append(cell_value)

    print (dados_linhas[1])