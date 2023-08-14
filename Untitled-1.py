import xlrd
import openpyxl
import pandas as pd
from docx import Document
from datetime import datetime, timedelta

data_do_dia_anterior = datetime.now() - timedelta(days=1)
data_formatada = data_do_dia_anterior.strftime('%Y%m%d')
nomeDoArquivo = (data_formatada + "-" + "impressao.xls")
nomeDoArquivoDocx = (data_formatada + "-" + "distribuição.docx")

caminho_arquivo = r"C:\Users\pedro\OneDrive - LIG CONTATO DIÁRIO FORENSE\DISTRIBUIÇÃO\DISTRIBUIÇÕES\\" + nomeDoArquivo

print("Caminho do arquivo:", caminho_arquivo)

try:
    workbook = xlrd.open_workbook(caminho_arquivo)
    sheet = workbook.sheet_by_index(0)  
    print("Planilha carregada com sucesso.")

    # Colunas para extrair 
    colunas_alvo = [0, 2, 4, 5, 7, 10, 12]  

    dados_das_linhas = []

    for row in range(1, sheet.nrows):  # Começando da segunda linha
        linha = {}
        for col_idx, col in enumerate(colunas_alvo):
            valor_celula = sheet.cell_value(row, col)
            nome_coluna = f'coluna_{col + 1}'
        
                
            linha[nome_coluna] = valor_celula
        dados_das_linhas.append(linha)
  
  #cria o doc word
    doc = Document()

# Adiciona informações ao documento
    doc.add_heading('Distribuição', level=0,)

# Corpo do e-mail
    email_template = (
    "Prezado(a),\n\n"
    "Segue abaixo o relatório de distribuição referente ao dia anterior:\n\n"
    "Detalhes:\n"
    "Data: {data}\n"
    "Dados coletados:\n\n"
    "{dados}\n\n"
    "Atenciosamente,\n"
    "Lig Contato"
)

# Preencher o corpo do e-mail com os dados capturados
    dados_formatados = ""
    contador = 0
    for linha in dados_das_linhas:
        contador += 1
        dados_formatados += f"Distribuição: {contador}\n\n\n"
        dados_formatados += f"Data da distribuição: {linha['coluna_1']}\n\n"
        dados_formatados += f"Numero Processo: {linha['coluna_3']}\n\n"
        dados_formatados += f"Polo Passivo: {linha['coluna_5']}\n\n"
        dados_formatados += f"Polo Ativo: {linha['coluna_6']}\n\n"
        dados_formatados += f"Orgão: {linha['coluna_8']}\n\n"
        dados_formatados += f"Link: {linha['coluna_11']}\n\n"
        dados_formatados += f"Classe Judicial: {linha['coluna_13']}\n\n"
       
    

    email_texto = email_template.format(data=data_do_dia_anterior.strftime('%d/%m/%y'),
                                    dados=dados_formatados)

    doc.add_paragraph(email_texto)

    caminho_arquivo_docx = r"C:\Users\pedro\OneDrive - LIG CONTATO DIÁRIO FORENSE\Envio Whatsapp\Documentos\\" + nomeDoArquivoDocx
    doc.save(caminho_arquivo_docx)

except FileNotFoundError:
    print("Arquivo Excel não encontrado.")
