import re
import requests
from bs4 import BeautifulSoup
from flask import Flask, request
import os
import time

app = Flask(__name__)

def baixar_csv(csv_url, csv_path):
    tentativas = 5
    while tentativas > 0:
        try:
            with open(csv_path, "wb") as csv_file:
                csv_response = requests.get(csv_url)
                csv_file.write(csv_response.content)
            print(f"CSV {csv_url} baixado com sucesso.")    
            return True
        except Exception as e:
            print(f"Erro ao baixar o CSV {csv_path}: {e}")
            tentativas -= 1
            if tentativas > 0:
                print("Tentando novamente...")
                time.sleep(1)
    return False

def extrair_links_da_pagina(url):
    response = requests.get(url)

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, "html.parser")
        
        # Encontrar todos os links com a classe "resource-url-analytics"
        links = soup.find_all("a", class_="resource-url-analytics")
        
        if not links:
            print("Nenhum link encontrado com a classe fornecida.")
        
        csv_links = [link.get("href") for link in links if link.get("href") and link.get("href").endswith("/dados?formato=csv")]
        
        return csv_links
    else:
        return None

def limpar_nome_arquivo(nome):
    nome_limpo = re.sub(r'[<>:"/\\|?*]', '_', nome)
    return nome_limpo

@app.route('/')
def extrair_links_e_baixar_csvs():
    url = request.args.get('url')
    
    pasta_csvs = r"C:\Users\pedro\OneDrive\Documentos\teste csv"

    links = extrair_links_da_pagina(url)
    
    if links:
        if not os.path.exists(pasta_csvs):
            os.makedirs(pasta_csvs)
        
        for link in links:
            if not link.startswith("http"):
                link = f"https://api.bcb.gov.br{link}"
            
            csv_name = limpar_nome_arquivo(link.split("/")[-1])
            
            if not csv_name.endswith(".csv"):
                csv_name += ".csv"
            
            csv_path = os.path.join(pasta_csvs, csv_name)
            
            # Baixar o arquivo CSV
            if not baixar_csv(link, csv_path):
                return {'error': f'Erro ao baixar o CSV {csv_name}'}
        
        return {'message': 'CSVs baixados com sucesso na pasta "csv"'}
    else:
        return {'error': 'Erro ao extrair os links ou página inválida'}

if __name__ == '__main__':
    app.run(debug=True)
