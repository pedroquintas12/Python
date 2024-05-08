import requests
from bs4 import BeautifulSoup
from flask import Flask, request
import os
import time

app = Flask(__name__)

def baixar_pdf(pdf_url, pdf_path):
    tentativas = 5
    while tentativas > 0:
        try:
            with open(pdf_path, "wb") as pdf_file:
                pdf_response = requests.get(pdf_url)
                pdf_file.write(pdf_response.content)
            print(f"PDF {pdf_url} baixado com sucesso.")    
            return True
        except Exception as e:
            print(f"Erro ao baixar o PDF {pdf_path}: {e}")
            tentativas -= 1
            if tentativas > 0:
                print("Tentando novamente...")
                time.sleep(1)
    return False

def extrair_links_da_pagina(url):
    response = requests.get(url)

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, "html.parser")
        
        links = soup.find_all("a", class_="summary url")
        
        return [link.get("href").replace("/view", "") for link in links]
    else:
        return None

@app.route('/')
def extrair_links_e_baixar_pdfs():
    url = request.args.get('url')
    
    response = requests.get(url)

    pasta_pdfs = r"C:\Users\pedro\OneDrive - LIG CONTATO DIÁRIO FORENSE\PDFS PARAIBA"

    if response.status_code == 200:
        soup = BeautifulSoup(response.content, "html.parser")
        
        next_page_link = soup.find("a", class_="proximo")
        
        links = extrair_links_da_pagina(url)
        
        if next_page_link:
            print("2 pagina")
            next_page_url = next_page_link.get("href")
            next_page_links = extrair_links_da_pagina(next_page_url)
            if next_page_links:
                links.extend(next_page_links)
        
        if not os.path.exists("pdfs"):
            os.makedirs("pdfs")
        
        for link in links:
            pdf_name = link.split("/")[-1]
            pdf_path = os.path.join(pasta_pdfs, pdf_name)
            if not baixar_pdf(link, pdf_path):
                return {'error': f'Erro ao baixar o PDF {pdf_name}'}
        
        return {'message': 'PDFs baixados com sucesso na pasta "pdfs"'}
    else:
        return {'error': 'Erro ao fazer a requisição: {}'.format(response.status_code)}

if __name__ == '__main__':
    app.run(debug=True)
