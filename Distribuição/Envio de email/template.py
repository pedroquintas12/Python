def generate_email_body(cliente, processos, logo, localizador, data_do_dia):
    email_body = ""
    total_processos = len(processos)

    for idx, processo in enumerate(processos, start=1):
        email_body += f"""
            <div class="processo">
                <p><strong>Distribuição {idx} de {total_processos}</strong></p> </br>
                <p>Tribunal: {processo['tribunal']}</p></br>
                <p>UF/Instância/Comarca: {processo['uf']}/{processo['instancia']}/{processo['comarca']}</p></br>
                <p>Número do Processo: {processo['numero_processo']}</p></br>
                <p>Data de Distribuição: {processo['data_distribuicao']}</p></br>
                <p>Órgão: {processo['orgao']}</p></br>
                <p>Classe Judicial: {processo['classe_judicial']}</p></br>
                <p>Polo Ativo: {processo['polo_ativo']}</p></br>
                <p>Polo Passivo: {processo['polo_passivo']}</p></br>
                <div class="links">
                    <p><strong>Links:</strong></p></br>
        """
        for link_info in processo['links']:
            email_body += f'<p>({link_info["tipoLink"]}): <a href="{link_info["link_doc"]}">{processo["tipo_processo"]}({link_info["id_link"]})</a></p></br>'
        email_body += "</div></div>"

    email_body = f"""
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>{cliente}</title>
            <style>
                        body {{
                            font-family: Arial, sans-serif; 
                            margin: 0; 
                            padding: 0; 
                            background-color: #f4f4f4; 
                            color: #333; 
                            line-height: 1.6; 
                        }}
                        .container {{
                            padding: 20px; 
                            background-color: #fff; 
                            margin: 20px auto; 
                            border-radius: 8px; 
                            box-shadow: 0 4px 8px rgba(0,0,0,0.1); 
                            max-width: 800px; 
                        }}
                        .header {{
                            display: flex;
                            align-items: center;
                            justify-content: space-between;
                            padding: 10px 0;
                            border-bottom: 1px solid #ccc;
                        }}
                        .header img {{
                            max-height: 80px; 
                            margin-right: 20px;
                        }}
                        .header div {{
                            flex-grow: 1; 
                        }}
                        .processo {{
                            border: 1px solid #e0e0e0; 
                            border-radius: 8px; 
                            padding: 15px; 
                            margin-bottom: 20px; 
                            background-color: #f9f9f9; 
                        }}
                        .processo p {{
                            margin: 5px 0;
                        }}
                        .links {{
                            margin-top: 10px; 
                        }}
                        .alert {{
                            background-color: #ff4d4d; 
                            color: #fff;
                            padding: 15px;
                            border-radius: 8px;
                            margin-bottom: 20px;
                            font-weight: bold;
                        }}
                        .processo p{{
                            margin-bottom: 10px ;
                        }}
              </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <img src="{logo}" alt="Logo">
                    <div>
                        <h1>{cliente}</h1>
                        <p>Data: {data_do_dia.strftime('%d/%m/%y')}</p>
                        <p>Localizador: {localizador}</p>
                        <p>Total    Distribuições: {total_processos}</p>
                    </div>
                </div>
                <div class="alert">
                    *Atenção* Esta mensagem pode conter mais conteúdo no corpo do e-mail, portanto verifique no final da mensagem se existe a opção de "Exibir toda a mensagem" para visualizar mais conteúdo.
                </div>
                <div class="content">
                    {email_body}
                </div>
            </div>
        </body>
        </html>
    """
    return email_body
