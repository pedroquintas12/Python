def generate_email_body(cliente, processos, logo, localizador, data_do_dia):
    email_body = ""
    total_processos = len(processos)

    
    exibir_max = 1

    for idx, processo in enumerate(processos, start=1):
        email_body += f"""
            <div class="processo">
                <p><strong>Distribuição {idx} de {total_processos}</strong></p> </br>
                <p>Tribunal: {processo['tribunal']}</p></br>
                <p>UF/Instância/Comarca: {processo['uf']}/{processo['instancia']}/{processo['comarca']}</p></br>
                <p>Número do Processo: {processo['numero_processo']}</p></br>
                <p>Data de Distribuição: {processo['data_distribuicao']}</p></br>
                <p>Órgão: {processo['orgao']}</p></br>
                <p>Classe Judicial: {processo['classe_judicial']}</p></br>"""
        
        autor_list = processo['autor']
        total_autores = len(autor_list)
        if total_autores > exibir_max:
            autores_exibidos = ', '.join(autor['nomeAutor'] for autor in autor_list[:exibir_max])
            email_body += f'<p>Polo Ativo: {autores_exibidos} (+{total_autores - exibir_max})</p></br>'
        else:
            autores_exibidos = ', '.join(autor['nomeAutor'] for autor in autor_list)
            email_body += f'<p>Polo Ativo: {autores_exibidos}</p></br>'

        # Processa réus
        reu_list = processo['reu']
        total_reus = len(reu_list)
        if total_reus > exibir_max:
            reus_exibidos = ', '.join(reu['nomeReu'] for reu in reu_list[:exibir_max])
            email_body += f'<p>Polo Passivo: {reus_exibidos} (+{total_reus - exibir_max})</p></br>'
        else:
            reus_exibidos = ', '.join(reu['nomeReu'] for reu in reu_list)
            email_body += f'<p>Polo Passivo: {reus_exibidos}</p></br>'
        """    
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
                .footer {{
                    background-color: #000; 
                    color: #ffffff; 
                    padding: 10px; 
                    text-align: justify; 
                    font-size: 8px; 
                    line-height: 1.2; 
                    font-family: Arial, 'Helvetica Neue', Helvetica, sans-serif;
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
                        <p>Total Distribuições: {total_processos}</p>
                    </div>
                </div>
                <div class="alert">
                    *Atenção* Esta mensagem pode conter mais conteúdo no corpo do e-mail, portanto verifique no final da mensagem se existe a opção de "Exibir toda a mensagem" para visualizar mais conteúdo.
                </div>
                <div class="content">
                    {email_body}
                </div>
                <!-- FOOTER -->
                <div class="footer">
                    <p>
                        Esta mensagem constitui informação privilegiada e confidencial, legalmente
                        resguardada por segredo profissional, nos termos do art. 7º, inc. II, e ss. da lei nº 8.906/94,
                        referindo-se exclusivamente ao relacionamento pessoal e profissional entre o remetente e o
                        destinatário, sendo vedada a utilização, divulgação ou reprodução do seu conteúdo.
                    </p>
                    <!--
                    <p style="text-align: center;">
                        Localizador: {localizador}>--</span>
                    </p>
                    -->
                </div>
            </div>
        </body>
        </html>
    """
    return email_body
