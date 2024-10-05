from datetime import datetime
from db_conexão import get_db_connection
from processo_data import fetch_processes_and_clients
from template import generate_email_body
from mail_sender import send_email
import mysql.connector
import uuid

def main():
    data_do_dia = datetime.now()

    clientes_data = fetch_processes_and_clients()

    try:
        db_connection = get_db_connection()
        db_cursor = db_connection.cursor()

        # Puxar dados de configuração do SMTP
        db_cursor.execute("SELECT smtp_host, smtp_port, smtp_username, smtp_password, url_thumbnail FROM companies LIMIT 1")
        smtp_config = db_cursor.fetchone()

        if smtp_config:
            smtp_host, smtp_port, smtp_user, smtp_password,logo = smtp_config
        else:
            print("Configuração SMTP não encontrada.")
            exit()
    except mysql.connector.Error as err:
        print(f"Erro ao conectar ao banco de dados: {err}")
        exit()

    # SMTP configuration
    smtp_config = (smtp_host, smtp_port, smtp_user, smtp_password, logo)

    for cliente ,processos in clientes_data.items():
        localizador = str(uuid.uuid4()) 
        email_body = generate_email_body(cliente, processos, logo, localizador, data_do_dia)
        email_receiver = "pedroquintas1213@gmail.com"
        #email_receiver= processos[0]['emails']
        bcc_receivers = ""  
        subject = f"LIGCONTATO - DISTRIBUIÇÃO {data_do_dia.strftime('%d/%m/%y')} - {cliente}"

        # Send email
        send_email(smtp_config, email_body, email_receiver, bcc_receivers, subject)

        try:
            for processo in processos:
                processo_id = processo['ID_processo']  
                db_cursor.execute("UPDATE processo SET status = 'S' WHERE ID_processo = %s", (processo_id,))
                db_cursor.execute("""INSERT INTO envio_emails (ID_processo, numero_processo, cod_escritorio, localizador_processo, data_envio, localizador, email_envio, data_hora_envio)
              VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
              """,(processo_id, processo['numero_processo'],processo['cod_escritorio'],processo['localizador'],data_do_dia.strftime('%Y-%m-%d'),
                   localizador,email_receiver, datetime.now()))

            db_connection.commit()  # Commit all changes
            print(f"E-mail enviado para {cliente} e status atualizado para 'S'.")

        except mysql.connector.Error as err:
            print(f"Erro ao atualizar o status ou registrar o envio: {err}")

if __name__ == "__main__":
    main()
