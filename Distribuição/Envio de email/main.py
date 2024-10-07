from datetime import datetime
from db_conexão import get_db_connection
from processo_data import fetch_processes_and_clients
from template import generate_email_body
from mail_sender import send_email
import mysql.connector
import uuid
import time
import schedule

def enviar_emails():
    data_do_dia = datetime.now()

    # Busca os dados dos clientes e processos
    clientes_data = fetch_processes_and_clients()

    # Calcula o total de escritórios e total de processos por escritório
    total_escritorios = len(clientes_data)  # Total de escritórios
    total_processos_por_escritorio = {cliente: len(processos) for cliente, processos in clientes_data.items()}

    # Exibe os totais
    print(f"Total de escritórios a serem enviados: {total_escritorios}")
    for cliente, total_processos in total_processos_por_escritorio.items():
        print(f"Escritório: {cliente} - Total de processos: {total_processos}")

    try:
        db_connection = get_db_connection()
        db_cursor = db_connection.cursor()

        # Puxar dados de configuração do SMTP
        db_cursor.execute("SELECT smtp_host, smtp_port, smtp_username, smtp_password, url_thumbnail FROM companies LIMIT 1")
        smtp_config = db_cursor.fetchone()

        if smtp_config:
            smtp_host, smtp_port, smtp_user, smtp_password, logo = smtp_config
        else:
            print("Configuração SMTP não encontrada.")
            exit()
    except mysql.connector.Error as err:
        print(f"Erro ao conectar ao banco de dados: {err}")
        exit()

    # Configuração do SMTP
    smtp_config = (smtp_host, smtp_port, smtp_user, smtp_password, logo)

    for cliente, processos in clientes_data.items():
        localizador = str(uuid.uuid4()) 
        email_body = generate_email_body(cliente, processos, logo, localizador, data_do_dia)
        email_receiver = "pedroquintas1213@gmail.com"
        #email_receiver = processos[0]['emails']
        bcc_receivers = ""  
        subject = f"LIGCONTATO - DISTRIBUIÇÃO {data_do_dia.strftime('%d/%m/%y')} - {cliente}"

        # Envia o e-mail
        send_email(smtp_config, email_body, email_receiver, bcc_receivers, subject)

        try:
            for processo in processos:
                processo_id = processo['ID_processo']  
                db_cursor.execute("UPDATE processo SET status = 'S' WHERE ID_processo = %s", (processo_id,))
                db_cursor.execute("""
                    INSERT INTO envio_emails (ID_processo, numero_processo, cod_escritorio, localizador_processo, data_envio, localizador, email_envio, data_hora_envio)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """, (processo_id, processo['numero_processo'], processo['cod_escritorio'], processo['localizador'], 
                      data_do_dia.strftime('%Y-%m-%d'), localizador, email_receiver, datetime.now()))

            db_connection.commit()  # Commit all changes
            print(f"E-mail enviado para {cliente} {datetime.now().strftime('%H:%M')}.")

        except mysql.connector.Error as err:
            print(f"Erro ao atualizar o status ou registrar o envio: {err}")

# Agenda o envio para todos os dias às 16:00
schedule.every().day.at("21:20").do(enviar_emails)

if __name__ == "__main__":
    while True:
        schedule.run_pending()

            # A cada hora, recalcula os totais e exibe na tela
        clientes_data = fetch_processes_and_clients()
        total_escritorios = len(clientes_data)  
        total_processos_por_escritorio = {cliente: len(processos) for cliente, processos in clientes_data.items()}

            # Atualiza a exibição
        print(f"\nAguardando o horário de envio... (Atualizado: {datetime.now().strftime('%d-%m-%y %H:%M')})")
        print(f"Total de escritórios a serem enviados: {total_escritorios}")
        for cliente, total_processos in total_processos_por_escritorio.items():
            print(f"Escritório: {cliente} - Total de processos: {total_processos}")

        time.sleep(3600)  # Verifica a cada hora