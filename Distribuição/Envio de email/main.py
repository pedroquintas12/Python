from datetime import datetime
from db_conexão import get_db_connection
from processo_data import fetch_processes_and_clients
from template import generate_email_body
from mail_sender import send_email
import mysql.connector
import uuid
import time
import schedule
from logger_config import logger


def enviar_emails():

    
    data_do_dia = datetime.now()

    # Busca os dados dos clientes e processos
    clientes_data = fetch_processes_and_clients()
    total_escritorios = len(clientes_data)  
    total_processos_por_escritorio = {cliente: len(processos) for cliente, processos in clientes_data.items()}

    try:
        db_connection = get_db_connection()
        db_cursor = db_connection.cursor()

        # Puxar dados de configuração do SMTP
        db_cursor.execute("""SELECT smtp_host, smtp_port, smtp_username, smtp_password, smtp_from_email, 
                          smtp_from_name,smtp_reply_to, smtp_cc_emails,smtp_bcc_emails, url_thumbnail FROM companies LIMIT 1""")
        smtp_config = db_cursor.fetchone()

        if smtp_config:
            smtp_host, smtp_port, smtp_user, smtp_password,smtp_from_email,smtp_from_name,smtp_reply_to,smtp_cc_emails,smtp_bcc_emails,logo = smtp_config
        else:
            logger.warning("Configuração SMTP não encontrada.")
            exit()
    except mysql.connector.Error as err:
        logger.error(f"Erro ao conectar ao banco de dados: {err}")
        exit()

    # Configuração do SMTP
    smtp_config = (smtp_host, smtp_port, smtp_user, smtp_password,smtp_from_email,smtp_from_name,smtp_reply_to,smtp_cc_emails,smtp_bcc_emails,logo)

    for cliente, processos in clientes_data.items():
        localizador = str(uuid.uuid4()) 
        email_body = generate_email_body(cliente, processos, logo, localizador, data_do_dia)
        email_receiver= processos[0]['emails']
        bcc_receivers = smtp_bcc_emails 
        cc_receiver = smtp_cc_emails
        subject = f"LIGCONTATO - DISTRIBUIÇÃO {data_do_dia.strftime('%d/%m/%y')} - {cliente}"

        # Envia o e-mail
        send_email(smtp_config, email_body, email_receiver, bcc_receivers,cc_receiver, subject)

        try:
            for processo in processos:
                processo_id = processo['ID_processo']  
                db_cursor.execute("UPDATE processo SET status = 'S' WHERE ID_processo = %s", (processo_id,))
                db_cursor.execute("""INSERT INTO envio_emails (ID_processo, numero_processo, 
                                  cod_escritorio, localizador_processo, data_envio, localizador, email_envio, data_hora_envio)
                                  VALUES (%s, %s, %s, %s, %s, %s, %s, %s)""",
                                    (processo_id, processo['numero_processo'], processo['cod_escritorio'], processo['localizador'], 
                      data_do_dia.strftime('%Y-%m-%d'), localizador, email_receiver, datetime.now()))

            db_connection.commit()  # Commit all changes
            for cliente, total_processos in total_processos_por_escritorio.items():
                logger.info(f"E-mail enviado para {cliente} AS {datetime.now().strftime('%H:%M:%S')} - total: {total_processos.items}.")

        except mysql.connector.Error as err:
            logger.error(f"Erro ao atualizar o status ou registrar o envio: {err}")

#Agenda o envio para todos os dias às 16:00
schedule.every().day.at("16:16").do(enviar_emails)

if __name__ == "__main__":

    while True:
        schedule.run_pending()
            # A cada hora, recalcula os totais e exibe na tela
        clientes_data = fetch_processes_and_clients()
        total_escritorios = len(clientes_data)  
        total_processos_por_escritorio = {cliente: len(processos) for cliente, processos in clientes_data.items()}

            # Atualiza a exibição
        logger.info(f"\nAguardando o horário de envio... (Atualizado: {datetime.now().strftime('%d-%m-%y %H:%M')})")
        logger.info(f"Total de escritórios a serem enviados: {total_escritorios}")
        for cliente, total_processos in total_processos_por_escritorio.items():
            logger.info(f"Escritório: {cliente} - Total de processos: {total_processos}")
        time.sleep(10)
