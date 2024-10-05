import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_email(smtp_config, email_body, email_receiver, bcc_receivers, subject):
    smtp_host, smtp_port, smtp_user, smtp_password, logo = smtp_config



    # Envio do e-mail
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = email_receiver
    msg['Bcc'] = bcc_receivers
    msg['Subject'] = subject
    msg.attach(MIMEText(email_body, 'html'))

    # Envio do e-mail usando SMTP
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls() 
        server.login(smtp_user, smtp_password)
        server.send_message(msg)
