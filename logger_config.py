import logging
import os
from datetime import datetime

# Configurando o log
log_directory = 'C:\\Users\\pedro\\OneDrive - LIG CONTATO DIÁRIO FORENSE\\DISTRIBUIÇÃO\\LOG DISTRIBUICÃO\\'

# Cria o diretório se não existir
if not os.path.exists(log_directory):
    os.makedirs(log_directory)


current_date = datetime.now().strftime('%y-%m-%d')
log_name = f'log-envio-{current_date}.log'

log_file_path = os.path.join(log_directory,log_name)

logger = logging.getLogger('Log')
logger.setLevel(logging.DEBUG)

file_handler = logging.FileHandler(log_file_path)
console_handler = logging.StreamHandler()

formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

logger.addHandler(file_handler)
logger.addHandler(console_handler)

# Mensagem inicial no log
logger.info("Início do programa.")
