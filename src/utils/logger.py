import logging
import os
from datetime import datetime
from logging.handlers import RotatingFileHandler
from src.config.settings import settings

config = settings


class Logger:
    def __init__(self):
        # Definir o diretório de logs
        self.log_dir = config.paths.logs
        # Criar o diretório de logs se não existir
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)

        self.dia_mes_ano = datetime.today().strftime('%d_%m_%Y')

        # Criar logger principal
        self.logger = logging.getLogger('MiddlewareGraphAPI')

        # Verificar se o logger já foi configurado (evitar adicionar handlers duplicados)
        if not self.logger.hasHandlers():

            self.logger.setLevel(logging.DEBUG)  # Permite registrar todos os níveis de log

            # Configurar o caminho do arquivo de log no novo diretório
            log_file_path = os.path.join(self.log_dir, f"log_{self.dia_mes_ano}.txt")

            # Configurar handler para arquivo com rotação
            file_handler = RotatingFileHandler(
                log_file_path, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
            )
            file_handler.setLevel(logging.DEBUG)  # Registra todos os logs no arquivo

            # Configurar handler para console
            console_handler = logging.StreamHandler()
            
            console_handler.setLevel(logging.INFO)

            # Formatar mensagens de log
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s - %(module)s - %(funcName)s - Line: %(lineno)d',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            file_handler.setFormatter(formatter)
            console_handler.setFormatter(formatter)

            # Adicionar handlers ao logger
            self.logger.addHandler(file_handler)
            self.logger.addHandler(console_handler)

    def get_logger(self):
        """
        Retorna o logger configurado para uso externo.
        """
        return self.logger

# --- Teste ---
if __name__ == "__main__":
    logger = Logger().get_logger()