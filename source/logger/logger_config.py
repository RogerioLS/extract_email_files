import atexit
import json
import os
import traceback
from datetime import datetime
from pathlib import Path

from colorama import Fore, Style
from dotenv import load_dotenv

load_dotenv()

# Carrega as variáveis de ambiente necessárias
PASTA_LOG = os.getenv("PASTA_LOG")

# --- FUNÇÃO DE APRESENTAÇÃO PARA O TERMINAL ---
LOG_COLORS = {
    "INFO": Fore.GREEN,
    "ERROR": Fore.RED,
    "AVISO": Fore.YELLOW,
    "AÇÃO": Fore.CYAN,
}


def print_log(level: str, message: str, theme_color: str = None):
    """Imprime uma mensagem formatada e colorida no terminal."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # Lógica de seleção de cor corrigida e mais clara
    if level == "ERROR":
        color = LOG_COLORS.get("ERROR")  # Prioridade 1: Erros são sempre vermelhos
    elif theme_color:
        color = theme_color  # Prioridade 2: Usar a cor do tema se ela for fornecida
    else:
        # Prioridade 3: Usar a cor padrão do nível, com branco como fallback
        color = LOG_COLORS.get(level, Fore.WHITE)

    print(
        f"{color}{Style.BRIGHT}[{level}]{Style.RESET_ALL} "
        f"{color}{Style.BRIGHT}[{timestamp}]{Style.RESET_ALL} {message}"
    )


# --- CLASSE DE LOGGING SILENCIOSA PARA ARQUIVOS ---
class Logger:
    """
    Classe de logging silenciosa. Apenas escreve em arquivos JSON (info e error)
    ao final da execução. Não imprime no terminal.
    """

    def __init__(self, log_directory: str, name_prefix: str = "log"):
        """
        Inicializa o Logger e configura os caminhos dos arquivos de log com base na data.

        Args:
            log_directory (str): O diretório base para todos os logs (ex: 'logs').
            name_prefix (str): Um prefixo para os nomes dos arquivos de log (ex: 'quantum').
        """
        self.base_log_dir = Path(log_directory)
        self.name_prefix = name_prefix
        self.info_entries = []
        self.error_entries = []

        # --- Configuração do caminho do log com base na data ---
        now = datetime.now()
        year = now.strftime("%Y")
        month = now.strftime("%m")
        date_str = now.strftime("%Y%m%d")

        # Cria a estrutura de diretórios ano/mês
        todays_log_dir = self.base_log_dir / year / month
        todays_log_dir.mkdir(parents=True, exist_ok=True)

        # Define os caminhos completos para os arquivos de log de info e erro
        self.info_log_path = todays_log_dir / f"{self.name_prefix}_info_{date_str}.json"
        self.error_log_path = (
            todays_log_dir / f"{self.name_prefix}_error_{date_str}.json"
        )

        atexit.register(self.save_logs)

    def _add_log_entry(
        self, level: str, message: str, entry_list: list, extra_data: dict = None
    ):
        """Adiciona uma entrada de log à lista apropriada na memória."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = {"timestamp": timestamp, "level": level, "message": message}
        if extra_data:
            log_entry.update(extra_data)
        entry_list.append(log_entry)

    def info(self, message: str, extra_data: dict = None):
        """Registra uma mensagem de informação na memória."""
        self._add_log_entry("INFO", message, self.info_entries, extra_data=extra_data)

    def error(self, message: str, exc: Exception = None):
        """Registra uma mensagem de erro na memória."""
        extra = {}
        full_message = message
        if exc:
            extra["traceback"] = traceback.format_exc()
            full_message = f"{message} | Exception: {str(exc)}"
        self._add_log_entry("ERROR", full_message, self.error_entries, extra_data=extra)

    def save_logs(self):
        """Salva as listas de logs em seus respectivos arquivos JSON, se não estiverem vazias."""
        if self.info_entries:
            with open(self.info_log_path, "w", encoding="utf-8") as f:
                json.dump(self.info_entries, f, indent=4, ensure_ascii=False)
        if self.error_entries:
            with open(self.error_log_path, "w", encoding="utf-8") as f:
                json.dump(self.error_entries, f, indent=4, ensure_ascii=False)


# Cria uma instância do logger com o diretório de logs e o prefixo
logger_quantum = Logger(PASTA_LOG, name_prefix="quantum")
