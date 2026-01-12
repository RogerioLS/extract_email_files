import json
import os

import requests
import urllib3
from colorama import Fore
from dotenv import load_dotenv

# Importações locais
from source.logger.logger_config import logger_quantum, print_log

# Desabilita avisos de certificado inseguro
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Carrega as variáveis de ambiente
load_dotenv()


def enviar_teams_sucesso():
    """
    Envia uma confirmação de sucesso para o Microsoft Teams via Webhook.
    """
    theme_color = Fore.GREEN
    print_log(
        "AÇÃO",
        "Preparando envio de confirmação para o Teams...",
        theme_color=theme_color,
    )

    webhook_url = os.getenv("TEAMS_WEBHOOK_URL")

    if not webhook_url:
        error_msg = "A variável de ambiente TEAMS_WEBHOOK_URL não foi encontrada."
        print_log("ERROR", error_msg)
        logger_quantum.error(error_msg)
        return

    # Card do Teams
    card_data = {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "themeColor": "5cb85c",  # Verde
        "summary": "Processo Concluído com Sucesso",
        "sections": [
            {
                "activityTitle": "✅ Processo Concluído com Sucesso",
                "activitySubtitle": "Quantum - Automated Data Check",
                "text": "A verificação automática da planilha e o salvamento dos dados foram concluídos sem problemas. Os dados estão íntegros.",
            }
        ],
    }

    try:
        response = requests.post(
            webhook_url,
            data=json.dumps(card_data),
            headers={"Content-Type": "application/json"},
            verify=False,
        )
        response.raise_for_status()

        success_msg = "Confirmação de sucesso enviada para o Teams!"
        print_log("INFO", success_msg, theme_color=theme_color)
        logger_quantum.info(success_msg)

    except requests.exceptions.RequestException as e:
        error_msg = f"Falha ao enviar confirmação para o Teams: {e}"
        print_log("ERROR", error_msg)
        logger_quantum.error(error_msg, exc=e)
