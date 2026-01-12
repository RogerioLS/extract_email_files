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


def enviar_teams_alerta(contagem_nan: int, limite: int):
    """
    Envia um alerta para o Microsoft Teams via Webhook sobre a baixa qualidade dos dados.

    Args:
        contagem_nan (int): O número de valores NaN encontrados.
        limite (int): O limite máximo permitido.
    """
    theme_color = Fore.RED
    print_log(
        "AÇÃO", "Preparando envio de alerta para o Teams...", theme_color=theme_color
    )

    webhook_url = os.getenv("TEAMS_WEBHOOK_URL")

    if not webhook_url:
        error_msg = "A variável de ambiente TEAMS_WEBHOOK_URL não foi encontrada."
        print_log("ERROR", error_msg)
        logger_quantum.error(error_msg)
        return

    # Card do Teams (Adaptive Card ou Message Card simples)
    # Usando formato simples de MessageCard para compatibilidade geral com Webhooks
    card_data = {
        "@type": "MessageCard",
        "@context": "http://schema.org/extensions",
        "themeColor": "d9534f",  # Vermelho
        "summary": "Alerta de Qualidade de Dados",
        "sections": [
            {
                "activityTitle": "⚠️ Alerta de Qualidade de Dados",
                "activitySubtitle": "Quantum - Automated Data Check",
                "facts": [
                    {"name": "Status:", "value": "FALHA NA VALIDAÇÃO"},
                    {"name": "Valores Nulos Encontrados:", "value": str(contagem_nan)},
                    {"name": "Limite Permitido:", "value": str(limite)},
                ],
                "text": "A quantidade de dados ausentes na planilha excedeu o limite configurado. Por favor, verifique a planilha de origem.",
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

        success_msg = "Alerta enviado com sucesso para o Teams!"
        print_log("INFO", success_msg, theme_color=Fore.GREEN)
        logger_quantum.info(success_msg)

    except requests.exceptions.RequestException as e:
        error_msg = f"Falha ao enviar alerta para o Teams: {e}"
        print_log("ERROR", error_msg)
        logger_quantum.error(error_msg, exc=e)
