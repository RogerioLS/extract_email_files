import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from colorama import Fore
from dotenv import load_dotenv

# Importações locais
from source.logger.logger_config import logger_quantum, print_log

# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()


def enviar_email_alerta(contagem_nan: int, limite: int):
    """
    Envia um e-mail de alerta sobre a baixa qualidade dos
    dados em uma planilha.

    Esta função é acionada quando a contagem de valores nulos (NaN) em uma verificação
    excede o limite pré-configurado. Ela lê as credenciais de e-mail e as configurações
    do servidor SMTP a partir de variáveis de ambiente e envia um e-mail formatado
    em HTML para o destinatário especificado.

    Args:
            contagem_nan (int): O número de valores NaN que foram encontrados.
            limite (int): O limite máximo de valores NaN que era permitido.

    Raises:
            - smtplib.SMTPAuthenticationError: Se as credenciais (usuário/senha) forem inválidas.
            - Exception: Para outros erros relacionados à conexão ou envio do e-mail.
    """
    theme_color = Fore.RED  # Vermelho para indicar alerta
    print_log(
        "AÇÃO", "Iniciando a preparação do e-mail de alerta...", theme_color=theme_color
    )

    # --- CONFIGURAÇÕES DO E-MAIL ---
    email_remetente = os.getenv("EMAIL_USER")
    senha_remetente = os.getenv("EMAIL_PASSWORD")
    email_destinatario = os.getenv("EMAIL_DESTINATARIO")

    if not all([email_remetente, senha_remetente, email_destinatario]):
        error_msg = (
            "As variáveis de ambiente EMAIL_USER, EMAIL_PASSWORD ou"
            " EMAIL_DESTINATARIO não foram encontradas."
        )
        print_log("ERROR", error_msg)
        logger_quantum.error(error_msg)
        return

    logger_quantum.info(
        "Credenciais e destinatário carregados. Preparando para enviar alerta para"
        f" {email_destinatario}."
    )

    # Configurações do servidor SMTP (Outlook)
    smtp_server = "smtp.office365.com"
    smtp_port = 587

    # --- CRIAÇÃO DA MENSAGEM ---
    msg = MIMEMultipart()
    msg["From"] = email_remetente
    msg["To"] = email_destinatario
    msg["Subject"] = (
        f"⚠️ Alerta de Qualidade de Dados: {contagem_nan} Valores Nulos Encontrados"
    )

    corpo_email = f"""
    <html>
    <body style="font-family: sans-serif;">
        <h2 style="color: #d9534f;">Alerta de Qualidade de Dados</h2>
        <p>Olá,</p>
        <p>A verificação automática detectou um problema na planilha recém-processada.</p>
        <ul style="list-style-type: none; padding: 0;">
            <li style="padding: 5px;"><strong>Valores Nulos/NaN Encontrados:</strong> <span style="color: #d9534f; font-weight: bold;">{contagem_nan}</span></li>
            <li style="padding: 5px;"><strong>Limite Permitido:</strong> {limite}</li>
        </ul>
        <p>A quantidade de dados ausentes excedeu o limite configurado.</p>
        <p><strong>Ação recomendada:</strong> Por favor, verifique a planilha de origem para garantir a integridade dos dados antes de uma nova execução.</p>
        <br>
        <p><em>Este é um e-mail automático.</em></p>
    </body>
    </html>
    """
    msg.attach(MIMEText(corpo_email, "html"))
    logger_quantum.info("Corpo do e-mail de alerta construído.")

    # --- ENVIO DO E-MAIL ---
    try:
        print_log(
            "INFO",
            f"Conectando ao servidor SMTP ({smtp_server})...",
            theme_color=theme_color,
        )
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            logger_quantum.info("Conexão TLS estabelecida.")
            server.login(email_remetente, senha_remetente)
            logger_quantum.info("Login no servidor SMTP realizado com sucesso.")
            server.send_message(msg)
            success_msg = (
                f"E-mail de alerta enviado com sucesso para {email_destinatario}!"
            )
            print_log("INFO", success_msg, theme_color=Fore.GREEN)
            logger_quantum.info(success_msg)

    except smtplib.SMTPAuthenticationError as e:
        error_msg = (
            "Erro de autenticação SMTP. Verifique seu e-mail e senha (ou senha de app)."
        )
        print_log("ERROR", error_msg)
        logger_quantum.error(f"{error_msg} | Exception: {e}", exc=e)

    except Exception as e:
        error_msg = f"Falha inesperada ao enviar o e-mail: {e}"
        print_log("ERROR", error_msg)
        logger_quantum.error(error_msg, exc=e)
