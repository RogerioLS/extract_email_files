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


def enviar_email_sucesso():
    """
    Envia um e-mail de confirmação após um processamento bem-sucedido.

    Esta função é chamada quando a verificação de qualidade dos dados passa e os dados
    são salvos corretamente. Ela informa ao destinatário que o processo foi concluído
    sem problemas. As credenciais e o destinatário são lidos de variáveis de ambiente.
    """
    theme_color = Fore.GREEN  # Verde para indicar sucesso
    print_log(
        "AÇÃO",
        "Iniciando a preparação do e-mail de sucesso...",
        theme_color=theme_color,
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
        "Credenciais e destinatário carregados. Preparando para enviar e-mail de"
        f" sucesso para {email_destinatario}."
    )

    # Configurações do servidor SMTP (Outlook)
    smtp_server = "smtp.office365.com"
    smtp_port = 587

    # --- CRIAÇÃO DA MENSAGEM ---
    msg = MIMEMultipart()
    msg["From"] = email_remetente
    msg["To"] = email_destinatario
    msg["Subject"] = "✅ Processo Concluído com Sucesso"

    corpo_email = """
    <html>
    <body style="font-family: sans-serif;">
        <h2 style="color: #5cb85c;">Relatório de Qualidade de Dados</h2>
        <p>Olá,</p>
        <p>A verificação automática da planilha e o salvamento dos dados foram concluídos com <strong>sucesso</strong>.</p>
        <p>Nenhum problema que exigisse atenção imediata foi detectado e os dados foram atualizados no destino.</p>
        <br>
        <p><em>Este é um e-mail automático.</em></p>
    </body>
    </html>
    """
    msg.attach(MIMEText(corpo_email, "html"))
    logger_quantum.info("Corpo do e-mail de sucesso construído.")

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
            success_msg = f"E-mail de sucesso enviado para {email_destinatario}!"
            print_log("INFO", success_msg, theme_color=theme_color)
            logger_quantum.info(success_msg)

    except smtplib.SMTPAuthenticationError as e:
        error_msg = "Erro de autenticação SMTP. Verifique seu e-mail e senha."
        print_log("ERROR", error_msg)
        logger_quantum.error(f"{error_msg} | Exception: {e}", exc=e)

    except Exception as e:
        error_msg = f"Falha inesperada ao enviar o e-mail: {e}"
        print_log("ERROR", error_msg)
        logger_quantum.error(error_msg, exc=e)
