import os
import sys
import time
from datetime import datetime
from pathlib import Path

import win32com.client
from colorama import Fore

# Importações locais padronizadas
from source.logger.logger_config import logger_quantum, print_log

# Ajusta o PATH se estiver rodando como um executável PyInstaller
if getattr(sys, "frozen", False):
    dll_path = os.path.join(sys._MEIPASS, "libs")
    os.environ["PATH"] = dll_path + os.pathsep + os.environ.get("PATH", "")


def inicializar_outlook():
    """
    Tenta estabelecer uma conexão com a aplicação Outlook.

    Se a conexão inicial falhar, a função tentará forçar o encerramento do processo
    do Outlook, reiniciá-lo e tentar a conexão novamente.

    Returns:
            win32com.client.Dispatch: O objeto 'namespace' do Outlook em caso de sucesso,
                                                            ou None se a conexão falhar mesmo após a reinicialização.
    """
    theme_color = Fore.BLUE
    print_log("AÇÃO", "Conectando ao Outlook...", theme_color=theme_color)
    try:
        outlook_app = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook_app.GetNamespace("MAPI")
        _ = namespace.Folders  # Acesso para forçar a inicialização
        print_log(
            "INFO",
            "Conexão com o Outlook estabelecida com sucesso.",
            theme_color=theme_color,
        )
        logger_quantum.info("Conexão com Outlook bem-sucedida.")
        return namespace
    except Exception as e:
        print_log(
            "AVISO",
            f"Falha ao conectar ao Outlook: {e}. Tentando reiniciar o aplicativo.",
            theme_color=Fore.YELLOW,
        )
        logger_quantum.error(f"Falha inicial ao conectar ao Outlook: {e}", exc=e)
        os.system("taskkill /f /im outlook.exe > nul 2>&1")
        time.sleep(2)
        os.startfile("outlook.exe")
        time.sleep(8)
        try:
            outlook_app = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook_app.GetNamespace("MAPI")
            _ = namespace.Folders
            print_log(
                "INFO",
                "Conexão com o Outlook restabelecida após reinicialização.",
                theme_color=theme_color,
            )
            logger_quantum.info("Conexão com Outlook restabelecida.")
            return namespace
        except Exception as e_after_restart:
            error_msg = (
                "Não foi possível conectar ao Outlook mesmo após reiniciar:"
                f" {e_after_restart}"
            )
            print_log("ERROR", error_msg)
            logger_quantum.error(error_msg, exc=e_after_restart)
            return None


def extrair_excel_email(pasta_raiz_quantum: str, headline_prefix: str):
    """
    Busca e-mails recentes no Outlook, encontra um com um assunto específico
    e salva seu anexo .xlsx.

    A função se conecta à conta 'asa.com.br' no Outlook, procura na Caixa de Entrada
    pelos e-mails mais recentes recebidos no dia atual que correspondam ao
    'headline_prefix'. Ao encontrar o primeiro e-mail correspondente, salva seu
    anexo .xlsx na pasta raiz especificada.

    Args:
            pasta_raiz_quantum (str): O caminho da pasta onde o anexo Excel será salvo.
            headline_prefix (str): O prefixo ou sufixo do assunto do e-mail a ser procurado.
    """
    theme_color = Fore.CYAN
    print_log(
        "INFO", "--- INICIANDO BUSCA POR E-MAIL QUANTUM ---", theme_color=theme_color
    )
    namespace = inicializar_outlook()
    if not namespace:
        logger_quantum.error(
            "Processo de extração de e-mail interrompido: namespace do Outlook não"
            " inicializado."
        )
        return

    try:
        asa_account = next(
            (f for f in namespace.Folders if "@asa.com.br" in f.Name.lower()), None
        )
        if not asa_account:
            print_log("ERROR", "Conta ASA não encontrada no Outlook.")
            logger_quantum.error(
                "Conta de e-mail da ASA não foi encontrada no Outlook."
            )
            return

        inbox = asa_account.Folders["Caixa de Entrada"]
        mensagens = inbox.Items
        mensagens.Sort("[ReceivedTime]", True)
        print_log(
            "INFO",
            "Caixa de entrada encontrada. Verificando os e-mails mais recentes...",
            theme_color=theme_color,
        )
        logger_quantum.info(
            f"Verificando {mensagens.Count} mensagens na caixa de entrada, ordenadas"
            " por data."
        )

    except Exception as e:
        print_log("ERROR", f"Erro ao acessar pastas do Outlook: {e}")
        logger_quantum.error(f"Erro fatal ao acessar pastas do Outlook: {e}", exc=e)
        return

    email_encontrado = False
    for i in range(min(50, mensagens.Count)):  # Aumentado o limite de busca para 50
        try:
            msg = mensagens.Item(i + 1)
            data_msg = msg.ReceivedTime.strftime("%Y-%m-%d")
            data_hoje = datetime.now().strftime("%Y-%m-%d")

            if data_msg != data_hoje:
                print_log(
                    "INFO",
                    "Busca encerrada. E-mails de hoje já foram todos verificados.",
                    theme_color=theme_color,
                )
                logger_quantum.info(
                    "Busca finalizada ao encontrar e-mail de data anterior."
                )
                return

            if msg.Subject and (
                msg.Subject.endswith(headline_prefix)
                or msg.Subject.startswith(headline_prefix)
            ):
                email_encontrado = True
                msg_processando = f"E-mail correspondente encontrado: '{msg.Subject}'"
                print_log("INFO", msg_processando, theme_color=theme_color)
                logger_quantum.info(msg_processando)

                for anexo in msg.Attachments:
                    if anexo.FileName.lower().endswith(".xlsx"):
                        nome_formatado = (
                            anexo.FileName.lower().replace(" ", "_").replace("-", "_")
                        )
                        Path(pasta_raiz_quantum).mkdir(parents=True, exist_ok=True)
                        caminho_anexo_salvo = os.path.join(
                            pasta_raiz_quantum, nome_formatado
                        )
                        anexo.SaveAsFile(caminho_anexo_salvo)

                        msg_anexo_salvo = (
                            f"Anexo '{nome_formatado}' salvo em: {caminho_anexo_salvo}"
                        )
                        print_log("INFO", msg_anexo_salvo, theme_color=theme_color)
                        logger_quantum.info(msg_anexo_salvo)
                        return

        except Exception as e:
            error_detail = f"Falha ao processar o e-mail: {msg.Subject if 'msg' in locals() else 'Desconhecido'}."
            print_log("ERROR", f"{error_detail} Detalhes: {e}")
            logger_quantum.error(error_detail, exc=e)

    if not email_encontrado:
        print_log(
            "AVISO",
            "Nenhum e-mail correspondente encontrado nos e-mails de hoje.",
            theme_color=Fore.YELLOW,
        )
        logger_quantum.info("Nenhum e-mail do Quantum encontrado para a data de hoje.")
    print_log("INFO", "--- BUSCA POR E-MAIL CONCLUÍDA ---", theme_color=Fore.CYAN)
