import os
import time
import warnings
from pathlib import Path

import pandas as pd
from colorama import Fore
from dotenv import load_dotenv

from source.email.envia_email_alerta import enviar_email_alerta
from source.email.envia_email_sucesso import enviar_email_sucesso
from source.email.extrair_excel_email import extrair_excel_email
from source.logger.logger_config import logger_quantum, print_log
from source.manipulacao_excel.manipulacao_excel import processar_excel_extraido

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# --- CONFIGURAÇÃO INICIAL ---
# Carrega as variáveis de ambiente do arquivo .env
load_dotenv()

# Define uma cor tema para os logs desta execução
THEME_COLOR = Fore.MAGENTA

# Carrega as variáveis de ambiente necessárias
PASTA_RAIZ_QUANTUM = os.getenv("PASTA_RAIZ_QUANTUM")
HEADLINE_PREFIX = os.getenv("HEADLINE_PREFIX")

# Configurações da lógica de retentativa
MAX_TENTATIVAS = 3
INTERVALO_TENTATIVAS_SEGUNDOS = 60


def main():
    # --- ETAPA 1: Extrair anexo do e-mail com lógica de retentativa ---
    print_log(
        "AÇÃO", "Iniciando extração de anexo do e-mail...", theme_color=THEME_COLOR
    )
    arquivo_salvo = False
    for tentativa in range(1, MAX_TENTATIVAS + 1):
        print_log(
            "AÇÃO",
            f"Tentativa {tentativa}/{MAX_TENTATIVAS}: Buscando e-mail e extraindo anexo...",
            theme_color=THEME_COLOR,
        )
        nome_arquivo_esperado = extrair_excel_email(PASTA_RAIZ_QUANTUM, HEADLINE_PREFIX)

        if nome_arquivo_esperado:
            caminho_completo_arquivo = Path(PASTA_RAIZ_QUANTUM) / nome_arquivo_esperado
            if caminho_completo_arquivo.is_file():
                print_log(
                    "INFO",
                    f"Sucesso! Arquivo '{nome_arquivo_esperado}' encontrado na pasta.",
                    theme_color=Fore.GREEN,
                )
                logger_quantum.info(
                    f"Arquivo '{nome_arquivo_esperado}' validado com sucesso na tentativa {tentativa}."
                )
                arquivo_salvo = True
                break  # Sai do loop de tentativas

        print_log(
            "AVISO",
            f"Arquivo não encontrado na tentativa {tentativa}. Aguardando {INTERVALO_TENTATIVAS_SEGUNDOS} segundos...",
            theme_color=Fore.YELLOW,
        )
        if tentativa < MAX_TENTATIVAS:
            time.sleep(INTERVALO_TENTATIVAS_SEGUNDOS)

    if not arquivo_salvo:
        msg_falha = "Arquivo não foi baixado após todas as tentativas."
        print_log("ERROR", msg_falha)
        logger_quantum.error(msg_falha)
        return print_log(
            "INFO",
            "❌ --- PROCESSO QUANTUM INTERROMPIDO --- ❌",
            theme_color=THEME_COLOR,
        )

    # --- ETAPA 2: Processar a planilha e verificar a qualidade dos dados ---
    limites_null = 30
    print_log(
        "AÇÃO",
        f"Processando planilha e verificando se há mais de {limites_null} valores nulos...",
        theme_color=THEME_COLOR,
    )
    resultado_processamento = processar_excel_extraido(
        Path(PASTA_RAIZ_QUANTUM), limites_null
    )
    logger_quantum.info("Processamento da planilha concluído.")

    # --- ETAPA 3: Tomar decisão com base na qualidade dos dados ---
    if not isinstance(resultado_processamento, pd.DataFrame):
        # CASO DE FALHA: Muitos valores nulos
        contagem_nan = resultado_processamento
        print_log(
            "AVISO",
            f"Limite de valores nulos excedido! Encontrados: {contagem_nan}. Limite: {limites_null}.",
        )
        logger_quantum.error(
            f"Validação falhou: {contagem_nan} valores nulos encontrados (limite: {limites_null})."
        )

        print_log("AÇÃO", "Enviando e-mail de alerta...", theme_color=THEME_COLOR)
        enviar_email_alerta(contagem_nan, limites_null)
        logger_quantum.info("E-mail de alerta enviado.")
        return print_log(
            "INFO",
            "❌ --- PROCESSO QUANTUM INTERROMPIDO DEVIDO A ERRO --- ❌",
            theme_color=THEME_COLOR,
        )
    else:
        # CASO DE SUCESSO: Dados válidos
        print_log(
            "AÇÃO",
            "Enviando e-mail de confirmação de sucesso...",
            theme_color=THEME_COLOR,
        )
        enviar_email_sucesso()
        logger_quantum.info("E-mail de sucesso enviado.")
        return print_log(
            "INFO",
            "✅ --- PROCESSO QUANTUM CONCLUÍDO COM SUCESSO --- ✅",
            theme_color=THEME_COLOR,
        )


if __name__ == "__main__":
    print_log(
        "INFO", "🚀 --- INICIANDO PROCESSO QUANTUM --- 🚀", theme_color=THEME_COLOR
    )
    logger_quantum.info("Processo Quantum iniciado.")

    try:
        main()
    except Exception as e:
        # Tratamento de erro para qualquer falha inesperada no processo
        print_log("ERROR", f"Ocorreu um erro crítico que interrompeu o processo: {e}")
        logger_quantum.error(f"Erro fatal na execução do main.py: {e}", exc=e)
        print_log(
            "INFO",
            "❌ --- PROCESSO QUANTUM INTERROMPIDO DEVIDO A ERRO --- ❌",
            theme_color=THEME_COLOR,
        )
        raise
