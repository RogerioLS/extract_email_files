import os
from pathlib import Path

import pandas as pd
from colorama import Fore

# Importações locais
from source.logger.logger_config import logger_quantum, print_log


def ler_arquivo_excel(caminho_excel: Path):
    """
    Lê um arquivo Excel e o carrega em um DataFrame do pandas.

    Args:
            caminho_excel (Path): O caminho completo para o arquivo Excel.

    Returns:
            pd.DataFrame: Um DataFrame com os dados do arquivo, ou None se ocorrer um erro.
    """
    if not caminho_excel.is_file():
        error_msg = (
            f"O arquivo não foi encontrado no caminho especificado: {caminho_excel}"
        )
        print_log("ERROR", error_msg)
        logger_quantum.error(error_msg)
        return None

    try:
        print_log("INFO", f"Lendo o arquivo Excel: {caminho_excel.name}...")
        df = pd.read_excel(caminho_excel)
        logger_quantum.info(f"Arquivo '{caminho_excel.name}' lido com sucesso.")
        return df
    except Exception as e:
        error_msg = (
            "Ocorreu um erro inesperado ao tentar ler o arquivo"
            f" '{caminho_excel.name}': {e}"
        )
        print_log("ERROR", error_msg)
        logger_quantum.error(error_msg, exc=e)
        return None


def quantidade_nan(dataframe: pd.DataFrame):
    """
    Conta a quantidade de valores NaN especificamente na coluna 'Retorno' de um DataFrame.

    Args:
            dataframe (pd.DataFrame): O DataFrame a ser analisado.

    Returns:
            int: A quantidade total de valores NaN encontrados na coluna 'Retorno'.
    """
    if "Retorno" not in dataframe.columns:
        error_msg = (
            "A coluna 'Retorno' não foi encontrada no DataFrame para a contagem de"
            " NaNs."
        )
        print_log("ERROR", error_msg)
        logger_quantum.error(error_msg)
        return 0

    contagem_nan = dataframe["Retorno"].isna().sum()
    logger_quantum.info(
        f"Contagem de NaNs na coluna 'Retorno' finalizada: {contagem_nan}"
        " encontrados."
    )
    return contagem_nan


def ler_excel_mais_recente_da_pasta(caminho_pasta: Path):
    """
    Encontra e lê o arquivo Excel (.xlsx ou .xls) mais recente em uma pasta.

    Args:
            caminho_pasta (Path): O caminho para a pasta que contém os arquivos.

    Returns:
            pd.DataFrame: Um DataFrame com os dados do arquivo mais recente, ou None se não for encontrado.
    """
    caminho_pasta = Path(caminho_pasta)
    if not caminho_pasta.is_dir():
        error_msg = f"A pasta não foi encontrada em: {caminho_pasta}"
        print_log("ERROR", error_msg)
        logger_quantum.error(error_msg)
        return None

    print_log("INFO", f"Procurando por arquivos Excel em: {caminho_pasta}")
    arquivos_excel = list(caminho_pasta.glob("*.xlsx")) + list(
        caminho_pasta.glob("*.xls")
    )

    if not arquivos_excel:
        msg = f"Nenhum arquivo Excel (.xlsx ou .xls) encontrado na pasta."
        print_log("AVISO", msg, theme_color=Fore.YELLOW)
        logger_quantum.info(msg)
        return None

    arquivo_mais_recente = max(arquivos_excel, key=os.path.getmtime)
    print_log("INFO", f"Arquivo mais recente encontrado: {arquivo_mais_recente.name}")
    logger_quantum.info(
        f"Arquivo mais recente para processamento: {arquivo_mais_recente.name}"
    )
    return ler_arquivo_excel(arquivo_mais_recente)


def processar_excel_extraido(caminho_pasta: Path, limites_null: int):
    """
    Orquestra a leitura do Excel mais recente e a verificação de qualidade (contagem de NaNs).

    Args:
            caminho_pasta (Path): A pasta onde o arquivo Excel de entrada está localizado.
            limites_null (int): O número máximo de valores nulos permitidos.

    Returns:
            pd.DataFrame: O DataFrame do Excel se a contagem de nulos for aceitável.
            int: A contagem de nulos se o limite for excedido.
            None: Se nenhum arquivo for encontrado ou ocorrer um erro.
    """
    df_excel = ler_excel_mais_recente_da_pasta(caminho_pasta)
    if df_excel is None:
        return None  # Erro já logado pela função anterior

    contagem_nan = quantidade_nan(df_excel)
    print_log(
        "INFO",
        f"Verificação de qualidade: {contagem_nan} nulos encontrados (Limite: {limites_null}).",
    )

    if contagem_nan > limites_null:
        return contagem_nan
    return df_excel
