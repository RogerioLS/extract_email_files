
from src.extrair_excel_email import extrair_excel_email
from src.manipulacao_excel import processar_excel_extraido
from logger.logger_config import Logger, print_log

# --- CONFIGURAÇÃO CENTRAL ---
PASTA_RAIZ_QUANTUM = r"W:\Users\RL\projetos\quantum"
HEADLINE_PREFIX = "Daily Fundos"

if __name__ == "__main__":
    try:
        extrair_excel_email(PASTA_RAIZ_QUANTUM, HEADLINE_PREFIX)
        if processar_excel_extraido(PASTA_RAIZ_QUANTUM) > 10:
            print_log("INFO", "🔍 Mais de 10 valores NaN encontrados.")
        else:
            print_log("INFO", "✅ A planilha está dentro do limite de qualidade de dados.")
        print_log("INFO", "Processamento concluído com sucesso.")
    except Exception as e:
        print_log("ERROR", f"Ocorreu um erro: {e}")