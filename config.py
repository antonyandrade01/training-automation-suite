import os
from dotenv import load_dotenv

# Carrega as variáveis do arquivo .env para o ambiente
load_dotenv()

# --- Banco de Dados ---
DB_HOST = os.getenv("DB_HOST")
DB_USER = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")
DB_NAME = os.getenv("DB_NAME")

# --- Caminhos ---
TRAINING_ASSETS_BASE_PATH = os.getenv("TRAINING_ASSETS_BASE_PATH")

# --- API Movidesk ---
MOVIDESK_API_TOKEN = os.getenv("MOVIDESK_API_TOKEN")
MOVIDESK_VERSION_FIELD_ID = int(os.getenv("MOVIDESK_VERSION_FIELD_ID", 0))
MOVIDESK_OTHER_FIELD_ID = int(os.getenv("MOVIDESK_OTHER_FIELD_ID", 0))
MOVIDESK_OTHER_FIELD_RULE_ID = int(os.getenv("MOVIDESK_OTHER_FIELD_RULE_ID", 0))

# --- IDs Movidesk ---
MOVIDESK_OWNER_ID = os.getenv("MOVIDESK_OWNER_ID")
MOVIDESK_OWNER_TEAM_NAME = os.getenv("MOVIDESK_OWNER_TEAM_NAME")
MOVIDESK_ACTION_CREATOR_ID = os.getenv("MOVIDESK_ACTION_CREATOR_ID")

# --- Assinatura de Email ---
ACTION_HTML_SIGNATURE = os.getenv("ACTION_HTML_SIGNATURE", "<p>Atenciosamente,</p>")

def validar_configuracoes():
    """Verifica se as configurações essenciais foram carregadas do .env"""
    essenciais = {
        "DB_HOST": DB_HOST, "DB_USER": DB_USER, "DB_PASSWORD": DB_PASSWORD,
        "DB_NAME": DB_NAME, "MOVIDESK_API_TOKEN": MOVIDESK_API_TOKEN,
        "TRAINING_ASSETS_BASE_PATH": TRAINING_ASSETS_BASE_PATH
    }
    ausentes = [chave for chave, valor in essenciais.items() if not valor]
    if ausentes:
        print(f"\033[91mERRO: As seguintes variáveis de ambiente essenciais não foram definidas: {', '.join(ausentes)}\033[0m")
        print("\033[93mPor favor, copie o arquivo '.env.example' para '.env' e preencha com suas credenciais.\033[0m")
        exit(1)