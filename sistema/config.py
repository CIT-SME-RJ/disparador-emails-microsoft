from pathlib import Path

# =========================
# CAMINHOS DO PROJETO
# =========================

PASTA_SISTEMA = Path(__file__).resolve().parent
PASTA_PROJETO = PASTA_SISTEMA.parent

BASE_DIR = PASTA_PROJETO / "Arquivo"

PASTA_PLANILHAS = BASE_DIR / "Planilhas"
PASTA_ANEXOS_DINAMICOS = BASE_DIR / "Anexos_Personalizados"
PASTA_ANEXOS_FIXOS = BASE_DIR / "Anexos_Fixos"
PASTA_ASSINATURA = BASE_DIR / "Assinatura"
PASTA_LOGS = BASE_DIR / "Logs"

PASTAS_SISTEMA = [
    PASTA_PLANILHAS,
    PASTA_ANEXOS_DINAMICOS,
    PASTA_ANEXOS_FIXOS,
    PASTA_ASSINATURA,
    PASTA_LOGS
]


def garantir_pastas():
    for pasta in PASTAS_SISTEMA:
        pasta.mkdir(parents=True, exist_ok=True)