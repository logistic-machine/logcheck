from __future__ import annotations

import os
import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Resolução de diretórios (compatível com PyInstaller)
# ---------------------------------------------------------------------------


def _base_dir() -> Path:
    """Diretório base onde ficam os arquivos empacotados (planilhas de dados)."""
    if getattr(sys, "frozen", False):  # rodando dentro de um .exe do PyInstaller
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    # .../inventario/src/config.py  ->  .../inventario
    return Path(__file__).resolve().parent.parent


def _output_dir() -> Path:
    
    override = os.environ.get("INVENTARIO_OUTPUT_ROOT", "").strip() #AINDA NÃO ESTÁ FUNCIONANDO, PRECISO TESTE#
    if override:
        return Path(override).expanduser().resolve()
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


BASE_DIR = _base_dir()
OUTPUT_DIR = _output_dir()

DATA_DIR = BASE_DIR / "data"
CRONOGRAMA_PATH = DATA_DIR / "Cronograma_Atualizado.xlsx"
CHECKLIST_PATH = DATA_DIR / "Checklist_Avaliacao.xlsx"

RESULTADOS_DIR = OUTPUT_DIR / "resultados"
EVIDENCIAS_DIR = OUTPUT_DIR / "evidencias"
RESULTADOS_XLSX = RESULTADOS_DIR / "Inventario_Resultados.xlsx"

# ---------------------------------------------------------------------------
# Constantes de domínio
# ---------------------------------------------------------------------------

# Planilha e colunas usadas na base "Cronograma Atualizado" (Ação 1.2).
CRONOGRAMA_SHEET = "Planilha1"
CRONOGRAMA_HEADER_ROW = 1
# Mapeamento (letra da coluna -> campo lógico) conforme solicitado: B, C, D, F, I.
CRONOGRAMA_COLS = {
    "area": "B",
    "centro": "C",
    "empresa": "D",
    "localidade": "F",
    "tipo_centro": "I",
}

# Rótulos exibidos na interface, na ordem das listas suspensas em cascata.
COMBO_FIELDS = [
    ("area", "Área"),
    ("centro", "Centro"),
    ("empresa", "Empresa"),
    ("localidade", "Localidade"),
    ("tipo_centro", "Tipo de Centro"),
]

# Planilha do questionário (checklist).
CHECKLIST_SHEET = "Questionário"

# Opções de avaliação de cada pergunta (Ação 2.3).
RESP_SIM = "SIM"
RESP_NAO = "NÃO"
RESP_NA = "N/A"
RESPOSTAS = [RESP_SIM, RESP_NAO, RESP_NA]

# ---------------------------------------------------------------------------
# Autenticação / perfis (versão web)
# ---------------------------------------------------------------------------

# Arquivo (JSON) onde os usuários cadastrados são armazenados.
USUARIOS_JSON = RESULTADOS_DIR / "usuarios.json"

# Senha padrão atribuída a novos auditores cadastrados pelo Master.
# Pode ser sobrescrita pela variável de ambiente SENHA_PADRAO (ex.: no deploy).
SENHA_PADRAO = os.environ.get("SENHA_PADRAO", "SENHA_A_DEFINIR")

# Configuração SMTP para envio de relatório por e-mail
SMTP_SERVER = os.environ.get("SMTP_SERVER", "")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD", "")
SMTP_SENDER_EMAIL = os.environ.get("SMTP_SENDER_EMAIL", SMTP_USER)
SMTP_USE_TLS = os.environ.get("SMTP_USE_TLS", "true").strip().lower() in ("1", "true", "yes")
SMTP_USE_SSL = os.environ.get("SMTP_USE_SSL", "false").strip().lower() in ("1", "true", "yes")
# Use Outlook (COM) instead of SMTP when true. Set to 'true' in env to enable.
USE_OUTLOOK = os.environ.get("USE_OUTLOOK", "false").strip().lower() in ("1", "true", "yes")

# E-mail(s) Master inicial(is) — já entram liberados e com poderes de Master.
# Pode ser sobrescrito por MASTER_EMAILS (lista separada por vírgula).
_master_env = os.environ.get("MASTER_EMAILS", "").strip()
MASTER_EMAILS_INICIAIS = (
    [e.strip().lower() for e in _master_env.split(",") if e.strip()]
    if _master_env
    else ["assistente.logistica4@amaranzero.com"]
)


def ensure_output_dirs() -> None:
    """Garante que as pastas de saída existam."""
    RESULTADOS_DIR.mkdir(parents=True, exist_ok=True)
    EVIDENCIAS_DIR.mkdir(parents=True, exist_ok=True)
