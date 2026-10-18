"""Configuración centralizada del proyecto de automatización ABM Convenios."""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent


def _load_dotenv_if_present() -> None:
    """Carga variables desde .env sin requerir librerías externas si no están instaladas."""
    env_file = BASE_DIR / ".env"
    if not env_file.exists():
        return
    for raw_line in env_file.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip())


_load_dotenv_if_present()

# Casilla de correo a monitorear
TARGET_MAILBOX = os.getenv("TARGET_MAILBOX", "ABM.Convenios@swissmedical.com.ar")

# Credenciales Microsoft Graph API (Azure AD)
AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID", "")
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "")

# Credenciales IMAP (alternativa)
IMAP_SERVER = os.getenv("IMAP_SERVER", "outlook.office365.com")
IMAP_USER = os.getenv("IMAP_USER", TARGET_MAILBOX)
IMAP_PASSWORD = os.getenv("IMAP_PASSWORD", "")

# Configuración SharePoint Excel (Microsoft Graph Workbook API)
SHAREPOINT_SITE_ID = os.getenv("SHAREPOINT_SITE_ID", "")
SHAREPOINT_DRIVE_ID = os.getenv("SHAREPOINT_DRIVE_ID", "")
SHAREPOINT_ITEM_ID = os.getenv("SHAREPOINT_ITEM_ID", "")
SHAREPOINT_TABLE_NAME = os.getenv("SHAREPOINT_TABLE_NAME", "TablaExcepciones")

# Ruta local de salida (Excel sincronizado o de pruebas)
LOCAL_EXCEL_PATH = Path(os.getenv("LOCAL_EXCEL_PATH", str(BASE_DIR / "output" / "excepciones_abm.xlsx")))
INITIAL_INGRESO_NUMBER = int(os.getenv("INITIAL_INGRESO_NUMBER", "50607"))

# Reglas de negocio por defecto para Gerencia autorizante y Usuario de carga
DEFAULT_GERENCIA_AUTORIZANTE = os.getenv(
    "DEFAULT_GERENCIA_AUTORIZANTE",
    "Gerencia de Gestión de Prestadores Interior",
)
DEFAULT_USUARIO_CARGA = os.getenv("DEFAULT_USUARIO_CARGA", "")

# Mapeo opcional de solicitantes / remitentes conocidos -> (Gerencia Autorizante, Iniciales Usuario)
SOLICITANTES_MAP = {
    "maria elena velez": ("Gerencia de Gestión de Prestadores Interior", "MV"),
    "maria.velez@swissmedical.com.ar": ("Gerencia de Gestión de Prestadores Interior", "MV"),
}
