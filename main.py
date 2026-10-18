"""
Punto de entrada principal de la automatización ABM Convenios - Swiss Medical.

Modos de uso:
  1. Probar con el ejemplo real incluido (Email + PDF de Ledesma German):
     python3 main.py --demo

  2. Procesar un archivo .eml local o un par (--pdf y --html):
     python3 main.py --eml samples/email_ejemplo.eml

  3. Ejecutar lectura real de la bandeja ABM.Convenios@swissmedical.com.ar:
     python3 main.py --run-mailbox
"""

import argparse
from email.message import EmailMessage
from pathlib import Path
import sys
import time

import config
from src.email_client import GraphMailboxClient, ImapMailboxClient, get_graph_access_token
from src.email_parser import parse_email_body, parse_eml_file
from src.excel_writer import LocalExcelWriter, SharePointExcelWriter
from src.models import EXCEL_COLUMNS
from src.pdf_parser import parse_formulario_pdf
from src.processor import build_excel_record


SAMPLE_EMAIL_HTML = """<!DOCTYPE html>
<html>
<body>
<p>De: Maria Elena Velez &lt;Maria.Velez@swissmedical.com.ar&gt;<br>
Enviado: viernes, octubre 16, 2026 5:25:37 p.m.<br>
Para: ABM de Convenios &lt;ABM.Convenios@swissmedical.com.ar&gt;<br>
Cc: Elizabeth Martinez &lt;Elizabeth.Martinez@swissmedical.com.ar&gt;<br>
Asunto: Centro-Cba-E-01/10/26 al 31/10/26-P.1212 S Allende-Soc. 214421/01 LEDESMA, GERMAN ARIEL.</p>

<p>Sres.<br>ABM de Convenios<br>Presente</p>
<p>Se envía excepción para su carga Soc. 214421/01 LEDESMA, GERMAN ARIEL.</p>

<table border="1" cellspacing="0" cellpadding="4">
  <tr>
    <td>Convenio (1)</td>
    <td colspan="3">12421</td>
  </tr>
  <tr>
    <td>Valor de referencia (2)</td>
    <td>$ 457,650.00</td>
    <td>Prestador N°</td>
    <td>25966 Hospital Privado Rosario</td>
  </tr>
  <tr>
    <td></td>
    <td colspan="3">Trámite Excepción #302407236</td>
  </tr>
  <tr>
    <td>Observaciones (3)</td>
    <td colspan="3">
      Cob. por Excepción Soc. 214421/01 LEDESMA, GERMAN ARIEL.<br>
      Procedimiento: Cód. 1987 FILMARRAY PANEL MENINGITIS (BIOLOGÍA MOLECULAR) .<br><br>
      Trámite #302390046 Gestion Presupuestos de Prestadores<br>
      Trámite #302298742 Gestión Autorizaciones.
    </td>
  </tr>
</table>

<p>Ma. Elena Velez<br>
Ejecutiva de Cuentas<br>
Gcia. Prestadores Interior<br>
Tel: 3512088751</p>
</body>
</html>
"""


def ensure_sample_eml() -> Path:
    """Genera samples/email_ejemplo.eml con el cuerpo HTML y el PDF adjunto real si no existe."""
    eml_path = config.BASE_DIR / "samples" / "email_ejemplo.eml"
    pdf_path = config.BASE_DIR / "samples" / "Excep. LEDESMA, GERMAN.pdf"
    if eml_path.exists() or not pdf_path.exists():
        return eml_path

    msg = EmailMessage()
    msg["Subject"] = (
        "Centro-Cba-E-01/10/26 al 31/10/26-P.1212 S Allende-Soc. 214421/01 LEDESMA, GERMAN ARIEL."
    )
    msg["From"] = "Maria Elena Velez <Maria.Velez@swissmedical.com.ar>"
    msg["To"] = "ABM de Convenios <ABM.Convenios@swissmedical.com.ar>"
    msg["Cc"] = "Elizabeth Martinez <Elizabeth.Martinez@swissmedical.com.ar>"
    msg["Date"] = "Fri, 16 Oct 2026 17:25:37 -0300"

    msg.set_content("Se envía excepción para su carga Soc. 214421/01 LEDESMA, GERMAN ARIEL.")
    msg.add_alternative(SAMPLE_EMAIL_HTML, subtype="html")
    msg.add_attachment(
        pdf_path.read_bytes(),
        maintype="application",
        subtype="pdf",
        filename="Excep. LEDESMA, GERMAN.pdf",
    )
    eml_path.write_bytes(msg.as_bytes())
    return eml_path


def print_record_summary(record) -> None:
    """Imprime por consola una tabla legible con las 23 columnas generadas para el Excel."""
    print("\n" + "=" * 78)
    print(" REGISTRO PROCESADO PARA EXCEL EN SHAREPOINT (23 COLUMNAS)")
    print("=" * 78)
    for col_name, value in record.to_dict().items():
        print(f"  {col_name:<30} : {value}")
    print("=" * 78 + "\n")


def run_demo() -> int:
    """Ejecuta el flujo completo con el correo y el PDF de ejemplo y guarda el Excel local."""
    eml_path = ensure_sample_eml()
    if not eml_path.exists():
        print("Error: No se encontró el archivo de ejemplo en samples/", file=sys.stderr)
        return 1

    print(f"[1/4] Leyendo correo entrante de ejemplo: {eml_path.name}")
    email_data = parse_eml_file(eml_path)

    if not email_data.pdf_attachments:
        print("Error: El correo no contiene adjuntos PDF.", file=sys.stderr)
        return 1

    att = email_data.pdf_attachments[0]
    print(f"[2/4] Extrayendo datos del Formulario PDF adjunto: {att['filename']}")
    pdf_data = parse_formulario_pdf(att["content"])

    writer = LocalExcelWriter(config.LOCAL_EXCEL_PATH)
    next_ingreso = writer.get_next_ingreso_number()

    print(f"[3/4] Combinando Cuerpo del Mail + Formulario PDF (N° Ingreso #{next_ingreso})...")
    record = build_excel_record(
        email_data=email_data,
        pdf_data=pdf_data,
        numero_ingreso=next_ingreso,
    )

    print_record_summary(record)

    output_path = writer.append_record(record)
    print(f"[4/4] Fila agregada exitosamente en el archivo Excel: {output_path}")
    return 0


def process_single_eml(eml_file: Path) -> int:
    """Procesa un archivo .eml específico y agrega sus formularios PDF al Excel."""
    email_data = parse_eml_file(eml_file)
    writer = LocalExcelWriter(config.LOCAL_EXCEL_PATH)

    for att in email_data.pdf_attachments:
        pdf_data = parse_formulario_pdf(att["content"])
        next_ingreso = writer.get_next_ingreso_number()
        record = build_excel_record(email_data, pdf_data, numero_ingreso=next_ingreso)
        print_record_summary(record)
        output_path = writer.append_record(record)
        print(f"Guardado en: {output_path}")
    return 0


def run_mailbox_once() -> int:
    """
    Se conecta a ABM.Convenios@swissmedical.com.ar (vía Microsoft Graph API o IMAP),
    procesa todos los correos nuevos con adjunto PDF y escribe en SharePoint y/o Excel local.
    """
    use_graph = bool(
        config.AZURE_TENANT_ID
        and config.AZURE_CLIENT_ID
        and config.AZURE_CLIENT_SECRET
        and config.AZURE_TENANT_ID != "tu-tenant-id"
    )

    if use_graph:
        print(f"Conectando por Microsoft Graph API a {config.TARGET_MAILBOX}...")
        token = get_graph_access_token()
        mail_client = GraphMailboxClient(access_token=token, mailbox=config.TARGET_MAILBOX)
        emails = mail_client.fetch_unread_exception_emails()

        use_sharepoint = bool(
            config.SHAREPOINT_SITE_ID and config.SHAREPOINT_DRIVE_ID and config.SHAREPOINT_ITEM_ID
        )
        sp_writer = SharePointExcelWriter(access_token=token) if use_sharepoint else None
        local_writer = LocalExcelWriter(config.LOCAL_EXCEL_PATH)

        print(f"Correos pendientes con adjunto PDF encontrados: {len(emails)}")
        for email_data in emails:
            for att in email_data.pdf_attachments:
                pdf_data = parse_formulario_pdf(att["content"])
                if sp_writer:
                    next_num = sp_writer.get_next_ingreso_number()
                    record = build_excel_record(email_data, pdf_data, numero_ingreso=next_num)
                    sp_writer.append_record(record)
                    print(f"[SharePoint] Cargado N° Ingreso {next_num} - Trámite {record.n_de_llamado}")
                else:
                    next_num = local_writer.get_next_ingreso_number()
                    record = build_excel_record(email_data, pdf_data, numero_ingreso=next_num)
                    local_writer.append_record(record)
                    print(f"[Excel Local] Cargado N° Ingreso {next_num} - Trámite {record.n_de_llamado}")

            mail_client.mark_as_read(email_data.message_id)
        return 0

    if config.IMAP_PASSWORD:
        print(f"Conectando por IMAP a {config.IMAP_USER}...")
        imap_client = ImapMailboxClient()
        emails = imap_client.fetch_unread_exception_emails(mark_seen=True)
        local_writer = LocalExcelWriter(config.LOCAL_EXCEL_PATH)
        print(f"Correos pendientes con adjunto PDF encontrados: {len(emails)}")
        for email_data in emails:
            for att in email_data.pdf_attachments:
                pdf_data = parse_formulario_pdf(att["content"])
                next_num = local_writer.get_next_ingreso_number()
                record = build_excel_record(email_data, pdf_data, numero_ingreso=next_num)
                local_writer.append_record(record)
                print(f"[Excel Local] Cargado N° Ingreso {next_num} - Trámite {record.n_de_llamado}")
        return 0

    print(
        "No se configuraron credenciales en .env (AZURE_* o IMAP_*). "
        "Ejecutando modo demostración (--demo) con los archivos reales de prueba:\n"
    )
    return run_demo()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Automatización de carga de Excepciones ABM Convenios (Swiss Medical)"
    )
    parser.add_argument(
        "--demo",
        action="store_true",
        help="Ejecuta una prueba completa con el correo y el PDF de ejemplo (Ledesma, German Ariel)",
    )
    parser.add_argument(
        "--eml",
        type=Path,
        help="Procesa un archivo de correo .eml específico",
    )
    parser.add_argument(
        "--run-mailbox",
        action="store_true",
        help="Revisa la bandeja de entrada ABM.Convenios@swissmedical.com.ar y procesa correos no leídos",
    )
    parser.add_argument(
        "--watch",
        type=int,
        default=0,
        metavar="SEGUNDOS",
        help="Monitorea continuamente la casilla cada N segundos (ej. --watch 60)",
    )

    args = parser.parse_args()

    if args.demo:
        return run_demo()
    if args.eml:
        return process_single_eml(args.eml)
    if args.watch > 0:
        print(f"Iniciando monitoreo continuo cada {args.watch} segundos...")
        while True:
            run_mailbox_once()
            time.sleep(args.watch)
    if args.run_mailbox:
        return run_mailbox_once()

    return run_demo()


if __name__ == "__main__":
    raise SystemExit(main())
