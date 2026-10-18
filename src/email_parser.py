"""Extractor de datos del cuerpo y metadatos del correo electrónico."""

from datetime import datetime
import email
from email.message import Message
from email.utils import parsedate_to_datetime
from html.parser import HTMLParser
from pathlib import Path
import re
from typing import List, Optional, Union

from src.models import EmailExtractedData
from src.pdf_parser import parse_money_amount


MESES_ES = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}


class _HTMLTableAndTextExtractor(HTMLParser):
    """Convierte HTML en texto estructurado preservando celdas de tablas."""

    def __init__(self) -> None:
        super().__init__()
        self.cells: List[str] = []
        self._current: List[str] = []
        self._in_cell = False

    def handle_starttag(self, tag: str, attrs: list) -> None:
        if tag in ("td", "th", "p", "div", "tr", "br", "li"):
            self._flush()
            if tag in ("td", "th"):
                self._in_cell = True

    def handle_endtag(self, tag: str) -> None:
        if tag in ("td", "th", "p", "div", "tr"):
            self._flush()
            self._in_cell = False

    def handle_data(self, data: str) -> None:
        self._current.append(data)

    def _flush(self) -> None:
        text = re.sub(r"\s+", " ", "".join(self._current)).strip()
        if text:
            self.cells.append(text)
        self._current = []

    def get_tokens(self) -> List[str]:
        self._flush()
        return self.cells


def html_to_tokens(content: str) -> List[str]:
    """Extrae bloques de texto/celdas desde un cuerpo HTML o texto plano."""
    if "<" in content and ">" in content:
        parser = _HTMLTableAndTextExtractor()
        parser.feed(content)
        tokens = parser.get_tokens()
        if tokens:
            return tokens
    return [line.strip() for line in content.splitlines() if line.strip()]


def normalize_date_ddmmyyyy(raw_date: str) -> str:
    """Normaliza una fecha ISO, RFC2822 o en español ('octubre 16, 2026') a DD/MM/YYYY."""
    if not raw_date:
        return ""
    raw_date = raw_date.strip()

    # Ya está en formato DD/MM/YYYY
    m_direct = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", raw_date)
    if m_direct:
        d, m, y = int(m_direct.group(1)), int(m_direct.group(2)), int(m_direct.group(3))
        return f"{d:02d}/{m:02d}/{y:04d}"

    # Formato español: "viernes, octubre 16, 2026" o "16 de octubre de 2026"
    m_es1 = re.search(r"([a-záéíóú]+)\s+(\d{1,2}),?\s+(\d{4})", raw_date.lower())
    if m_es1 and m_es1.group(1) in MESES_ES:
        mes = MESES_ES[m_es1.group(1)]
        dia = int(m_es1.group(2))
        anio = int(m_es1.group(3))
        return f"{dia:02d}/{mes:02d}/{anio:04d}"

    m_es2 = re.search(r"(\d{1,2})\s+de\s+([a-záéíóú]+)\s+de\s+(\d{4})", raw_date.lower())
    if m_es2 and m_es2.group(2) in MESES_ES:
        dia = int(m_es2.group(1))
        mes = MESES_ES[m_es2.group(2)]
        anio = int(m_es2.group(3))
        return f"{dia:02d}/{mes:02d}/{anio:04d}"

    # Formato ISO 8601 (Microsoft Graph: 2026-10-16T17:25:37Z)
    try:
        iso_clean = raw_date.replace("Z", "+00:00")
        dt = datetime.fromisoformat(iso_clean)
        return dt.strftime("%d/%m/%Y")
    except Exception:
        pass

    # Formato RFC 2822 (Cabecera Date de correo)
    try:
        dt = parsedate_to_datetime(raw_date)
        return dt.strftime("%d/%m/%Y")
    except Exception:
        pass

    return raw_date


def parse_email_body(
    body_content: str,
    subject: str = "",
    sender_name: str = "",
    sender_email: str = "",
    received_date: str = "",
) -> EmailExtractedData:
    """
    Extrae del cuerpo del correo:
    - Fecha de recepción del mail
    - Número de convenio ('Convenio (1)')
    - Valor de referencia ('Valor de referencia (2)')
    - N° Trámite Excepción y datos de gerencia/remitente si están en el cuerpo.
    """
    tokens = html_to_tokens(body_content)
    full_text = "\n".join(tokens)

    result = EmailExtractedData(
        subject=subject,
        sender_name=sender_name,
        sender_email=sender_email,
        fecha_recepcion=normalize_date_ddmmyyyy(received_date),
    )

    # Si el correo fue reenviado o no trae cabecera de fecha, buscar "Enviado:" en el cuerpo
    if not result.fecha_recepcion:
        m_enviado = re.search(r"Enviado:\s*([^\n\r]+)", full_text, re.I)
        if m_enviado:
            result.fecha_recepcion = normalize_date_ddmmyyyy(m_enviado.group(1))

    # 1. Extraer Número de convenio ("Convenio (1)")
    for idx, tok in enumerate(tokens):
        if re.search(r"Convenio\s*\(\s*1\s*\)", tok, re.I):
            # Puede estar en el mismo token ("Convenio (1) 12421") o en la siguiente celda de la tabla
            m_same = re.search(r"Convenio\s*\(\s*1\s*\)\s*[:\-]?\s*(\d+)", tok, re.I)
            if m_same:
                result.numero_convenio = int(m_same.group(1))
                break
            if idx + 1 < len(tokens):
                next_tok = tokens[idx + 1].strip()
                if re.match(r"^\d+$", next_tok):
                    result.numero_convenio = int(next_tok)
                    break

    # 2. Extraer Valor de referencia ("Valor de referencia (2)")
    for idx, tok in enumerate(tokens):
        if re.search(r"Valor\s+de\s+referencia\s*\(\s*2\s*\)", tok, re.I):
            # Buscar en el mismo token o en los siguientes 2 tokens (por si el signo '$' viene en celda separada)
            remainder = re.sub(r".*Valor\s+de\s+referencia\s*\(\s*2\s*\)", "", tok, flags=re.I).strip()
            candidates = [remainder] + tokens[idx + 1 : idx + 3]
            for cand in candidates:
                if not cand:
                    continue
                # Cortar si el texto se pegó con "Prestador N°"
                cand_cut = re.split(r"Prestador\s*N", cand, flags=re.I)[0]
                amount = parse_money_amount(cand_cut)
                if amount is not None:
                    result.valor_referencia = amount
                    break
            if result.valor_referencia is not None:
                break

    # 3. Extraer Trámite Excepción del cuerpo (respaldo)
    m_tramite = re.search(r"Tr[aá]mite\s+Excepci[oó]n\s*#\s*(\d+)", full_text, re.I)
    if m_tramite:
        result.tramite_excepcion_body = m_tramite.group(1)

    # 4. Detectar Gerencia en la firma si aparece
    if re.search(r"Prestadores\s+Interior", full_text, re.I):
        result.gerencia_firma = "Gerencia de Gestión de Prestadores Interior"
    elif re.search(r"Auditor[ií]a\s+M[eé]dica", full_text, re.I):
        result.gerencia_firma = "Auditoria Médica"

    return result


def parse_eml_file(eml_path: Union[str, Path]) -> EmailExtractedData:
    """Lee un archivo estándar .eml extrayendo el cuerpo y todos los adjuntos .pdf."""
    raw_bytes = Path(eml_path).read_bytes()
    msg: Message = email.message_from_bytes(raw_bytes)

    subject = msg.get("Subject", "")
    from_header = msg.get("From", "")
    date_header = msg.get("Date", "")

    sender_name, sender_email = email.utils.parseaddr(from_header)

    html_body = ""
    text_body = ""
    pdf_attachments = []

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            filename = part.get_filename() or ""
            if filename.lower().endswith(".pdf") or content_type == "application/pdf":
                payload = part.get_payload(decode=True)
                if payload:
                    pdf_attachments.append({"filename": filename or "adjunto.pdf", "content": payload})
            elif content_type == "text/html" and not html_body:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    html_body = payload.decode(charset, errors="replace")
            elif content_type == "text/plain" and not text_body:
                payload = part.get_payload(decode=True)
                if payload:
                    charset = part.get_content_charset() or "utf-8"
                    text_body = payload.decode(charset, errors="replace")
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            charset = msg.get_content_charset() or "utf-8"
            text_body = payload.decode(charset, errors="replace")

    extracted = parse_email_body(
        body_content=html_body or text_body,
        subject=subject,
        sender_name=sender_name,
        sender_email=sender_email,
        received_date=date_header,
    )
    extracted.pdf_attachments = pdf_attachments
    return extracted
