"""Extractor de datos del PDF 'FORMULARIO DE EXCEPCIONES' de Swiss Medical."""

import re
import zlib
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union

from src.models import PdfFormularioData


def _unescape_pdf_string(text: str) -> str:
    """Limpia secuencias escapadas típicas de cadenas PDF como \\( o \\)."""
    text = text.replace(r"\(", "(").replace(r"\)", ")").replace(r"\\", "\\")
    text = text.replace("&aacute;", "á").replace("&eacute;", "é").replace("&iacute;", "í")
    text = text.replace("&oacute;", "ó").replace("&uacute;", "ú").replace("&ntilde;", "ñ")
    return re.sub(r"\s+", " ", text).strip()


def _extract_blocks_from_pdf_bytes(pdf_bytes: bytes) -> List[str]:
    """
    Extrae los bloques de texto en orden de lectura directamente desde los streams PDF,
    soportando tanto streams sin comprimir (%PDF-1.3 de Swiss Medical) como FlateDecode,
    y usando pypdf como respaldo adicional si está instalado.
    """
    blocks: List[str] = []

    for match in re.finditer(b"stream\r?\n(.*?)\r?\nendstream", pdf_bytes, re.S):
        raw_stream = match.group(1)
        decoded_stream: Optional[bytes] = None
        if b"BT" in raw_stream and b"Tj" in raw_stream:
            decoded_stream = raw_stream
        else:
            try:
                decompressed = zlib.decompress(raw_stream)
                if b"BT" in decompressed:
                    decoded_stream = decompressed
            except Exception:
                pass

        if decoded_stream:
            text_content = decoded_stream.decode("latin1", errors="ignore")
            for bt_block in re.findall(r"BT\s+(.*?)\s+ET", text_content, re.S):
                tj_parts = [
                    _unescape_pdf_string(part)
                    for part in re.findall(r"\((.*?)\)\s*Tj", bt_block)
                ]
                joined = " ".join(p for p in tj_parts if p)
                if joined:
                    blocks.append(joined)

    if blocks:
        return blocks

    # Fallback con pypdf si el PDF utiliza otra codificación de fuentes
    try:
        import io
        from pypdf import PdfReader  # type: ignore

        reader = PdfReader(io.BytesIO(pdf_bytes))
        lines: List[str] = []
        for page in reader.pages:
            page_text = page.extract_text() or ""
            for line in page_text.splitlines():
                clean = line.strip()
                if clean:
                    lines.append(clean)
        return lines
    except Exception:
        return []


def parse_money_amount(raw: str) -> Optional[float]:
    """
    Convierte importes como '$ 62350', '$ 457,650.00' o '$ 42.428,45' a float.
    """
    if not raw:
        return None
    cleaned = re.sub(r"[^\d.,-]", "", raw).strip()
    if not cleaned or cleaned == "-":
        return None

    # Caso con coma y punto
    if "," in cleaned and "." in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            # Formato europeo/argentino: 42.428,45
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            # Formato anglosajón: 457,650.00
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        # Si termina con 2 decimales tras la coma (ej. 511108,51)
        parts = cleaned.split(",")
        if len(parts[-1]) == 2:
            cleaned = "".join(parts[:-1]) + "." + parts[-1]
        else:
            cleaned = cleaned.replace(",", "")

    try:
        return float(cleaned)
    except ValueError:
        return None


def _parse_socio_integrante(raw: str) -> Tuple[Optional[int], Optional[int]]:
    """
    Separa '0214421/01' o '214421/01' en (214421, 1).
    """
    match = re.search(r"0*(\d+)\s*/\s*0*(\d+)", raw)
    if not match:
        return None, None
    return int(match.group(1)), int(match.group(2))


def _parse_prestacion(raw: str) -> Tuple[Optional[int], str]:
    """
    Separa '12099 - FilmArray panel meningitis encefalitis' en (12099, 'FilmArray panel meningitis encefalitis').
    """
    match = re.match(r"^\s*(\d+)\s*-\s*(.+)$", raw)
    if match:
        return int(match.group(1)), match.group(2).strip()
    return None, raw.strip()


def parse_formulario_pdf(source: Union[str, Path, bytes]) -> PdfFormularioData:
    """
    Extrae todos los campos estructurados del 'FORMULARIO DE EXCEPCIONES' en PDF.
    Acepta una ruta de archivo (str / Path) o los bytes del PDF directamente.
    """
    if isinstance(source, (str, Path)):
        pdf_bytes = Path(source).read_bytes()
    else:
        pdf_bytes = source

    blocks = _extract_blocks_from_pdf_bytes(pdf_bytes)
    result = PdfFormularioData()

    # Detectar tipo de excepción por el título de sección
    full_text = "\n".join(blocks)
    if re.search(r"Excepci[oó]n\s+al\s+Pres?t?ador", full_text, re.I):
        result.tipo_excepcion = "Al Prestador"
    elif re.search(r"Excepci[oó]n\s+al\s+(Cliente|Afiliado|Socio)", full_text, re.I):
        result.tipo_excepcion = "Al Cliente"

    # Recorrer bloques etiqueta -> valor
    current_section = ""
    i = 0
    while i < len(blocks):
        token = blocks[i].strip()
        next_val = blocks[i + 1].strip() if (i + 1) < len(blocks) else ""

        # Actualizar sección activa para desambiguar etiquetas repetidas (ej. 'Apellido y Nombre')
        if token in (
            "Datos del Prestador",
            "Datos del Afiliado",
            "Detalle de la excepción a otorgar",
        ) or token.startswith("Excepción al"):
            current_section = token
            i += 1
            continue

        if token == "Nro Trámite:":
            result.nro_tramite = next_val
            i += 2
            continue
        if token == "Fecha:" and not result.fecha_formulario:
            result.fecha_formulario = next_val
            i += 2
            continue
        if token == "Solicitante:":
            result.solicitante = next_val
            i += 2
            continue
        if token == "Descripción:":
            result.motivo_excepcion = next_val
            i += 2
            continue
        if token == "Razón:":
            result.razon = next_val
            i += 2
            continue
        if token == "Código:" and "Prestador" in current_section:
            if next_val.isdigit():
                result.codigo_prestador = int(next_val)
            i += 2
            continue
        if token == "Apellido y Nombre / Razón Social:":
            result.nombre_prestador = next_val
            i += 2
            continue
        if token == "Contra/Inte:":
            socio, integrante = _parse_socio_integrante(next_val)
            result.numero_socio = socio
            result.numero_integrante = integrante
            i += 2
            continue
        if token == "Apellido y Nombre:" and "Afiliado" in current_section:
            result.afiliado_nombre = next_val
            i += 2
            continue
        if token == "Prestación:":
            cod, nom = _parse_prestacion(next_val)
            result.codigo_prestacion = cod
            result.nombre_prestacion = nom
            i += 2
            continue
        if token == "Importe de la excepción:":
            result.importe_excepcion = parse_money_amount(next_val)
            i += 2
            continue
        if token == "Importe presupuestado:":
            result.importe_presupuestado = parse_money_amount(next_val)
            i += 2
            continue
        if token == "Importe a liquidar:":
            result.importe_a_liquidar = parse_money_amount(next_val)
            i += 2
            continue
        if token == "Cantidad:":
            if next_val.isdigit():
                result.cantidad = int(next_val)
            i += 2
            continue
        if token == "Fecha desde:":
            result.vigencia_desde = next_val
            i += 2
            continue
        if token == "Fecha hasta:":
            result.vigencia_hasta = next_val
            i += 2
            continue

        i += 1

    return result
