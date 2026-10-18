"""Orquestador que combina los datos del Cuerpo del Email + Formulario PDF en una fila de Excel."""

from datetime import datetime
from typing import Optional, Tuple

import config
from src.models import EmailExtractedData, ExcelRowRecord, PdfFormularioData


def resolve_gerencia_and_usuario(
    email_data: EmailExtractedData,
    pdf_data: PdfFormularioData,
) -> Tuple[str, str]:
    """
    Determina 'Gerencia autorizante' y 'Usuario de carga' a partir del solicitante
    del formulario PDF o remitente del correo, con respaldo en los valores por defecto.
    """
    solicitante_key = (pdf_data.solicitante or email_data.sender_name or "").strip().lower()
    email_key = (email_data.sender_email or "").strip().lower()

    if solicitante_key in config.SOLICITANTES_MAP:
        gerencia, usuario = config.SOLICITANTES_MAP[solicitante_key]
    elif email_key in config.SOLICITANTES_MAP:
        gerencia, usuario = config.SOLICITANTES_MAP[email_key]
    else:
        gerencia = email_data.gerencia_firma or config.DEFAULT_GERENCIA_AUTORIZANTE
        usuario = config.DEFAULT_USUARIO_CARGA
        if not usuario and solicitante_key:
            parts = [p for p in solicitante_key.split() if p.isalpha()]
            if len(parts) >= 2:
                # Inicial del primer nombre + inicial del apellido (ej. Maria Elena Velez -> MV)
                usuario = (parts[0][0] + parts[-1][0]).upper()
            elif len(parts) == 1:
                usuario = parts[0][:2].upper()

    return gerencia, usuario


def build_excel_record(
    email_data: EmailExtractedData,
    pdf_data: PdfFormularioData,
    numero_ingreso: int,
    fecha_carga: Optional[str] = None,
) -> ExcelRowRecord:
    """
    Construye el registro final para el Excel combinando:
    - Cuerpo del mail: Fecha de recepción del mail, Número de convenio, Valor de referencia.
    - Formulario PDF: Vigencias, Socio/Integrante, Prestador, Prestación, Importes, Trámite, Tipo y Motivo.
    - Campos de carga: Gerencia autorizante, Usuario de carga, Fecha de la carga.
    """
    gerencia, usuario = resolve_gerencia_and_usuario(email_data, pdf_data)
    fecha_carga_str = fecha_carga or datetime.now().strftime("%d/%m/%Y")

    # Fecha de recepción del mail (si no vino en cabecera del mail, usa la fecha del formulario sin hora)
    fecha_recepcion = email_data.fecha_recepcion
    if not fecha_recepcion and pdf_data.fecha_formulario:
        fecha_recepcion = pdf_data.fecha_formulario.split()[0]

    # Valor de referencia: prioriza el Cuerpo del Mail ('Valor de referencia (2)')
    valor_ref = email_data.valor_referencia
    if (
        valor_ref is None
        and pdf_data.importe_presupuestado is not None
        and pdf_data.importe_excepcion is not None
    ):
        valor_ref = round(pdf_data.importe_presupuestado - pdf_data.importe_excepcion, 2)

    n_llamado = pdf_data.nro_tramite or email_data.tramite_excepcion_body

    return ExcelRowRecord(
        numero_ingreso=numero_ingreso,
        fecha_recepcion_mail=fecha_recepcion,
        vigencia_desde=pdf_data.vigencia_desde,
        vigencia_hasta=pdf_data.vigencia_hasta,
        numero_de_socio=pdf_data.numero_socio,
        numero_socio_integrante=pdf_data.numero_integrante,
        codigo_prestador=pdf_data.codigo_prestador,
        nombre_prestador=pdf_data.nombre_prestador,
        numero_convenio=email_data.numero_convenio,
        codigo_prestacion=pdf_data.codigo_prestacion,
        nombre_prestacion=pdf_data.nombre_prestacion,
        valor_referencia=valor_ref,
        valor_presupuestado=pdf_data.importe_presupuestado,
        valor_excepcion=pdf_data.importe_excepcion,
        valor_a_liquidar=pdf_data.importe_a_liquidar,
        gerencia_autorizante=gerencia,
        usuario_carga=usuario,
        fecha_carga=fecha_carga_str,
        n_de_llamado=n_llamado,
        tipo_de_excepcion=pdf_data.tipo_excepcion,
        motivo_excepcion=pdf_data.motivo_excepcion,
        autorizaciones_si_no="",
        detalle="",
    )
