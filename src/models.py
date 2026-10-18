"""Modelos de datos para el procesamiento de excepciones ABM Convenios."""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Union

# Columnas exactas del Excel en SharePoint (habiendo desestimado 'Tipo de convenio')
EXCEL_COLUMNS: List[str] = [
    "Número de ingreso",
    "Fecha de recepción del mail",
    "Vigencia desde",
    "Vigencia hasta",
    "Numero_de_Socio",
    "Numero del socio_integrante",
    "Código del prestador",
    "Nombre del prestador",
    "Número de convenio",
    "Código de la Prestación",
    "Nombre de la Prestación",
    "Valor de referencia",
    "Valor presupuestado",
    "Valor de la excepción",
    "Valor a liquidar",
    "Gerencia autorizante",
    "Usuario de carga",
    "Fecha de la carga",
    "N_de_llamado",
    "Tipo_de_excepcion",
    "Motivo de la Excepción",
    "Autorizaciones si/no?",
    "Detalle",
]


@dataclass
class EmailExtractedData:
    """Datos extraídos del correo electrónico (cabeceras y cuerpo)."""

    message_id: str = ""
    subject: str = ""
    sender_name: str = ""
    sender_email: str = ""
    fecha_recepcion: str = ""          # Formato DD/MM/YYYY
    numero_convenio: Optional[int] = None
    valor_referencia: Optional[float] = None
    tramite_excepcion_body: str = ""
    gerencia_firma: str = ""
    pdf_attachments: List[Dict[str, bytes]] = field(default_factory=list)


@dataclass
class PdfFormularioData:
    """Datos extraídos del PDF adjunto ('FORMULARIO DE EXCEPCIONES')."""

    nro_tramite: str = ""
    fecha_formulario: str = ""
    solicitante: str = ""
    tipo_excepcion: str = ""           # Ej. 'Al Prestador' / 'Al Cliente'
    motivo_excepcion: str = ""         # Ej. 'Prestación no cargada en el convenio del Prestador'
    razon: str = ""
    codigo_prestador: Optional[int] = None
    nombre_prestador: str = ""
    numero_socio: Optional[int] = None
    numero_integrante: Optional[int] = None
    afiliado_nombre: str = ""
    codigo_prestacion: Optional[int] = None
    nombre_prestacion: str = ""
    importe_excepcion: Optional[float] = None
    importe_presupuestado: Optional[float] = None
    importe_a_liquidar: Optional[float] = None
    cantidad: Optional[int] = None
    vigencia_desde: str = ""           # Formato DD/MM/YYYY
    vigencia_hasta: str = ""           # Formato DD/MM/YYYY
    resolucion: str = ""
    observaciones: str = ""


@dataclass
class ExcelRowRecord:
    """Representa una fila lista para insertarse en la tabla de Excel en SharePoint."""

    numero_ingreso: int
    fecha_recepcion_mail: str
    vigencia_desde: str
    vigencia_hasta: str
    numero_de_socio: Optional[int]
    numero_socio_integrante: Optional[int]
    codigo_prestador: Optional[int]
    nombre_prestador: str
    numero_convenio: Optional[Union[int, str]]
    codigo_prestacion: Optional[int]
    nombre_prestacion: str
    valor_referencia: Optional[float]
    valor_presupuestado: Optional[float]
    valor_excepcion: Optional[float]
    valor_a_liquidar: Optional[float]
    gerencia_autorizante: str
    usuario_carga: str
    fecha_carga: str
    n_de_llamado: Optional[Union[int, str]]
    tipo_de_excepcion: str
    motivo_excepcion: str
    autorizaciones_si_no: str = ""
    detalle: str = ""

    def to_list(self) -> List[Union[str, int, float, None]]:
        """Devuelve los valores ordenados exactamente según EXCEL_COLUMNS."""
        return [
            self.numero_ingreso,
            self.fecha_recepcion_mail,
            self.vigencia_desde,
            self.vigencia_hasta,
            self.numero_de_socio,
            self.numero_socio_integrante,
            self.codigo_prestador,
            self.nombre_prestador,
            self.numero_convenio if self.numero_convenio is not None else "",
            self.codigo_prestacion,
            self.nombre_prestacion,
            self.valor_referencia,
            self.valor_presupuestado,
            self.valor_excepcion,
            self.valor_a_liquidar,
            self.gerencia_autorizante,
            self.usuario_carga,
            self.fecha_carga,
            int(self.n_de_llamado) if str(self.n_de_llamado).isdigit() else self.n_de_llamado,
            self.tipo_de_excepcion,
            self.motivo_excepcion,
            self.autorizaciones_si_no,
            self.detalle,
        ]

    def to_dict(self) -> Dict[str, Union[str, int, float, None]]:
        """Devuelve un diccionario {Nombre de Columna: Valor}."""
        return dict(zip(EXCEL_COLUMNS, self.to_list()))
