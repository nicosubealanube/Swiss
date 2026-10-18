"""Tests integrales para validar la extracción del Cuerpo del Mail + PDF Adjunto -> Excel."""

from pathlib import Path
import tempfile
import unittest

import config
from main import ensure_sample_eml
from src.email_parser import parse_eml_file
from src.excel_writer import LocalExcelWriter
from src.models import EXCEL_COLUMNS
from src.pdf_parser import parse_formulario_pdf
from src.processor import build_excel_record


class TestAbmConveniosPipeline(unittest.TestCase):
    def setUp(self) -> None:
        self.eml_path = ensure_sample_eml()
        self.pdf_path = config.BASE_DIR / "samples" / "Excep. LEDESMA, GERMAN.pdf"

    def test_pdf_extraction_ledesma(self) -> None:
        """Verifica que todos los campos del formulario PDF real se extraigan con exactitud."""
        pdf_data = parse_formulario_pdf(self.pdf_path)

        self.assertEqual(pdf_data.nro_tramite, "302407236")
        self.assertEqual(pdf_data.solicitante, "Maria Elena Velez")
        self.assertEqual(pdf_data.tipo_excepcion, "Al Prestador")
        self.assertEqual(
            pdf_data.motivo_excepcion,
            "Prestación no cargada en el convenio del Prestador",
        )
        self.assertEqual(pdf_data.codigo_prestador, 1212)
        self.assertEqual(pdf_data.nombre_prestador, "Sanatorio Allende")
        self.assertEqual(pdf_data.numero_socio, 214421)
        self.assertEqual(pdf_data.numero_integrante, 1)
        self.assertEqual(pdf_data.codigo_prestacion, 12099)
        self.assertEqual(
            pdf_data.nombre_prestacion,
            "FilmArray panel meningitis encefalitis",
        )
        self.assertEqual(pdf_data.importe_excepcion, 62350.0)
        self.assertEqual(pdf_data.importe_presupuestado, 520000.0)
        self.assertEqual(pdf_data.importe_a_liquidar, 520000.0)
        self.assertEqual(pdf_data.vigencia_desde, "01/10/2026")
        self.assertEqual(pdf_data.vigencia_hasta, "31/10/2026")

    def test_email_and_pdf_to_excel_row(self) -> None:
        """Verifica que la combinación del cuerpo del mail y el PDF genere las 23 columnas exactas."""
        email_data = parse_eml_file(self.eml_path)
        self.assertEqual(email_data.fecha_recepcion, "16/10/2026")
        self.assertEqual(email_data.numero_convenio, 12421)
        self.assertEqual(email_data.valor_referencia, 457650.0)
        self.assertEqual(len(email_data.pdf_attachments), 1)

        pdf_data = parse_formulario_pdf(email_data.pdf_attachments[0]["content"])
        record = build_excel_record(
            email_data=email_data,
            pdf_data=pdf_data,
            numero_ingreso=50608,
            fecha_carga="18/10/2026",
        )

        row_dict = record.to_dict()
        self.assertEqual(list(row_dict.keys()), EXCEL_COLUMNS)
        self.assertNotIn("Tipo de convenio", row_dict)
        self.assertEqual(row_dict["Número de ingreso"], 50608)
        self.assertEqual(row_dict["Fecha de recepción del mail"], "16/10/2026")
        self.assertEqual(row_dict["Vigencia desde"], "01/10/2026")
        self.assertEqual(row_dict["Vigencia hasta"], "31/10/2026")
        self.assertEqual(row_dict["Numero_de_Socio"], 214421)
        self.assertEqual(row_dict["Numero del socio_integrante"], 1)
        self.assertEqual(row_dict["Código del prestador"], 1212)
        self.assertEqual(row_dict["Nombre del prestador"], "Sanatorio Allende")
        self.assertEqual(row_dict["Número de convenio"], 12421)
        self.assertEqual(row_dict["Código de la Prestación"], 12099)
        self.assertEqual(
            row_dict["Nombre de la Prestación"],
            "FilmArray panel meningitis encefalitis",
        )
        self.assertEqual(row_dict["Valor de referencia"], 457650.0)
        self.assertEqual(row_dict["Valor presupuestado"], 520000.0)
        self.assertEqual(row_dict["Valor de la excepción"], 62350.0)
        self.assertEqual(row_dict["Valor a liquidar"], 520000.0)
        self.assertEqual(
            row_dict["Gerencia autorizante"],
            "Gerencia de Gestión de Prestadores Interior",
        )
        self.assertEqual(row_dict["Usuario de carga"], "MV")
        self.assertEqual(row_dict["Fecha de la carga"], "18/10/2026")
        self.assertEqual(row_dict["N_de_llamado"], 302407236)
        self.assertEqual(row_dict["Tipo_de_excepcion"], "Al Prestador")
        self.assertEqual(
            row_dict["Motivo de la Excepción"],
            "Prestación no cargada en el convenio del Prestador",
        )
        self.assertEqual(row_dict["Autorizaciones si/no?"], "")
        self.assertEqual(row_dict["Detalle"], "")

    def test_excel_writer_roundtrip(self) -> None:
        """Verifica la escritura y lectura autoincremental en archivo .xlsx."""
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_file = Path(tmpdir) / "test_excepciones.xlsx"
            writer = LocalExcelWriter(xlsx_file)

            num1 = writer.get_next_ingreso_number(default_base=50607)
            self.assertEqual(num1, 50608)

            email_data = parse_eml_file(self.eml_path)
            pdf_data = parse_formulario_pdf(email_data.pdf_attachments[0]["content"])
            rec1 = build_excel_record(email_data, pdf_data, numero_ingreso=num1)
            writer.append_record(rec1)

            num2 = writer.get_next_ingreso_number(default_base=50607)
            self.assertEqual(num2, 50609)

            rows = writer.read_all_rows()
            self.assertEqual(len(rows), 2)
            self.assertEqual(rows[0], EXCEL_COLUMNS)
            self.assertEqual(rows[1][0], 50608)


if __name__ == "__main__":
    unittest.main()
