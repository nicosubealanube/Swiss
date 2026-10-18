# Automatización de Carga de Excepciones - ABM Convenios (Swiss Medical)

Proyecto en Python para automatizar la lectura de correos entrantes en la casilla **`ABM.Convenios@swissmedical.com.ar`**, extraer los datos requeridos del **cuerpo del email** y del **Formulario de Excepciones (.pdf)** adjunto, e insertar automáticamente la fila correspondiente en el archivo **Excel en SharePoint** (o archivo `.xlsx` sincronizado/local).

---

## Origen de cada columna en el Excel (23 columnas)

*(Nota: se desestimó la columna `Tipo de convenio` según lo acordado con el equipo operativo).*

| # | Columna Excel | Origen | Ejemplo (Caso Ledesma) |
|---|---|---|---|
| 1 | `Número de ingreso` | Autoincremental (`Último + 1`) | `50608` |
| 2 | `Fecha de recepción del mail` | **Email** (Fecha de recepción) | `16/10/2026` |
| 3 | `Vigencia desde` | **Formulario PDF** (`Fecha desde:`) | `01/10/2026` |
| 4 | `Vigencia hasta` | **Formulario PDF** (`Fecha hasta:`) | `31/10/2026` |
| 5 | `Numero_de_Socio` | **Formulario PDF** (`Contra/Inte:` antes de `/`) | `214421` |
| 6 | `Numero del socio_integrante` | **Formulario PDF** (`Contra/Inte:` después de `/`) | `1` |
| 7 | `Código del prestador` | **Formulario PDF** (`Datos del Prestador -> Código:`) | `1212` |
| 8 | `Nombre del prestador` | **Formulario PDF** (`Apellido y Nombre / Razón Social:`) | `Sanatorio Allende` |
| 9 | `Número de convenio` | **Cuerpo del Email** (`Convenio (1)`) | `12421` |
| 10 | `Código de la Prestación` | **Formulario PDF** (`Prestación:` código numérico) | `12099` |
| 11 | `Nombre de la Prestación` | **Formulario PDF** (`Prestación:` descripción) | `FilmArray panel meningitis encefalitis` |
| 12 | `Valor de referencia` | **Cuerpo del Email** (`Valor de referencia (2)`) | `457650.0` |
| 13 | `Valor presupuestado` | **Formulario PDF** (`Importe presupuestado:`) | `520000.0` |
| 14 | `Valor de la excepción` | **Formulario PDF** (`Importe de la excepción:`) | `62350.0` |
| 15 | `Valor a liquidar` | **Formulario PDF** (`Importe a liquidar:`) | `520000.0` |
| 16 | `Gerencia autorizante` | Deducido del solicitante/firma o valor fijo | `Gerencia de Gestión de Prestadores Interior` |
| 17 | `Usuario de carga` | Iniciales del solicitante (ej. Maria Elena Velez) | `MV` |
| 18 | `Fecha de la carga` | Fecha de ejecución del proceso | `18/10/2026` |
| 19 | `N_de_llamado` | **Formulario PDF** (`Nro Trámite:`) | `302407236` |
| 20 | `Tipo_de_excepcion` | **Formulario PDF** (`Excepción al Prestador / Convenio`) | `Al Prestador` |
| 21 | `Motivo de la Excepción` | **Formulario PDF** (`Descripción:`) | `Prestación no cargada en el convenio del Prestador` |
| 22 | `Autorizaciones si/no?` | Se deja vacío | `""` |
| 23 | `Detalle` | Se deja vacío | `""` |

---

## Estructura del Proyecto

* `config.py`: Configuración centralizada, variables de entorno y mapeo de solicitantes/gerencias.
* `main.py`: Punto de entrada por línea de comandos (`--demo`, `--eml`, `--run-mailbox`, `--watch`).
* `src/models.py`: Definición de las 23 columnas del Excel y estructuras de datos.
* `src/email_parser.py`: Extractor de tablas HTML, fechas y metadatos del cuerpo del correo.
* `src/pdf_parser.py`: Extractor del PDF *Formulario de Excepciones*.
* `src/processor.py`: Orquestador que cruza los datos del mail + PDF y aplica las reglas de negocio.
* `src/email_client.py`: Conector con `ABM.Convenios@swissmedical.com.ar` vía **Microsoft Graph API** o **IMAP**.
* `src/excel_writer.py`: Escritor en **SharePoint Excel Online** (vía Graph API) y en archivo `.xlsx` local/sincronizado.
* `samples/`: Incluye el PDF real de ejemplo (`Excep. LEDESMA, GERMAN.pdf`) y el correo de ejemplo (`email_ejemplo.eml`).
* `tests/test_pipeline.py`: Suite de pruebas automatizadas.

---

## Ejecución Rápida

### 1. Probar con el caso real de ejemplo (sin necesidad de credenciales)
```bash
python3 main.py --demo
```
Esto procesa el correo y el PDF de `LEDESMA, GERMAN ARIEL`, muestra por pantalla las 23 columnas extraídas y genera el archivo Excel en `output/excepciones_abm.xlsx`.

### 2. Ejecutar los tests unitarios e integrales
```bash
python3 -m unittest discover -s tests -v
```

### 3. Conectar con la casilla `ABM.Convenios@swissmedical.com.ar` y SharePoint
1. Copiar `.env.example` a `.env`:
   ```bash
   cp .env.example .env
   ```
2. Completar las credenciales de **Microsoft Graph API** (`AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`) y los IDs del archivo en **SharePoint** (`SHAREPOINT_SITE_ID`, `SHAREPOINT_DRIVE_ID`, `SHAREPOINT_ITEM_ID`, `SHAREPOINT_TABLE_NAME`), o bien configurar una ruta de carpeta sincronizada con OneDrive/SharePoint en `LOCAL_EXCEL_PATH`.
3. Ejecutar una pasada o dejar monitoreando continuamente:
   ```bash
   # Procesar correos no leídos una vez
   python3 main.py --run-mailbox

   # Monitorear la bandeja cada 60 segundos
   python3 main.py --watch 60
   ```
