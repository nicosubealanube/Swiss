"""
Genera un único archivo autocontenido 'Automatizacion_ABM_Convenios.html'
que funciona 100% offline/local con doble clic en Windows (sin .exe y sin hosting externo).
Incluye el logo de Swiss Medical embebido en Base64, el motor extractor de PDF/Email,
la interfaz completa y la guía de automatización interna con Office 365 (Excel + Power Automate).
"""

import base64
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent


def build_standalone_html() -> Path:
    logo_bytes = (BASE_DIR / "public" / "logo.png").read_bytes()
    logo_b64 = base64.b64encode(logo_bytes).decode("ascii")
    logo_data_uri = f"data:image/png;base64,{logo_b64}"

    css_content = (BASE_DIR / "public" / "styles.css").read_text(encoding="utf-8")
    parser_js = (BASE_DIR / "public" / "parser.js").read_text(encoding="utf-8")
    app_js = (BASE_DIR / "public" / "app.js").read_text(encoding="utf-8")

    # Asegurar que el copiado al portapapeles funcione tanto en https:// como en file:///
    clipboard_helper = """
  async function safeCopyText(text) {
    if (navigator.clipboard && window.isSecureContext) {
      try {
        await navigator.clipboard.writeText(text);
        return;
      } catch (e) {}
    }
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.left = "-9999px";
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
  }
"""
    app_js = app_js.replace(
        "await navigator.clipboard.writeText(tsv);",
        "await safeCopyText(tsv);",
    ).replace(
        "await navigator.clipboard.writeText(OFFICE_SCRIPT_CODE);",
        "await safeCopyText(OFFICE_SCRIPT_CODE);",
    ).replace(
        "function showToast(message) {",
        clipboard_helper + "\n  function showToast(message) {",
    )

    extra_css = """
.local-security-pill {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  background: #ecfdf5;
  border: 1px solid #a7f3d0;
  color: #065f46;
  font-size: 12px;
  font-weight: 600;
  padding: 5px 11px;
  border-radius: 999px;
}
"""

    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>ABM de Convenios - Carga Automática de Excepciones | Swiss Medical</title>
  <style>
{css_content}
{extra_css}
  </style>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>
<body>
  <header class="topbar">
    <div class="topbar-brand">
      <img src="{logo_data_uri}" alt="Swiss Medical" class="swiss-logo" />
      <div>
        <h1>Automatización ABM de Convenios</h1>
        <p class="subtitle">Casilla: <strong>ABM.Convenios@swissmedical.com.ar</strong> &rarr; Excel en SharePoint</p>
      </div>
    </div>
    <div class="topbar-actions">
      <span class="local-security-pill" title="Todos los archivos PDF y correos se procesan localmente en este equipo sin enviarse a ningún servidor externo.">
        &#128274; Ejecución 100% Local y Privada
      </span>
      <button type="button" id="btnModeManual" class="btn btn-mode active">
        Carga Asistida
      </button>
      <button type="button" id="btnModeAuto" class="btn btn-mode">
        Automatización Microsoft 365 (Excel + Power Automate)
      </button>
    </div>
  </header>

  <main class="container">
    <div id="toastBanner" class="toast hidden"></div>

    <!-- ================================================================== -->
    <!-- VISTA 1: GUÍA PASO A PASO 100% INTERNA EN MICROSOFT 365            -->
    <!-- ================================================================== -->
    <section id="autoSetupView" class="hidden">
      <div class="card guide-hero">
        <h2>Automatización 100% Interna en Microsoft 365 (Outlook &rarr; SharePoint Excel)</h2>
        <p>
          Con esta configuración, cada vez que ingresa un correo con un formulario PDF a
          <strong>ABM.Convenios@swissmedical.com.ar</strong>, los datos se cargan
          <strong>automáticamente en el Excel de SharePoint</strong> dentro del entorno corporativo de Microsoft 365,
          sin pasar por servidores externos.
        </p>
      </div>

      <div id="tabMethodExcel" class="method-content">
        <div class="guide-grid">
          <!-- Parte 1: En el Excel de SharePoint -->
          <div class="card guide-card">
            <div class="step-header">
              <span class="step-badge">Paso 1 (En Excel)</span>
              <h3>Cargar el Script Extractor en el Excel de SharePoint</h3>
            </div>
            <p class="help-text">
              Este script interno de Excel lee el correo y el formulario PDF adjunto y escribe la nueva fila en la planilla:
            </p>
            <ol class="friendly-steps">
              <li>
                Abrí en tu navegador el archivo <strong>Excel en SharePoint</strong> donde se registran las excepciones.
              </li>
              <li>
                En el menú superior del Excel, hacé clic en la pestaña
                <strong>Automatizar</strong> (<em>Automate</em>) y seleccioná <strong>Nuevo script</strong> (<em>New Script</em>).
              </li>
              <li>
                En el panel de código que se abre a la derecha, <strong>borrá el contenido inicial</strong>.
              </li>
              <li>
                Hacé clic en el botón verde para copiar el script y pegalo en el panel con <strong>Ctrl + V</strong>:
                <div class="code-copy-box">
                  <button type="button" id="btnCopyExcelScript" class="btn btn-success">
                    Copiar Código para Excel
                  </button>
                  <button type="button" id="btnToggleScriptPreview" class="btn btn-outline">
                    Ver / Ocultar código
                  </button>
                </div>
                <pre id="excelScriptPre" class="code-preview hidden"></pre>
              </li>
              <li>
                Nombrá el script como <strong><code>CargarExcepcionABM</code></strong> y hacé clic en
                <strong>Guardar script</strong> (<em>Save script</em>).
              </li>
            </ol>
          </div>

          <!-- Parte 2: En Power Automate -->
          <div class="card guide-card">
            <div class="step-header">
              <span class="step-badge">Paso 2 (En Power Automate)</span>
              <h3>Vincular la casilla ABM.Convenios con el Excel</h3>
            </div>
            <p class="help-text">
              Regla automática dentro de Microsoft 365 para que cada email entrante dispare el script del Excel:
            </p>
            <ol class="friendly-steps">
              <li>
                Ingresá a
                <a href="https://make.powerautomate.com" target="_blank" rel="noopener"><strong>make.powerautomate.com</strong></a>
                con tu cuenta corporativa de Swiss Medical.
              </li>
              <li>
                En el menú izquierdo seleccioná <strong>Crear</strong> &rarr;
                <strong>Flujo de nube automatizado</strong>.
              </li>
              <li>
                Asignale el nombre <code>ABM Convenios a Excel</code>, buscá el disparador
                <strong>"Cuando llega un correo electrónico nuevo a un buzón compartido (V2)"</strong> (Office 365 Outlook) y presioná <em>Crear</em>.
              </li>
              <li>
                Configurá el bloque del correo con:
                <ul>
                  <li><strong>Dirección del buzón original:</strong> <code>ABM.Convenios@swissmedical.com.ar</code></li>
                  <li><strong>Carpeta:</strong> <code>Inbox</code></li>
                  <li><strong>Incluir datos adjuntos:</strong> <strong>Sí</strong></li>
                  <li><strong>Solo con datos adjuntos:</strong> <strong>Sí</strong></li>
                </ul>
              </li>
              <li>
                Agregá un <strong>(+) Nuevo paso</strong>, elegí <strong>Excel Online (Empresas)</strong> y seleccioná la acción
                <strong>"Ejecutar script" (Run script)</strong>.
              </li>
              <li>
                Seleccioná el archivo Excel de Excepciones en SharePoint y elegí el script <strong><code>CargarExcepcionABM</code></strong>.
              </li>
              <li>
                Completá los 3 parámetros con el contenido dinámico del correo:
                <ul>
                  <li><strong><code>emailBody</code>:</strong> <strong>Cuerpo (Body)</strong> del correo.</li>
                  <li><strong><code>receivedDate</code>:</strong> <strong>Hora de recepción (Received Time)</strong>.</li>
                  <li><strong><code>pdfBase64</code>:</strong> <strong>Contenido de datos adjuntos (Attachments Content)</strong>.</li>
                </ul>
              </li>
              <li>
                Hacé clic en <strong>Guardar</strong>. A partir de ese momento, cada correo nuevo se procesará automáticamente.
              </li>
            </ol>
          </div>
        </div>
      </div>
    </section>

    <!-- ================================================================== -->
    <!-- VISTA 2: CARGA ASISTIDA LOCAL                                      -->
    <!-- ================================================================== -->
    <div id="manualAssistView">
      <section class="config-bar card">
        <div class="config-item">
          <label for="decimalSeparatorSelect">Separador decimal al pegar en Excel</label>
          <select id="decimalSeparatorSelect">
            <option value=",">Coma decimal (ej. 457650,00 - Excel Español)</option>
            <option value=".">Punto decimal (ej. 457650.00)</option>
          </select>
        </div>
        <div class="config-item">
          <label for="defaultGerenciaInput">Gerencia Autorizante por defecto</label>
          <input type="text" id="defaultGerenciaInput" value="Gerencia de Gestión de Prestadores Interior" />
        </div>
        <div class="config-item">
          <label for="fechaCargaInput">Fecha de la carga</label>
          <input type="text" id="fechaCargaInput" placeholder="DD/MM/YYYY" />
        </div>
      </section>

      <div class="workspace-grid">
        <section class="card">
          <div class="step-header">
            <span class="step-badge">Paso 1</span>
            <h2>Datos del Email y Formulario PDF Adjunto</h2>
          </div>

          <div id="pdfDropzone" class="dropzone">
            <input type="file" id="pdfFileInput" accept=".pdf,.eml" hidden />
            <div class="dropzone-content">
              <div class="dropzone-icon">PDF</div>
              <div>
                <p class="dropzone-title" id="pdfStatusTitle">
                  Arrastrá acá el archivo <strong>.PDF</strong> adjunto ("Formulario de Excepciones")
                </p>
                <p class="dropzone-sub" id="pdfStatusSub">
                  O hacé clic para seleccionarlo desde tu computadora
                </p>
              </div>
            </div>
          </div>

          <div class="form-group">
            <div class="label-row">
              <label for="emailBodyInput">Cuerpo del Email (pegar desde Outlook con Ctrl+V)</label>
              <button type="button" id="btnClearInputs" class="btn-link">Limpiar campos</button>
            </div>
            <textarea
              id="emailBodyInput"
              rows="7"
              placeholder="Copiá y pegá acá el texto o la tabla del cuerpo del mail (donde figura Convenio (1) y Valor de referencia (2))..."
            ></textarea>
          </div>

          <div class="inline-fields">
            <div class="form-group">
              <label for="manualFechaMailInput">Fecha de recepción del mail (DD/MM/YYYY)</label>
              <input
                type="text"
                id="manualFechaMailInput"
                placeholder="Se detecta sola o ingresala"
              />
            </div>
            <div class="form-group">
              <label for="manualConvenioInput">Número de convenio (del mail)</label>
              <input
                type="text"
                id="manualConvenioInput"
                placeholder="Auto desde Convenio (1)"
              />
            </div>
            <div class="form-group">
              <label for="manualValorRefInput">Valor de referencia (del mail)</label>
              <input
                type="text"
                id="manualValorRefInput"
                placeholder="Auto desde Valor de referencia (2)"
              />
            </div>
          </div>
        </section>

        <section class="card">
          <div class="step-header">
            <span class="step-badge">Paso 2</span>
            <h2>Verificación de Campos Extraídos (22 Columnas)</h2>
          </div>
          <p class="help-text">
            Revisá los datos extraídos automáticamente del <strong>Cuerpo del Mail</strong> y del <strong>Formulario PDF</strong>.
            Podés editar cualquier valor antes de sumarlo a la tabla.
          </p>

          <div id="previewFieldsGrid" class="preview-grid"></div>

          <div class="card-footer-actions">
            <button type="button" id="btnAddRow" class="btn btn-primary btn-large">
              + Agregar Fila a la Tabla de Carga
            </button>
          </div>
        </section>
      </div>

      <section class="card batch-section">
        <div class="batch-header">
          <div>
            <div class="step-header">
              <span class="step-badge">Paso 3</span>
              <h2>Filas Listas para Pasar al Excel de SharePoint (<span id="rowCountBadge">0</span>)</h2>
            </div>
            <p class="help-text">
              Hacé clic en <strong>"Copiar filas para SharePoint"</strong>, andá al Excel abierto en SharePoint,
              parate en la celda <strong>Fecha de recepción del mail</strong> de la nueva fila y presioná <strong>Ctrl + V</strong>
              (el <em>Número de ingreso</em> lo numera solo el Excel).
            </p>
          </div>
          <div class="batch-actions">
            <button type="button" id="btnCopyForSharePoint" class="btn btn-success" disabled>
              Copiar filas para SharePoint (Ctrl+V)
            </button>
            <button type="button" id="btnCopyWithHeaders" class="btn btn-outline" disabled>
              Copiar con cabeceras
            </button>
            <button type="button" id="btnDownloadXlsx" class="btn btn-secondary" disabled>
              Descargar Excel (.xlsx)
            </button>
            <button type="button" id="btnClearTable" class="btn btn-danger-outline" disabled>
              Vaciar tabla
            </button>
          </div>
        </div>

        <div class="table-wrapper">
          <table id="batchTable">
            <thead>
              <tr id="batchTableHeader"></tr>
            </thead>
            <tbody id="batchTableBody">
              <tr class="empty-row">
                <td colspan="23">
                  Todavía no hay filas en el lote. Subí un archivo PDF y pegá el cuerpo del mail en el Paso 1.
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </section>
    </div>
  </main>

  <script>
{parser_js}
  </script>
  <script>
{app_js}
  </script>
</body>
</html>
"""

    target_path = BASE_DIR / "Automatizacion_ABM_Convenios.html"
    target_path.write_text(html, encoding="utf-8")
    return target_path


if __name__ == "__main__":
    out = build_standalone_html()
    print(f"Archivo único generado exitosamente: {out} ({out.stat().st_size // 1024} KB)")
