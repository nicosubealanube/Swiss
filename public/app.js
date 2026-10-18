/**
 * Lógica principal de la aplicación web ABM Convenios (Netlify Frontend)
 */
(function () {
  "use strict";

  const COLUMNS = window.SwissParser.EXCEL_COLUMNS;

  const COLUMN_SOURCE_META = {
    "Fecha de recepción del mail": { tag: "Mail", cls: "source-mail" },
    "Vigencia desde": { tag: "PDF", cls: "source-pdf" },
    "Vigencia hasta": { tag: "PDF", cls: "source-pdf" },
    "Numero_de_Socio": { tag: "PDF", cls: "source-pdf" },
    "Numero del socio_integrante": { tag: "PDF", cls: "source-pdf" },
    "Código del prestador": { tag: "PDF", cls: "source-pdf" },
    "Nombre del prestador": { tag: "PDF", cls: "source-pdf" },
    "Número de convenio": { tag: "Mail", cls: "source-mail" },
    "Código de la Prestación": { tag: "PDF", cls: "source-pdf" },
    "Nombre de la Prestación": { tag: "PDF", cls: "source-pdf", wide: true },
    "Valor de referencia": { tag: "Mail", cls: "source-mail" },
    "Valor presupuestado": { tag: "PDF", cls: "source-pdf" },
    "Valor de la excepción": { tag: "PDF", cls: "source-pdf" },
    "Valor a liquidar": { tag: "PDF", cls: "source-pdf" },
    "Gerencia autorizante": { tag: "Auto", cls: "source-auto", wide: true },
    "Usuario de carga": { tag: "Auto", cls: "source-auto" },
    "Fecha de la carga": { tag: "Auto", cls: "source-auto" },
    "N_de_llamado": { tag: "PDF", cls: "source-pdf" },
    "Tipo_de_excepcion": { tag: "PDF", cls: "source-pdf" },
    "Motivo de la Excepción": { tag: "PDF", cls: "source-pdf", wide: true },
    "Autorizaciones si/no?": { tag: "Vacío", cls: "source-auto" },
    "Detalle": { tag: "Vacío", cls: "source-auto" },
  };

  // Código completo de Office Script (TypeScript para Excel Online en SharePoint)
  const OFFICE_SCRIPT_CODE = `/**
 * Office Script para Excel en SharePoint - Automatización ABM Convenios (Swiss Medical)
 * Recibe el cuerpo del mail y el PDF adjunto desde Power Automate y agrega la fila en el Excel.
 */
function main(
  workbook: ExcelScript.Workbook,
  emailBody: string = "",
  receivedDate: string = "",
  pdfBase64: string = ""
) {
  const rawPdf = decodeBase64ToLatin1(pdfBase64);
  const blocks = extractBlocksFromPdf(rawPdf);
  const pdf = parseFormularioBlocks(blocks);
  const mail = parseEmailBody(emailBody, receivedDate);

  const hoy = formatTodayDDMMYYYY();
  const fechaRecepcion = mail.fechaRecepcion || (pdf.fechaFormulario ? pdf.fechaFormulario.split(" ")[0] : hoy);

  let valorRef: number | string = mail.valorReferencia;
  if (valorRef === "" && typeof pdf.importePresupuestado === "number" && typeof pdf.importeExcepcion === "number") {
    valorRef = Number((pdf.importePresupuestado - pdf.importeExcepcion).toFixed(2));
  }

  let usuario = "MV";
  if (pdf.solicitante) {
    const words = pdf.solicitante.trim().split(/\\s+/);
    if (words.length >= 2) {
      usuario = (words[0][0] + words[words.length - 1][0]).toUpperCase();
    }
  }
  const gerencia = mail.gerenciaFirma || "Gerencia de Gestión de Prestadores Interior";
  const nLlamado = pdf.nroTramite || mail.tramiteBody || "";

  const sheet = workbook.getActiveWorksheet();
  const usedRange = sheet.getUsedRange();
  const lastRowIndex = usedRange ? usedRange.getRowCount() : 1;

  // Calcular Número de ingreso (+1 respecto de la fila anterior en Columna A)
  let nextIngreso: number | string = 1;
  if (lastRowIndex >= 2) {
    const prevVal = sheet.getCell(lastRowIndex - 1, 0).getValue();
    const parsedNum = Number(prevVal);
    if (!isNaN(parsedNum) && parsedNum > 0) {
      nextIngreso = parsedNum + 1;
    }
  }

  const newRowValues: (string | number)[][] = [[
    nextIngreso,
    fechaRecepcion,
    pdf.vigenciaDesde,
    pdf.vigenciaHasta,
    pdf.numeroSocio,
    pdf.numeroIntegrante,
    pdf.codigoPrestador,
    pdf.nombrePrestador,
    mail.numeroConvenio,
    pdf.codigoPrestacion,
    pdf.nombrePrestacion,
    valorRef,
    pdf.importePresupuestado,
    pdf.importeExcepcion,
    pdf.importeALiquidar,
    gerencia,
    usuario,
    hoy,
    nLlamado ? Number(nLlamado) : "",
    pdf.tipoExcepcion,
    pdf.motivoExcepcion,
    "",
    ""
  ]];

  const tables = sheet.getTables();
  if (tables.length > 0) {
    tables[0].addRow(-1, newRowValues[0]);
  } else {
    const targetRange = sheet.getRangeByIndexes(lastRowIndex, 0, 1, newRowValues[0].length);
    targetRange.setValues(newRowValues);
  }
}

function decodeBase64ToLatin1(b64: string): string {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  const clean = (b64 || "").replace(/[^A-Za-z0-9+/=]/g, "");
  let output = "";
  let i = 0;
  while (i < clean.length) {
    const enc1 = chars.indexOf(clean.charAt(i++));
    const enc2 = chars.indexOf(clean.charAt(i++));
    const enc3 = chars.indexOf(clean.charAt(i++));
    const enc4 = chars.indexOf(clean.charAt(i++));
    const chr1 = (enc1 << 2) | (enc2 >> 4);
    const chr2 = ((enc2 & 15) << 4) | (enc3 >> 2);
    const chr3 = ((enc3 & 3) << 6) | enc4;
    output += String.fromCharCode(chr1);
    if (enc3 !== 64 && enc3 !== -1) output += String.fromCharCode(chr2);
    if (enc4 !== 64 && enc4 !== -1) output += String.fromCharCode(chr3);
  }
  return output;
}

function extractBlocksFromPdf(rawLatin1: string): string[] {
  const blocks: string[] = [];
  const streamRegex = /stream\\r?\\n([\\s\\S]*?)\\r?\\nendstream/g;
  let sm: RegExpExecArray | null;
  while ((sm = streamRegex.exec(rawLatin1)) !== null) {
    const content = sm[1];
    if (content.indexOf("BT") !== -1 && content.indexOf("Tj") !== -1) {
      const btRegex = /BT\\s+([\\s\\S]*?)\\s+ET/g;
      let bm: RegExpExecArray | null;
      while ((bm = btRegex.exec(content)) !== null) {
        const tjRegex = /\\((.*?)\\)\\s*Tj/g;
        let tm: RegExpExecArray | null;
        const parts: string[] = [];
        while ((tm = tjRegex.exec(bm[1])) !== null) {
          const clean = tm[1].replace(/\\\\\\(/g, "(").replace(/\\\\\\)/g, ")").replace(/\\s+/g, " ").trim();
          if (clean) parts.push(clean);
        }
        if (parts.length > 0) blocks.push(parts.join(" "));
      }
    }
  }
  return blocks;
}

function parseMoney(raw: string): number | string {
  const cleaned = (raw || "").replace(/[^\\d.,-]/g, "").trim();
  if (!cleaned) return "";
  const norm = cleaned.indexOf(",") > cleaned.indexOf(".") && cleaned.indexOf(".") !== -1
    ? cleaned.replace(/\\./g, "").replace(",", ".")
    : cleaned.replace(/,/g, "");
  const num = Number(norm);
  return isNaN(num) ? "" : num;
}

function parseFormularioBlocks(blocks: string[]) {
  const res = {
    nroTramite: "",
    fechaFormulario: "",
    solicitante: "",
    tipoExcepcion: "Al Prestador",
    motivoExcepcion: "",
    codigoPrestador: "" as number | string,
    nombrePrestador: "",
    numeroSocio: "" as number | string,
    numeroIntegrante: "" as number | string,
    codigoPrestacion: "" as number | string,
    nombrePrestacion: "",
    importeExcepcion: "" as number | string,
    importePresupuestado: "" as number | string,
    importeALiquidar: "" as number | string,
    vigenciaDesde: "",
    vigenciaHasta: ""
  };
  let section = "";
  for (let i = 0; i < blocks.length; i++) {
    const tok = (blocks[i] || "").trim();
    const next = i + 1 < blocks.length ? (blocks[i + 1] || "").trim() : "";
    if (tok === "Datos del Prestador" || tok === "Datos del Afiliado" || tok === "Detalle de la excepción a otorgar") {
      section = tok;
      continue;
    }
    if (tok === "Nro Trámite:") { res.nroTramite = next; i++; }
    else if (tok === "Fecha:" && !res.fechaFormulario) { res.fechaFormulario = next; i++; }
    else if (tok === "Solicitante:") { res.solicitante = next; i++; }
    else if (tok === "Descripción:") { res.motivoExcepcion = next; i++; }
    else if (tok === "Código:" && section.indexOf("Prestador") !== -1) { res.codigoPrestador = Number(next) || next; i++; }
    else if (tok === "Apellido y Nombre / Razón Social:") { res.nombrePrestador = next; i++; }
    else if (tok === "Contra/Inte:") {
      const m = next.match(/0*(\\d+)\\s*\\/\\s*0*(\\d+)/);
      if (m) { res.numeroSocio = Number(m[1]); res.numeroIntegrante = Number(m[2]); }
      i++;
    }
    else if (tok === "Prestación:") {
      const m = next.match(/^\\s*(\\d+)\\s*-\\s*(.+)$/);
      if (m) { res.codigoPrestacion = Number(m[1]); res.nombrePrestacion = m[2].trim(); }
      i++;
    }
    else if (tok === "Importe de la excepción:") { res.importeExcepcion = parseMoney(next); i++; }
    else if (tok === "Importe presupuestado:") { res.importePresupuestado = parseMoney(next); i++; }
    else if (tok === "Importe a liquidar:") { res.importeALiquidar = parseMoney(next); i++; }
    else if (tok === "Fecha desde:") { res.vigenciaDesde = next; i++; }
    else if (tok === "Fecha hasta:") { res.vigenciaHasta = next; i++; }
  }
  return res;
}

function parseEmailBody(body: string, receivedDate: string) {
  const cleanText = (body || "")
    .replace(/<\\/(td|th|tr|p|div)>/gi, "\\n")
    .replace(/<br\\s*\\/?>/gi, "\\n")
    .replace(/<[^>]+>/g, " ");
  const tokens = cleanText.split(/\\r?\\n|\\t/).map(s => s.replace(/\\s+/g, " ").trim()).filter(Boolean);

  let numeroConvenio: number | string = "";
  let valorReferencia: number | string = "";
  for (let i = 0; i < tokens.length; i++) {
    if (/Convenio\\s*\\(\\s*1\\s*\\)/i.test(tokens[i])) {
      const m = tokens[i].match(/Convenio\\s*\\(\\s*1\\s*\\)\\s*(\\d+)/i);
      if (m) numeroConvenio = Number(m[1]);
      else if (i + 1 < tokens.length && /^\\d+$/.test(tokens[i + 1])) numeroConvenio = Number(tokens[i + 1]);
    }
    if (/Valor\\s+de\\s+referencia\\s*\\(\\s*2\\s*\\)/i.test(tokens[i])) {
      const rem = tokens[i].replace(/.*Valor\\s+de\\s+referencia\\s*\\(\\s*2\\s*\\)/i, "");
      const cands = [rem, tokens[i + 1] || "", tokens[i + 2] || ""];
      for (const c of cands) {
        const amt = parseMoney(c.split(/Prestador/i)[0]);
        if (amt !== "") { valorReferencia = amt; break; }
      }
    }
  }
  let fechaRecepcion = "";
  if (receivedDate) {
    const dt = new Date(receivedDate);
    if (!isNaN(dt.getTime())) {
      const d = ("0" + dt.getDate()).slice(-2);
      const m = ("0" + (dt.getMonth() + 1)).slice(-2);
      fechaRecepcion = d + "/" + m + "/" + dt.getFullYear();
    }
  }
  return {
    fechaRecepcion,
    numeroConvenio,
    valorReferencia,
    tramiteBody: "",
    gerenciaFirma: "Gerencia de Gestión de Prestadores Interior"
  };
}

function formatTodayDDMMYYYY(): string {
  const dt = new Date();
  const d = ("0" + dt.getDate()).slice(-2);
  const m = ("0" + (dt.getMonth() + 1)).slice(-2);
  return d + "/" + m + "/" + dt.getFullYear();
}`;

  // Estado de la aplicación
  let currentPdfData = {};
  let batchRows = [];
  const previewInputs = {};

  // Elementos del DOM
  const btnModeManual = document.getElementById("btnModeManual");
  const btnModeAuto = document.getElementById("btnModeAuto");
  const autoSetupView = document.getElementById("autoSetupView");
  const manualAssistView = document.getElementById("manualAssistView");

  const btnCopyExcelScript = document.getElementById("btnCopyExcelScript");
  const btnToggleScriptPreview = document.getElementById("btnToggleScriptPreview");
  const excelScriptPre = document.getElementById("excelScriptPre");
  const btnCopyApiUrl = document.getElementById("btnCopyApiUrl");
  const apiEndpointUrl = document.getElementById("apiEndpointUrl");

  const decimalSeparatorSelect = document.getElementById("decimalSeparatorSelect");
  const defaultGerenciaInput = document.getElementById("defaultGerenciaInput");
  const fechaCargaInput = document.getElementById("fechaCargaInput");

  const pdfDropzone = document.getElementById("pdfDropzone");
  const pdfFileInput = document.getElementById("pdfFileInput");
  const pdfStatusTitle = document.getElementById("pdfStatusTitle");
  const pdfStatusSub = document.getElementById("pdfStatusSub");

  const emailBodyInput = document.getElementById("emailBodyInput");
  const manualFechaMailInput = document.getElementById("manualFechaMailInput");
  const manualConvenioInput = document.getElementById("manualConvenioInput");
  const manualValorRefInput = document.getElementById("manualValorRefInput");

  const previewFieldsGrid = document.getElementById("previewFieldsGrid");
  const btnAddRow = document.getElementById("btnAddRow");
  const btnClearInputs = document.getElementById("btnClearInputs");

  const batchTableHeader = document.getElementById("batchTableHeader");
  const batchTableBody = document.getElementById("batchTableBody");
  const rowCountBadge = document.getElementById("rowCountBadge");
  const btnCopyForSharePoint = document.getElementById("btnCopyForSharePoint");
  const btnCopyWithHeaders = document.getElementById("btnCopyWithHeaders");
  const btnDownloadXlsx = document.getElementById("btnDownloadXlsx");
  const btnClearTable = document.getElementById("btnClearTable");
  const toastBanner = document.getElementById("toastBanner");

  function getTodayDDMMYYYY() {
    const now = new Date();
    const d = String(now.getDate()).padStart(2, "0");
    const m = String(now.getMonth() + 1).padStart(2, "0");
    return `${d}/${m}/${now.getFullYear()}`;
  }

  function initConfig() {
    fechaCargaInput.value = getTodayDDMMYYYY();
    const savedSep = localStorage.getItem("swiss_abm_dec_sep");
    if (savedSep === "," || savedSep === ".") {
      decimalSeparatorSelect.value = savedSep;
    }

    decimalSeparatorSelect.addEventListener("change", () => {
      localStorage.setItem("swiss_abm_dec_sep", decimalSeparatorSelect.value);
    });
    defaultGerenciaInput.addEventListener("input", recomputePreview);
    fechaCargaInput.addEventListener("input", recomputePreview);

    if (excelScriptPre) {
      excelScriptPre.textContent = OFFICE_SCRIPT_CODE;
    }
    if (apiEndpointUrl && window.location.origin.startsWith("http")) {
      apiEndpointUrl.textContent = `${window.location.origin}/api/procesar`;
    }
  }

  function buildPreviewGrid() {
    previewFieldsGrid.innerHTML = "";
    COLUMNS.forEach((col) => {
      const meta = COLUMN_SOURCE_META[col] || { tag: "Auto", cls: "source-auto" };
      const wrapper = document.createElement("div");
      wrapper.className = "preview-field" + (meta.wide ? " wide" : "");

      const label = document.createElement("label");
      const nameSpan = document.createElement("span");
      nameSpan.textContent = col;
      const badge = document.createElement("span");
      badge.className = `source-tag ${meta.cls}`;
      badge.textContent = meta.tag;
      label.appendChild(nameSpan);
      label.appendChild(badge);

      const input = document.createElement("input");
      input.type = "text";
      input.dataset.column = col;

      wrapper.appendChild(label);
      wrapper.appendChild(input);
      previewFieldsGrid.appendChild(wrapper);
      previewInputs[col] = input;
    });
  }

  function buildBatchTableHeader() {
    batchTableHeader.innerHTML = "";
    const thAction = document.createElement("th");
    thAction.textContent = "#";
    batchTableHeader.appendChild(thAction);

    COLUMNS.forEach((col) => {
      const th = document.createElement("th");
      th.textContent = col;
      batchTableHeader.appendChild(th);
    });
  }

  async function extractPdfFallbackWithPdfJs(arrayBuffer) {
    if (!window.pdfjsLib) return [];
    try {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      const blocks = [];
      for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
        const page = await pdf.getPage(pageNum);
        const textContent = await page.getTextContent();
        for (const item of textContent.items) {
          const s = (item.str || "").trim();
          if (s) blocks.push(s);
        }
      }
      return blocks;
    } catch (e) {
      return [];
    }
  }

  async function loadPdfBuffer(arrayBuffer, filename) {
    const uint8 = new Uint8Array(arrayBuffer);
    const fallbackBlocks = await extractPdfFallbackWithPdfJs(arrayBuffer.slice(0));
    currentPdfData = window.SwissParser.parseFormularioPdfBytes(uint8, fallbackBlocks);

    pdfDropzone.classList.add("loaded");
    pdfStatusTitle.innerHTML = `PDF cargado: <strong>${filename}</strong>`;
    pdfStatusSub.textContent = `Trámite #${currentPdfData.nro_tramite || "-"} | Prestador: ${
      currentPdfData.codigo_prestador || "-"
    } ${currentPdfData.nombre_prestador || ""} | Socio: ${currentPdfData.numero_socio || "-"}`;

    recomputePreview();
  }

  function recomputePreview() {
    const emailParsed = window.SwissParser.parseEmailBody(emailBodyInput.value, {});

    if (emailParsed.fecha_recepcion && !manualFechaMailInput.dataset.userEdited) {
      manualFechaMailInput.value = emailParsed.fecha_recepcion;
    }
    if (emailParsed.numero_convenio !== "" && !manualConvenioInput.dataset.userEdited) {
      manualConvenioInput.value = emailParsed.numero_convenio;
    }
    if (emailParsed.valor_referencia !== "" && !manualValorRefInput.dataset.userEdited) {
      manualValorRefInput.value = emailParsed.valor_referencia;
    }

    const effectiveEmailData = Object.assign({}, emailParsed, {
      fecha_recepcion:
        manualFechaMailInput.value.trim() || emailParsed.fecha_recepcion || "",
      numero_convenio:
        manualConvenioInput.value.trim() !== ""
          ? manualConvenioInput.value.trim()
          : emailParsed.numero_convenio,
      valor_referencia:
        manualValorRefInput.value.trim() !== ""
          ? window.SwissParser.parseMoneyAmount(manualValorRefInput.value)
          : emailParsed.valor_referencia,
    });

    const record = window.SwissParser.buildExcelRecord(effectiveEmailData, currentPdfData, {
      fechaCarga: fechaCargaInput.value.trim() || getTodayDDMMYYYY(),
      gerenciaAutorizante: defaultGerenciaInput.value.trim(),
    });

    COLUMNS.forEach((col) => {
      if (previewInputs[col]) {
        previewInputs[col].value =
          record.dict[col] !== undefined && record.dict[col] !== null
            ? String(record.dict[col])
            : "";
      }
    });
  }

  function renderBatchTable() {
    rowCountBadge.textContent = String(batchRows.length);
    const hasRows = batchRows.length > 0;
    btnCopyForSharePoint.disabled = !hasRows;
    btnCopyWithHeaders.disabled = !hasRows;
    btnDownloadXlsx.disabled = !hasRows;
    btnClearTable.disabled = !hasRows;

    batchTableBody.innerHTML = "";
    if (!hasRows) {
      const tr = document.createElement("tr");
      tr.className = "empty-row";
      const td = document.createElement("td");
      td.colSpan = COLUMNS.length + 1;
      td.textContent =
        "Todavía no hay filas en el lote. Subí un archivo PDF y pegá el cuerpo del mail en el Paso 1.";
      tr.appendChild(td);
      batchTableBody.appendChild(tr);
      return;
    }

    batchRows.forEach((rowObj, rowIdx) => {
      const tr = document.createElement("tr");

      const tdDel = document.createElement("td");
      const btnDel = document.createElement("button");
      btnDel.className = "btn-delete-row";
      btnDel.title = "Eliminar fila";
      btnDel.textContent = "×";
      btnDel.addEventListener("click", () => {
        batchRows.splice(rowIdx, 1);
        renderBatchTable();
      });
      tdDel.appendChild(btnDel);
      tr.appendChild(tdDel);

      COLUMNS.forEach((col) => {
        const td = document.createElement("td");
        td.textContent = rowObj[col] !== undefined && rowObj[col] !== null ? String(rowObj[col]) : "";
        tr.appendChild(td);
      });

      batchTableBody.appendChild(tr);
    });
  }

  function showToast(message) {
    toastBanner.textContent = message;
    toastBanner.classList.remove("hidden");
    setTimeout(() => {
      toastBanner.classList.add("hidden");
    }, 4500);
  }

  function formatCellForClipboard(value, colName) {
    if (value === null || value === undefined || value === "") return "";
    const str = String(value).trim();
    const moneyCols = [
      "Valor de referencia",
      "Valor presupuestado",
      "Valor de la excepción",
      "Valor a liquidar",
    ];
    if (moneyCols.includes(colName)) {
      const num = Number(str.replace(",", "."));
      if (Number.isFinite(num)) {
        const decSep = decimalSeparatorSelect.value;
        const formatted = num.toFixed(2);
        return decSep === "," ? formatted.replace(".", ",") : formatted;
      }
    }
    return str.replace(/[\t\r\n]+/g, " ");
  }

  async function copyTableToClipboard(includeHeaders) {
    if (batchRows.length === 0) return;
    const lines = [];
    if (includeHeaders) {
      lines.push(COLUMNS.join("\t"));
    }
    for (const rowObj of batchRows) {
      const rowCells = COLUMNS.map((col) => formatCellForClipboard(rowObj[col], col));
      lines.push(rowCells.join("\t"));
    }
    const tsv = lines.join("\r\n");
    await navigator.clipboard.writeText(tsv);
    showToast(
      `¡${batchRows.length} fila(s) copiada(s)! Andá al Excel de SharePoint, parate en la celda 'Fecha de recepción del mail' de la fila nueva y presioná Ctrl + V.`
    );
  }

  function downloadExcelFile() {
    if (batchRows.length === 0 || !window.XLSX) return;
    const aoa = [COLUMNS];
    for (const rowObj of batchRows) {
      aoa.push(
        COLUMNS.map((col) => {
          const val = rowObj[col];
          if (val === "") return "";
          if (/^\d+(\.\d+)?$/.test(String(val))) {
            return Number(val);
          }
          return val;
        })
      );
    }
    const ws = window.XLSX.utils.aoa_to_sheet(aoa);
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, "Excepciones");
    window.XLSX.writeFile(wb, `Excepciones_ABM_${getTodayDDMMYYYY().replace(/\//g, "-")}.xlsx`);
  }

  function clearInputs() {
    currentPdfData = {};
    pdfDropzone.classList.remove("loaded");
    pdfStatusTitle.innerHTML =
      'Arrastrá acá el archivo <strong>.PDF</strong> adjunto ("Formulario de Excepciones")';
    pdfStatusSub.textContent = "O hacé clic para seleccionarlo desde tu computadora";
    pdfFileInput.value = "";
    emailBodyInput.value = "";
    manualFechaMailInput.value = "";
    manualConvenioInput.value = "";
    manualValorRefInput.value = "";
    delete manualFechaMailInput.dataset.userEdited;
    delete manualConvenioInput.dataset.userEdited;
    delete manualValorRefInput.dataset.userEdited;
    recomputePreview();
  }

  // Navegación entre modo Manual Asistido y Guía 100% Automático
  btnModeManual.addEventListener("click", () => {
    btnModeManual.classList.add("active");
    btnModeAuto.classList.remove("active");
    manualAssistView.classList.remove("hidden");
    autoSetupView.classList.add("hidden");
  });

  btnModeAuto.addEventListener("click", () => {
    btnModeAuto.classList.add("active");
    btnModeManual.classList.remove("active");
    autoSetupView.classList.remove("hidden");
    manualAssistView.classList.add("hidden");
  });

  document.querySelectorAll(".method-tab").forEach((tabBtn) => {
    tabBtn.addEventListener("click", () => {
      document.querySelectorAll(".method-tab").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".method-content").forEach((c) => c.classList.add("hidden"));
      tabBtn.classList.add("active");
      const targetId = tabBtn.getAttribute("data-target");
      const targetEl = document.getElementById(targetId);
      if (targetEl) targetEl.classList.remove("hidden");
    });
  });

  if (btnCopyExcelScript) {
    btnCopyExcelScript.addEventListener("click", async () => {
      await navigator.clipboard.writeText(OFFICE_SCRIPT_CODE);
      showToast(
        "¡Código para Excel copiado! Ahora pegalo con Ctrl + V en la pestaña 'Automatizar -> Nuevo script' de tu Excel en SharePoint."
      );
    });
  }

  if (btnToggleScriptPreview) {
    btnToggleScriptPreview.addEventListener("click", () => {
      excelScriptPre.classList.toggle("hidden");
    });
  }

  if (btnCopyApiUrl) {
    btnCopyApiUrl.addEventListener("click", async () => {
      await navigator.clipboard.writeText(apiEndpointUrl.textContent);
      showToast("URL del conector copiada al portapapeles.");
    });
  }

  // Eventos del cargador
  pdfDropzone.addEventListener("click", () => pdfFileInput.click());
  pdfDropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    pdfDropzone.classList.add("dragover");
  });
  pdfDropzone.addEventListener("dragleave", () => {
    pdfDropzone.classList.remove("dragover");
  });
  pdfDropzone.addEventListener("drop", async (e) => {
    e.preventDefault();
    pdfDropzone.classList.remove("dragover");
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) {
      const buf = await file.arrayBuffer();
      await loadPdfBuffer(buf, file.name);
    }
  });

  pdfFileInput.addEventListener("change", async (e) => {
    const file = e.target.files && e.target.files[0];
    if (file) {
      const buf = await file.arrayBuffer();
      await loadPdfBuffer(buf, file.name);
    }
  });

  emailBodyInput.addEventListener("input", () => {
    delete manualFechaMailInput.dataset.userEdited;
    delete manualConvenioInput.dataset.userEdited;
    delete manualValorRefInput.dataset.userEdited;
    recomputePreview();
  });

  manualFechaMailInput.addEventListener("input", () => {
    manualFechaMailInput.dataset.userEdited = "1";
    recomputePreview();
  });
  manualConvenioInput.addEventListener("input", () => {
    manualConvenioInput.dataset.userEdited = "1";
    recomputePreview();
  });
  manualValorRefInput.addEventListener("input", () => {
    manualValorRefInput.dataset.userEdited = "1";
    recomputePreview();
  });

  btnClearInputs.addEventListener("click", clearInputs);

  btnAddRow.addEventListener("click", () => {
    const rowObj = {};
    COLUMNS.forEach((col) => {
      rowObj[col] = previewInputs[col] ? previewInputs[col].value.trim() : "";
    });
    batchRows.push(rowObj);
    renderBatchTable();

    showToast(
      `Fila agregada a la tabla de carga (Trámite #${rowObj["N_de_llamado"] || "-"}).`
    );
    clearInputs();
  });

  btnCopyForSharePoint.addEventListener("click", () => copyTableToClipboard(false));
  btnCopyWithHeaders.addEventListener("click", () => copyTableToClipboard(true));
  btnDownloadXlsx.addEventListener("click", downloadExcelFile);
  btnClearTable.addEventListener("click", () => {
    batchRows = [];
    renderBatchTable();
  });

  // Inicializar
  buildPreviewGrid();
  buildBatchTableHeader();
  initConfig();
  recomputePreview();
})();
