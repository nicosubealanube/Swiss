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

  // Estado de la aplicación
  let currentPdfData = {};
  let batchRows = [];
  const previewInputs = {};

  // Elementos del DOM
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

  const btnToggleAutomate = document.getElementById("btnToggleAutomate");
  const btnCloseAutomate = document.getElementById("btnCloseAutomate");
  const automatePanel = document.getElementById("automatePanel");

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

  // Eventos
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

  btnToggleAutomate.addEventListener("click", () => {
    automatePanel.classList.toggle("hidden");
  });
  btnCloseAutomate.addEventListener("click", () => {
    automatePanel.classList.add("hidden");
  });

  // Inicializar
  buildPreviewGrid();
  buildBatchTableHeader();
  initConfig();
  recomputePreview();
})();
