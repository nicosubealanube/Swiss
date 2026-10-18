/**
 * Motor compartido de extracción (Navegador + Netlify Serverless Function)
 * Procesa el Cuerpo del Email + Formulario de Excepciones (.pdf) -> 22 Columnas de Excel en SharePoint
 * (La columna 'Número de ingreso' se autonumera sola en el Excel de SharePoint).
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.SwissParser = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const EXCEL_COLUMNS = [
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
  ];

  const MESES_ES = {
    enero: 1,
    febrero: 2,
    marzo: 3,
    abril: 4,
    mayo: 5,
    junio: 6,
    julio: 7,
    agosto: 8,
    septiembre: 9,
    setiembre: 9,
    octubre: 10,
    noviembre: 11,
    diciembre: 12,
  };

  const SOLICITANTES_MAP = {
    "maria elena velez": {
      gerencia: "Gerencia de Gestión de Prestadores Interior",
      usuario: "MV",
    },
    "maria.velez@swissmedical.com.ar": {
      gerencia: "Gerencia de Gestión de Prestadores Interior",
      usuario: "MV",
    },
  };

  function unescapePdfString(text) {
    return text
      .replace(/\\\(/g, "(")
      .replace(/\\\)/g, ")")
      .replace(/\\\\/g, "\\")
      .replace(/&aacute;/gi, "á")
      .replace(/&eacute;/gi, "é")
      .replace(/&iacute;/gi, "í")
      .replace(/&oacute;/gi, "ó")
      .replace(/&uacute;/gi, "ú")
      .replace(/&ntilde;/gi, "ñ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function uint8ArrayToLatin1(bytes) {
    let str = "";
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const sub = bytes.subarray(i, i + chunkSize);
      str += String.fromCharCode.apply(null, sub);
    }
    return str;
  }

  function extractBlocksFromPdfBytes(pdfBytes, fallbackBlocks) {
    const rawLatin1 = uint8ArrayToLatin1(pdfBytes);
    const blocks = [];
    const streamRegex = /stream\r?\n([\s\S]*?)\r?\nendstream/g;
    let streamMatch;

    while ((streamMatch = streamRegex.exec(rawLatin1)) !== null) {
      const streamContent = streamMatch[1];
      if (streamContent.indexOf("BT") !== -1 && streamContent.indexOf("Tj") !== -1) {
        const btRegex = /BT\s+([\s\S]*?)\s+ET/g;
        let btMatch;
        while ((btMatch = btRegex.exec(streamContent)) !== null) {
          const btBlock = btMatch[1];
          const tjRegex = /\((.*?)\)\s*Tj/g;
          let tjMatch;
          const parts = [];
          while ((tjMatch = tjRegex.exec(btBlock)) !== null) {
            const cleaned = unescapePdfString(tjMatch[1]);
            if (cleaned) parts.push(cleaned);
          }
          if (parts.length > 0) {
            blocks.push(parts.join(" "));
          }
        }
      }
    }

    if (blocks.length > 0) {
      return blocks;
    }
    return Array.isArray(fallbackBlocks) ? fallbackBlocks : [];
  }

  function parseMoneyAmount(raw) {
    if (raw === null || raw === undefined) return "";
    const str = String(raw).trim();
    if (!str) return "";
    let cleaned = str.replace(/[^\d.,-]/g, "").trim();
    if (!cleaned || cleaned === "-") return "";

    if (cleaned.includes(",") && cleaned.includes(".")) {
      if (cleaned.lastIndexOf(",") > cleaned.lastIndexOf(".")) {
        cleaned = cleaned.replace(/\./g, "").replace(",", ".");
      } else {
        cleaned = cleaned.replace(/,/g, "");
      }
    } else if (cleaned.includes(",")) {
      const parts = cleaned.split(",");
      if (parts[parts.length - 1].length === 2) {
        cleaned = parts.slice(0, -1).join("") + "." + parts[parts.length - 1];
      } else {
        cleaned = cleaned.replace(/,/g, "");
      }
    }

    const num = Number(cleaned);
    return Number.isFinite(num) ? num : "";
  }

  function parseSocioIntegrante(raw) {
    const m = String(raw || "").match(/0*(\d+)\s*\/\s*0*(\d+)/);
    if (!m) return { socio: "", integrante: "" };
    return {
      socio: parseInt(m[1], 10),
      integrante: parseInt(m[2], 10),
    };
  }

  function parsePrestacion(raw) {
    const str = String(raw || "").trim();
    const m = str.match(/^\s*(\d+)\s*-\s*(.+)$/);
    if (m) {
      return {
        codigo: parseInt(m[1], 10),
        nombre: m[2].trim(),
      };
    }
    return { codigo: "", nombre: str };
  }

  function parseFormularioPdfBlocks(blocks) {
    const result = {
      nro_tramite: "",
      fecha_formulario: "",
      solicitante: "",
      tipo_excepcion: "",
      motivo_excepcion: "",
      razon: "",
      codigo_prestador: "",
      nombre_prestador: "",
      numero_socio: "",
      numero_integrante: "",
      afiliado_nombre: "",
      codigo_prestacion: "",
      nombre_prestacion: "",
      importe_excepcion: "",
      importe_presupuestado: "",
      importe_a_liquidar: "",
      cantidad: "",
      vigencia_desde: "",
      vigencia_hasta: "",
    };

    const fullText = blocks.join("\n");
    if (/Excepci[oó]n\s+al\s+Pres?t?ador/i.test(fullText)) {
      result.tipo_excepcion = "Al Prestador";
    } else if (/Excepci[oó]n\s+al\s+(Cliente|Afiliado|Socio)/i.test(fullText)) {
      result.tipo_excepcion = "Al Cliente";
    }

    let currentSection = "";
    let i = 0;
    while (i < blocks.length) {
      const token = (blocks[i] || "").trim();
      const nextVal = i + 1 < blocks.length ? (blocks[i + 1] || "").trim() : "";

      if (
        token === "Datos del Prestador" ||
        token === "Datos del Afiliado" ||
        token === "Detalle de la excepción a otorgar" ||
        token.startsWith("Excepción al")
      ) {
        currentSection = token;
        i += 1;
        continue;
      }

      if (token === "Nro Trámite:") {
        result.nro_tramite = nextVal;
        i += 2;
        continue;
      }
      if (token === "Fecha:" && !result.fecha_formulario) {
        result.fecha_formulario = nextVal;
        i += 2;
        continue;
      }
      if (token === "Solicitante:") {
        result.solicitante = nextVal;
        i += 2;
        continue;
      }
      if (token === "Descripción:") {
        result.motivo_excepcion = nextVal;
        i += 2;
        continue;
      }
      if (token === "Razón:") {
        result.razon = nextVal;
        i += 2;
        continue;
      }
      if (token === "Código:" && currentSection.includes("Prestador")) {
        if (/^\d+$/.test(nextVal)) {
          result.codigo_prestador = parseInt(nextVal, 10);
        }
        i += 2;
        continue;
      }
      if (token === "Apellido y Nombre / Razón Social:") {
        result.nombre_prestador = nextVal;
        i += 2;
        continue;
      }
      if (token === "Contra/Inte:") {
        const parsed = parseSocioIntegrante(nextVal);
        result.numero_socio = parsed.socio;
        result.numero_integrante = parsed.integrante;
        i += 2;
        continue;
      }
      if (token === "Apellido y Nombre:" && currentSection.includes("Afiliado")) {
        result.afiliado_nombre = nextVal;
        i += 2;
        continue;
      }
      if (token === "Prestación:") {
        const prest = parsePrestacion(nextVal);
        result.codigo_prestacion = prest.codigo;
        result.nombre_prestacion = prest.nombre;
        i += 2;
        continue;
      }
      if (token === "Importe de la excepción:") {
        result.importe_excepcion = parseMoneyAmount(nextVal);
        i += 2;
        continue;
      }
      if (token === "Importe presupuestado:") {
        result.importe_presupuestado = parseMoneyAmount(nextVal);
        i += 2;
        continue;
      }
      if (token === "Importe a liquidar:") {
        result.importe_a_liquidar = parseMoneyAmount(nextVal);
        i += 2;
        continue;
      }
      if (token === "Cantidad:") {
        if (/^\d+$/.test(nextVal)) {
          result.cantidad = parseInt(nextVal, 10);
        }
        i += 2;
        continue;
      }
      if (token === "Fecha desde:") {
        result.vigencia_desde = nextVal;
        i += 2;
        continue;
      }
      if (token === "Fecha hasta:") {
        result.vigencia_hasta = nextVal;
        i += 2;
        continue;
      }
      i += 1;
    }

    return result;
  }

  function parseFormularioPdfBytes(pdfBytes, fallbackBlocks) {
    const blocks = extractBlocksFromPdfBytes(pdfBytes, fallbackBlocks);
    return parseFormularioPdfBlocks(blocks);
  }

  function normalizeDateDDMMYYYY(rawDate) {
    if (!rawDate) return "";
    const str = String(rawDate).trim();

    const mDirect = str.match(/\b(\d{1,2})\/(\d{1,2})\/(\d{4})\b/);
    if (mDirect) {
      const d = String(parseInt(mDirect[1], 10)).padStart(2, "0");
      const m = String(parseInt(mDirect[2], 10)).padStart(2, "0");
      return `${d}/${m}/${mDirect[3]}`;
    }

    const lower = str.toLowerCase();
    const mEs1 = lower.match(/([a-záéíóú]+)\s+(\d{1,2}),?\s+(\d{4})/);
    if (mEs1 && MESES_ES[mEs1[1]]) {
      const d = String(parseInt(mEs1[2], 10)).padStart(2, "0");
      const m = String(MESES_ES[mEs1[1]]).padStart(2, "0");
      return `${d}/${m}/${mEs1[3]}`;
    }

    const mEs2 = lower.match(/(\d{1,2})\s+de\s+([a-záéíóú]+)\s+de\s+(\d{4})/);
    if (mEs2 && MESES_ES[mEs2[2]]) {
      const d = String(parseInt(mEs2[1], 10)).padStart(2, "0");
      const m = String(MESES_ES[mEs2[2]]).padStart(2, "0");
      return `${d}/${m}/${mEs2[3]}`;
    }

    const parsedDate = new Date(str);
    if (!isNaN(parsedDate.getTime())) {
      const d = String(parsedDate.getDate()).padStart(2, "0");
      const m = String(parsedDate.getMonth() + 1).padStart(2, "0");
      const y = parsedDate.getFullYear();
      return `${d}/${m}/${y}`;
    }

    return "";
  }

  function htmlOrTextToTokens(content) {
    if (!content) return [];
    let text = String(content);
    if (text.includes("<") && text.includes(">")) {
      text = text
        .replace(/<\/(td|th|tr|p|div|li)>/gi, "\n")
        .replace(/<br\s*\/?>/gi, "\n")
        .replace(/<[^>]+>/g, " ");
    }
    return text
      .split(/\r?\n|\t/)
      .map((line) => line.replace(/\s+/g, " ").trim())
      .filter(Boolean);
  }

  function parseEmailBody(bodyContent, options) {
    const opts = options || {};
    const tokens = htmlOrTextToTokens(bodyContent);
    const fullText = tokens.join("\n");

    const result = {
      subject: opts.subject || "",
      sender_name: opts.senderName || "",
      sender_email: opts.senderEmail || "",
      fecha_recepcion: normalizeDateDDMMYYYY(opts.receivedDate || ""),
      numero_convenio: "",
      valor_referencia: "",
      tramite_excepcion_body: "",
      gerencia_firma: "",
    };

    if (!result.fecha_recepcion) {
      const mEnviado = fullText.match(/Enviado:\s*([^\n\r]+)/i);
      if (mEnviado) {
        result.fecha_recepcion = normalizeDateDDMMYYYY(mEnviado[1]);
      }
    }

    // 1. Convenio (1)
    for (let idx = 0; idx < tokens.length; idx++) {
      const tok = tokens[idx];
      if (/Convenio\s*\(\s*1\s*\)/i.test(tok)) {
        const mSame = tok.match(/Convenio\s*\(\s*1\s*\)\s*[:\-]?\s*(\d+)/i);
        if (mSame) {
          result.numero_convenio = parseInt(mSame[1], 10);
          break;
        }
        if (idx + 1 < tokens.length && /^\d+$/.test(tokens[idx + 1])) {
          result.numero_convenio = parseInt(tokens[idx + 1], 10);
          break;
        }
      }
    }

    // 2. Valor de referencia (2)
    for (let idx = 0; idx < tokens.length; idx++) {
      const tok = tokens[idx];
      if (/Valor\s+de\s+referencia\s*\(\s*2\s*\)/i.test(tok)) {
        const remainder = tok.replace(/.*Valor\s+de\s+referencia\s*\(\s*2\s*\)/i, "").trim();
        const candidates = [remainder, tokens[idx + 1] || "", tokens[idx + 2] || ""];
        for (const cand of candidates) {
          if (!cand) continue;
          const candCut = cand.split(/Prestador\s*N/i)[0];
          const amount = parseMoneyAmount(candCut);
          if (amount !== "") {
            result.valor_referencia = amount;
            break;
          }
        }
        if (result.valor_referencia !== "") break;
      }
    }

    // 3. Trámite Excepción #
    const mTramite = fullText.match(/Tr[aá]mite\s+Excepci[oó]n\s*#\s*(\d+)/i);
    if (mTramite) {
      result.tramite_excepcion_body = mTramite[1];
    }

    // 4. Gerencia en firma
    if (/Prestadores\s+Interior/i.test(fullText)) {
      result.gerencia_firma = "Gerencia de Gestión de Prestadores Interior";
    } else if (/Auditor[ií]a\s+M[eé]dica/i.test(fullText)) {
      result.gerencia_firma = "Auditoria Médica";
    }

    return result;
  }

  function resolveGerenciaAndUsuario(emailData, pdfData, defaults) {
    const defs = defaults || {};
    const solicitanteKey = String(pdfData.solicitante || emailData.sender_name || "")
      .trim()
      .toLowerCase();
    const emailKey = String(emailData.sender_email || "")
      .trim()
      .toLowerCase();

    if (SOLICITANTES_MAP[solicitanteKey]) {
      return SOLICITANTES_MAP[solicitanteKey];
    }
    if (SOLICITANTES_MAP[emailKey]) {
      return SOLICITANTES_MAP[emailKey];
    }

    const gerencia =
      defs.gerenciaAutorizante ||
      emailData.gerencia_firma ||
      "Gerencia de Gestión de Prestadores Interior";

    let usuario = defs.usuarioCarga || "";
    if (!usuario && solicitanteKey) {
      const parts = solicitanteKey.split(/\s+/).filter((w) => /^[a-záéíóúñ]+$/i.test(w));
      if (parts.length >= 2) {
        usuario = (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
      } else if (parts.length === 1) {
        usuario = parts[0].slice(0, 2).toUpperCase();
      }
    }

    return { gerencia, usuario };
  }

  function formatTodayDDMMYYYY() {
    const now = new Date();
    const d = String(now.getDate()).padStart(2, "0");
    const m = String(now.getMonth() + 1).padStart(2, "0");
    const y = now.getFullYear();
    return `${d}/${m}/${y}`;
  }

  function buildExcelRecord(emailData, pdfData, options) {
    const opts = options || {};
    const { gerencia, usuario } = resolveGerenciaAndUsuario(emailData, pdfData, opts);
    const fechaCarga = opts.fechaCarga || formatTodayDDMMYYYY();

    let fechaRecepcion = emailData.fecha_recepcion || "";
    if (!fechaRecepcion && pdfData.fecha_formulario) {
      fechaRecepcion = pdfData.fecha_formulario.split(" ")[0];
    }

    let valorRef = emailData.valor_referencia;
    if (
      (valorRef === "" || valorRef === null || valorRef === undefined) &&
      typeof pdfData.importe_presupuestado === "number" &&
      typeof pdfData.importe_excepcion === "number"
    ) {
      valorRef = Number((pdfData.importe_presupuestado - pdfData.importe_excepcion).toFixed(2));
    }

    const nLlamadoRaw = pdfData.nro_tramite || emailData.tramite_excepcion_body || "";
    const nLlamado = /^\d+$/.test(String(nLlamadoRaw)) ? parseInt(nLlamadoRaw, 10) : nLlamadoRaw;

    const recordObj = {
      "Fecha de recepción del mail": fechaRecepcion,
      "Vigencia desde": pdfData.vigencia_desde || "",
      "Vigencia hasta": pdfData.vigencia_hasta || "",
      "Numero_de_Socio": pdfData.numero_socio !== undefined ? pdfData.numero_socio : "",
      "Numero del socio_integrante":
        pdfData.numero_integrante !== undefined ? pdfData.numero_integrante : "",
      "Código del prestador":
        pdfData.codigo_prestador !== undefined ? pdfData.codigo_prestador : "",
      "Nombre del prestador": pdfData.nombre_prestador || "",
      "Número de convenio":
        emailData.numero_convenio !== undefined && emailData.numero_convenio !== null
          ? emailData.numero_convenio
          : "",
      "Código de la Prestación":
        pdfData.codigo_prestacion !== undefined ? pdfData.codigo_prestacion : "",
      "Nombre de la Prestación": pdfData.nombre_prestacion || "",
      "Valor de referencia": valorRef !== undefined && valorRef !== null ? valorRef : "",
      "Valor presupuestado":
        pdfData.importe_presupuestado !== undefined ? pdfData.importe_presupuestado : "",
      "Valor de la excepción":
        pdfData.importe_excepcion !== undefined ? pdfData.importe_excepcion : "",
      "Valor a liquidar":
        pdfData.importe_a_liquidar !== undefined ? pdfData.importe_a_liquidar : "",
      "Gerencia autorizante": gerencia,
      "Usuario de carga": usuario,
      "Fecha de la carga": fechaCarga,
      "N_de_llamado": nLlamado,
      "Tipo_de_excepcion": pdfData.tipo_excepcion || "",
      "Motivo de la Excepción": pdfData.motivo_excepcion || "",
      "Autorizaciones si/no?": "",
      "Detalle": "",
    };

    return {
      columns: EXCEL_COLUMNS,
      dict: recordObj,
      values: EXCEL_COLUMNS.map((col) => recordObj[col]),
    };
  }

  return {
    EXCEL_COLUMNS,
    parseFormularioPdfBytes,
    parseFormularioPdfBlocks,
    parseEmailBody,
    buildExcelRecord,
    parseMoneyAmount,
    normalizeDateDDMMYYYY,
  };
});
