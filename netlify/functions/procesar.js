/**
 * Netlify Serverless Function: /.netlify/functions/procesar (o /api/procesar)
 *
 * Permite conectar Power Automate (Microsoft 365) de forma 100% automática:
 * 1. Power Automate detecta el mail nuevo en ABM.Convenios@swissmedical.com.ar
 * 2. Envía un POST JSON con { emailBody, receivedDate, subject, senderName, senderEmail, pdfBase64, numeroIngreso }
 * 3. Esta función devuelve las 23 columnas exactas listas para insertar la fila en el Excel de SharePoint.
 */

const SwissParser = require("../../public/parser.js");

exports.handler = async function (event) {
  const headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Content-Type": "application/json; charset=utf-8",
  };

  if (event.httpMethod === "OPTIONS") {
    return { statusCode: 200, headers, body: "" };
  }

  if (event.httpMethod !== "POST") {
    return {
      statusCode: 405,
      headers,
      body: JSON.stringify({
        ok: false,
        error: "Usar método POST con JSON { emailBody, receivedDate, pdfBase64, numeroIngreso }",
      }),
    };
  }

  try {
    const payload = JSON.parse(event.body || "{}");

    const emailData = SwissParser.parseEmailBody(payload.emailBody || "", {
      subject: payload.subject || "",
      senderName: payload.senderName || "",
      senderEmail: payload.senderEmail || "",
      receivedDate: payload.receivedDate || "",
    });

    let pdfData = {};
    if (payload.pdfBase64) {
      const cleanB64 = String(payload.pdfBase64).replace(/^data:application\/pdf;base64,/, "");
      const pdfBytes = new Uint8Array(Buffer.from(cleanB64, "base64"));
      pdfData = SwissParser.parseFormularioPdfBytes(pdfBytes);
    }

    const record = SwissParser.buildExcelRecord(emailData, pdfData, {
      numeroIngreso: payload.numeroIngreso || 50608,
      fechaCarga: payload.fechaCarga || "",
      gerenciaAutorizante: payload.gerenciaAutorizante || "",
      usuarioCarga: payload.usuarioCarga || "",
    });

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({
        ok: true,
        rowObject: record.dict,
        rowValues: record.values,
        emailExtracted: emailData,
        pdfExtracted: pdfData,
      }),
    };
  } catch (err) {
    return {
      statusCode: 400,
      headers,
      body: JSON.stringify({
        ok: false,
        error: err.message || String(err),
      }),
    };
  }
};
