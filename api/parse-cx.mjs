import XLSX from "xlsx";
import { parseSystemReport } from "../lib/parse-cx-core.mjs";

function sendJson(res, status, body) {
  res.statusCode = status;
  res.setHeader("Content-Type", "application/json; charset=utf-8");
  res.end(JSON.stringify(body));
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("error", reject);
    req.on("end", () => resolve(Buffer.concat(chunks)));
  });
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    sendJson(res, 405, { error: "Metodo nao permitido." });
    return;
  }

  try {
    const body = await readBody(req);
    if (!body.length) {
      sendJson(res, 400, { error: "Envie um arquivo .xlsx do CX." });
      return;
    }

    const workbook = XLSX.read(body, {
      type: "buffer",
      cellDates: false,
      cellText: false,
      raw: true,
    });
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
      sendJson(res, 400, { error: "O arquivo enviado nao possui abas." });
      return;
    }

    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {
      header: 1,
      blankrows: true,
      defval: null,
      raw: true,
    });
    sendJson(res, 200, parseSystemReport(rows));
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Erro inesperado." });
  }
}
