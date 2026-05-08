import http from "node:http";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const publicDir = path.join(__dirname, "public");
const uploadDir = path.join(__dirname, "uploads");
const port = Number(process.env.PORT || 4173);

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
};

function sendJson(res, status, body) {
  const payload = JSON.stringify(body);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(payload),
  });
  res.end(payload);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => chunks.push(chunk));
    req.on("error", reject);
    req.on("end", () => resolve(Buffer.concat(chunks)));
  });
}

function parseNdjson(ndjson) {
  return ndjson
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => JSON.parse(line));
}

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function asNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return 0;
  const cleaned = value
    .replace(/[^\d,.-]/g, "")
    .replace(/\.(?=\d{3}(\D|$))/g, "")
    .replace(",", ".");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function findLabelValue(rows, label) {
  const target = normalizeText(label);
  for (const row of rows) {
    const index = row.findIndex((cell) => normalizeText(cell) === target);
    if (index === -1) continue;
    for (let col = index + 1; col < row.length; col += 1) {
      const value = asNumber(row[col]);
      if (value !== 0) return value;
    }
  }
  return 0;
}

function findTextAfterLabel(rows, label) {
  const target = normalizeText(label);
  for (const row of rows) {
    const index = row.findIndex((cell) => normalizeText(cell) === target);
    if (index === -1) continue;
    for (let col = index + 1; col < row.length; col += 1) {
      const value = row[col];
      if (value !== null && value !== undefined && String(value).trim()) {
        return String(value).trim();
      }
    }
  }
  return "";
}

function fixedColumn(row, index) {
  return asNumber(row[index]);
}

function getPaymentColumns(row) {
  return {
    contaCliente: fixedColumn(row, 10),
    dinheiro: fixedColumn(row, 11),
    cheques: fixedColumn(row, 13),
    pre: fixedColumn(row, 15),
    cartaoDebito: fixedColumn(row, 16),
    cartaoCredito: fixedColumn(row, 17),
    vales: fixedColumn(row, 19),
    parcelado: fixedColumn(row, 21) || fixedColumn(row, 22),
    outros: fixedColumn(row, 24),
  };
}

function paymentLabels(payments) {
  return [
    ["Dinheiro", payments.dinheiro],
    ["Cheques/Pix", payments.cheques],
    ["Pre", payments.pre],
    ["Debito", payments.cartaoDebito],
    ["Credito", payments.cartaoCredito],
    ["Vales", payments.vales],
    ["Parcelado", payments.parcelado],
    ["Outros", payments.outros],
    ["Conta cliente", payments.contaCliente],
  ]
    .filter(([, value]) => Math.abs(value) > 0.009)
    .map(([label]) => label);
}

function extractSequence(description) {
  const text = normalizeText(description);
  const match = text.match(/(?:saida|entrada|sequencia)\s+(\d+)/);
  return match?.[1] ?? "";
}

function amountForCashTotal(payments) {
  return (
    payments.contaCliente +
    payments.dinheiro +
    payments.cheques +
    payments.pre +
    payments.cartaoDebito +
    payments.cartaoCredito +
    payments.outros
  );
}

function parseTransactions(rows, totalRowIndex) {
  const transactions = [];

  for (let index = 0; index < totalRowIndex; index += 1) {
    const row = rows[index] ?? [];
    const description = String(row[1] ?? "").trim();
    const type = String(row[7] ?? "").trim().toUpperCase();
    if (!description || !type || normalizeText(description) === "descricao") continue;
    if (normalizeText(description).startsWith("tipo :")) continue;

    const payments = getPaymentColumns(row);
    const cashAmount = amountForCashTotal(payments);
    const visibleAmount = cashAmount;
    const sequence = extractSequence(description);
    if (!sequence && Math.abs(visibleAmount) < 0.009) continue;

    const isReturn = ["DEV", "DES", "ADS"].includes(type) || visibleAmount < 0;
    const details = [];
    for (let next = index + 1; next < Math.min(index + 4, totalRowIndex); next += 1) {
      const nextRow = rows[next] ?? [];
      const nextType = String(nextRow[7] ?? "").trim();
      if (nextType) break;
      const text = String(nextRow[1] ?? "").trim();
      if (text) details.push(text);
    }

    transactions.push({
      id: `${index + 1}-${sequence || type}`,
      row: index + 1,
      sequence: sequence || `${type}-${index + 1}`,
      description,
      details,
      employee: String(row[5] ?? "").trim(),
      type,
      payments,
      paymentMethods: paymentLabels(payments),
      amount: visibleAmount,
      cashAmount,
      isReturn,
      neutralized: false,
      neutralizedWith: "",
      needsPhysicalCheck: Math.abs(visibleAmount) > 0.009,
    });
  }

  markNeutralizedPairs(transactions);
  return transactions;
}

function sameMoneyAmount(left, right) {
  return Math.abs(Math.abs(left) - Math.abs(right)) < 0.01;
}

function canNeutralize(left, right) {
  if (left.neutralized || right.neutralized) return false;
  if (left.cashAmount === 0 || right.cashAmount === 0) return false;
  if (Math.sign(left.cashAmount) === Math.sign(right.cashAmount)) return false;
  if (!sameMoneyAmount(left.cashAmount, right.cashAmount)) return false;
  if (left.employee && right.employee && normalizeText(left.employee) !== normalizeText(right.employee)) {
    return false;
  }
  return true;
}

function markNeutralizedPairs(transactions) {
  for (let index = 0; index < transactions.length; index += 1) {
    const item = transactions[index];
    if (item.neutralized || item.cashAmount >= 0) continue;

    const candidates = [];
    for (let other = Math.max(0, index - 3); other <= Math.min(transactions.length - 1, index + 3); other += 1) {
      if (other !== index) candidates.push(transactions[other]);
    }

    const pair = candidates.find((candidate) => canNeutralize(item, candidate));
    if (!pair) continue;

    item.neutralized = true;
    item.neutralizedWith = pair.sequence;
    item.needsPhysicalCheck = false;

    pair.neutralized = true;
    pair.neutralizedWith = item.sequence;
    pair.needsPhysicalCheck = false;
  }
}

function parseSystemReport(rows) {
  const totalRowIndex = rows.findIndex((row) =>
    row.some((cell) => normalizeText(cell) === "total :"),
  );
  if (totalRowIndex === -1) {
    const looksLikeTicket = rows.some((row) =>
      row.some((cell) => ["pix (cheque)", "cartoes", "cheques(pre)"].includes(normalizeText(cell))),
    );
    if (looksLikeTicket) {
      throw new Error(
        "Este parece ser o arquivo Ticket 3. Importe aqui o relatorio CX exportado do sistema.",
      );
    }
    throw new Error(
      "Nao encontrei os totais do relatorio CX. Confira se o arquivo importado e o relatorio Caixa - Normal do sistema.",
    );
  }

  const totalRow = rows[totalRowIndex];
  const totals = {
    contaCliente: fixedColumn(totalRow, 10),
    dinheiro: fixedColumn(totalRow, 11),
    cheques: fixedColumn(totalRow, 13),
    pre: fixedColumn(totalRow, 15),
    cartaoDebito: fixedColumn(totalRow, 16),
    cartaoCredito: fixedColumn(totalRow, 17),
    vales: fixedColumn(totalRow, 19),
    parcelado: fixedColumn(totalRow, 21),
    outros: fixedColumn(totalRow, 24),
    totalCaixa: findLabelValue(rows, "Total deste caixa :"),
    totalRegistros: findLabelValue(rows, "Total de Registros :"),
  };

  return {
    totals,
    transactions: parseTransactions(rows, totalRowIndex),
    metadata: {
      empresa: findTextAfterLabel(rows, "Empresa :"),
      dataSerial: findLabelValue(rows, "Data :"),
      caixa: findTextAfterLabel(rows, "Caixa :"),
      linhasLidas: rows.length,
    },
  };
}

async function parseWorkbook(filePath) {
  const input = await FileBlob.load(filePath);
  const workbook = await SpreadsheetFile.importXlsx(input);
  const sheets = parseNdjson(
    (await workbook.inspect({ kind: "sheet", include: "name", maxChars: 20000 })).ndjson,
  ).filter((entry) => entry.kind === "sheet");
  const sheetName = sheets[0]?.name ?? "Sheet1";

  const table = await workbook.inspect({
    kind: "table",
    range: `${sheetName}!A1:Y1000`,
    include: "values",
    maxChars: 400000,
    tableMaxRows: 1000,
    tableMaxCols: 25,
    tableMaxCellChars: 180,
  });
  const tableRecord = parseNdjson(table.ndjson).find((entry) => entry.kind === "table");
  if (!tableRecord?.values) {
    throw new Error("Nao consegui ler os dados do arquivo CX.");
  }
  return parseSystemReport(tableRecord.values);
}

async function serveStatic(req, res) {
  const url = new URL(req.url, `http://${req.headers.host}`);
  const requestedPath = url.pathname === "/" ? "/index.html" : url.pathname;
  const filePath = path.normalize(path.join(publicDir, requestedPath));

  if (!filePath.startsWith(publicDir)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  try {
    const body = await fs.readFile(filePath);
    const ext = path.extname(filePath);
    res.writeHead(200, { "Content-Type": mimeTypes[ext] ?? "application/octet-stream" });
    res.end(body);
  } catch {
    res.writeHead(404);
    res.end("Not found");
  }
}

const server = http.createServer(async (req, res) => {
  try {
    if (req.method === "POST" && req.url === "/api/parse-cx") {
      await fs.mkdir(uploadDir, { recursive: true });
      const body = await readBody(req);
      if (!body.length) {
        sendJson(res, 400, { error: "Envie um arquivo .xlsx do CX." });
        return;
      }
      const filePath = path.join(uploadDir, `cx-${Date.now()}.xlsx`);
      await fs.writeFile(filePath, body);
      const parsed = await parseWorkbook(filePath);
      sendJson(res, 200, parsed);
      return;
    }

    await serveStatic(req, res);
  } catch (error) {
    sendJson(res, 500, { error: error.message || "Erro inesperado." });
  }
});

server.listen(port, () => {
  console.log(`Conferencia de caixa em http://localhost:${port}`);
});
