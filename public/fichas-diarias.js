const SERVICE_URL = "http://127.0.0.1:8765";
const DRAFT_KEY = "noxxus.daily-sheets.draft";

const money = new Intl.NumberFormat("pt-BR", {
  style: "currency",
  currency: "BRL",
});

const mainHeaders = [
  "DATA",
  "ECF",
  "ACERTO",
  "DINHEIRO",
  "CH AV",
  "CARTÃO",
  "VALES",
  "TOTAL1",
  "PRÉ",
  "VENDAS",
  "RECEB",
  "TOTAL2",
  "DESCON",
  "JUROS",
  "TOTAL3",
  "CH P/ EMI",
];

const valeHeaders = ["DATA", "DISCRIMINAÇÃO", "VALOR", "OBS DIA"];

let serviceInfo = { online: false, exists: false, path: "" };
let existingRows = { fichas: [], vales: [] };
let draft = loadDraft();

function loadDraft() {
  try {
    return JSON.parse(localStorage.getItem(DRAFT_KEY) || '{"fichas":[],"vales":[]}');
  } catch {
    return { fichas: [], vales: [] };
  }
}

function saveDraft() {
  localStorage.setItem(DRAFT_KEY, JSON.stringify(draft));
}

function parseAmount(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const cleaned = String(value || "")
    .replace(/[^\d,.-]/g, "")
    .replace(/\.(?=\d{3}(\D|$))/g, "")
    .replace(",", ".");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function rounded(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function displayDate(isoDate) {
  if (!isoDate) return "";
  const [year, month, day] = isoDate.split("-");
  return `${day}/${month}/${year}`;
}

function excelSerialToIso(value) {
  const date = new Date(Math.round((value - 25569) * 86400 * 1000));
  return Number.isNaN(date.getTime()) ? "" : date.toISOString().slice(0, 10);
}

function toIsoDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === "number" && value > 1000) return excelSerialToIso(value);
  const text = String(value || "").trim();
  if (!text) return "";
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return text;
  const br = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (!br) return "";
  const year = br[3].length === 2 ? `20${br[3]}` : br[3];
  return `${year}-${br[2].padStart(2, "0")}-${br[1].padStart(2, "0")}`;
}

function calculateTotals(row) {
  const total1 = rounded(row.dinheiro + row.chAv + row.cartao + row.vales);
  const total2 = rounded(total1 + row.pre + row.vendas + row.receb);
  const total3 = rounded(total2 - row.descon + row.juros);
  return { total1, total2, total3 };
}

function closeBrandToolbar() {
  const button = document.querySelector("#brandMenuButton");
  const toolbar = document.querySelector("#brandToolbar");
  const dim = document.querySelector("#pageDim");
  toolbar.hidden = true;
  dim.hidden = true;
  button.setAttribute("aria-expanded", "false");
}

function toggleBrandToolbar() {
  const button = document.querySelector("#brandMenuButton");
  const toolbar = document.querySelector("#brandToolbar");
  const dim = document.querySelector("#pageDim");
  const shouldOpen = toolbar.hidden;
  toolbar.hidden = !shouldOpen;
  dim.hidden = !shouldOpen;
  button.setAttribute("aria-expanded", String(shouldOpen));
}

function setStatus(kind, text) {
  const pill = document.querySelector("#dailyStatus");
  pill.className = `status-pill ${kind || ""}`.trim();
  pill.textContent = text;
}

async function serviceFetch(path, options = {}) {
  const response = await fetch(`${SERVICE_URL}${path}`, options);
  if (!response.ok) {
    const detail = await response.text();
    throw new Error(detail || `Erro ${response.status}`);
  }
  return response;
}

async function refreshWorkbook() {
  try {
    const status = await serviceFetch("/status").then((response) => response.json());
    serviceInfo = { online: true, ...status };
    setStatus("ok", "Serviço local OK");
    document.querySelector("#workbookState").textContent = status.exists ? "Encontrada" : "Nova";
    document.querySelector("#workbookPath").textContent = status.exists
      ? status.path
      : `Será criada em ${status.path}`;

    if (status.exists) {
      const workbookResponse = await serviceFetch("/workbook");
      const workbookBuffer = await workbookResponse.arrayBuffer();
      existingRows = parseWorkbook(workbookBuffer);
    } else {
      existingRows = { fichas: [], vales: [] };
    }
  } catch {
    serviceInfo = { online: false, exists: false, path: "" };
    existingRows = { fichas: [], vales: [] };
    setStatus("bad", "Serviço local offline");
    document.querySelector("#workbookState").textContent = "Offline";
    document.querySelector("#workbookPath").textContent =
      "Execute o serviço local para buscar ou criar a planilha automaticamente.";
  }
  render();
}

function sheetRows(workbook, expectedName) {
  const sheetName = workbook.SheetNames.find((name) => normalizeText(name) === normalizeText(expectedName));
  if (!sheetName) return [];
  return window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: "",
    raw: true,
  });
}

function findHeaderRow(rows, requiredHeaders) {
  return rows.findIndex((row) => {
    const labels = row.map(normalizeText);
    return requiredHeaders.every((header) => labels.includes(normalizeText(header)));
  });
}

function columnMap(headerRow) {
  const map = new Map();
  headerRow.forEach((cell, index) => map.set(normalizeText(cell), index));
  return map;
}

function valueAt(row, map, label) {
  return row[map.get(normalizeText(label))] ?? "";
}

function parseMainRows(rows) {
  const headerIndex = findHeaderRow(rows, ["DATA", "DINHEIRO", "TOTAL1"]);
  if (headerIndex === -1) return [];
  const map = columnMap(rows[headerIndex]);
  return rows.slice(headerIndex + 1)
    .map((row) => {
      const date = toIsoDate(valueAt(row, map, "DATA"));
      if (!date) return null;
      const base = {
        date,
        ecf: String(valueAt(row, map, "ECF") || ""),
        acerto: parseAmount(valueAt(row, map, "ACERTO")),
        dinheiro: parseAmount(valueAt(row, map, "DINHEIRO")),
        chAv: parseAmount(valueAt(row, map, "CH AV")),
        cartao: parseAmount(valueAt(row, map, "CARTÃO")),
        vales: parseAmount(valueAt(row, map, "VALES")),
        pre: parseAmount(valueAt(row, map, "PRÉ")),
        vendas: parseAmount(valueAt(row, map, "VENDAS")),
        receb: parseAmount(valueAt(row, map, "RECEB")),
        descon: parseAmount(valueAt(row, map, "DESCON")),
        juros: parseAmount(valueAt(row, map, "JUROS")),
        chPEmi: parseAmount(valueAt(row, map, "CH P/ EMI")),
      };
      return { ...base, ...calculateTotals(base) };
    })
    .filter(Boolean);
}

function parseValeRows(rows) {
  const headerIndex = findHeaderRow(rows, ["DATA", "DISCRIMINAÇÃO", "VALOR"]);
  if (headerIndex === -1) return [];
  const map = columnMap(rows[headerIndex]);
  return rows.slice(headerIndex + 1)
    .map((row) => {
      const date = toIsoDate(valueAt(row, map, "DATA"));
      if (!date) return null;
      return {
        id: `${date}-${Math.random()}`,
        date,
        descricao: String(valueAt(row, map, "DISCRIMINAÇÃO") || ""),
        valor: parseAmount(valueAt(row, map, "VALOR")),
        obs: String(valueAt(row, map, "OBS DIA") || ""),
      };
    })
    .filter(Boolean);
}

function parseWorkbook(buffer) {
  if (!window.XLSX) throw new Error("Leitor de Excel não carregou.");
  const workbook = window.XLSX.read(buffer, { type: "array", raw: true, cellDates: false });
  return {
    fichas: parseMainRows(sheetRows(workbook, "FICHADIARIAGERAL")),
    vales: parseValeRows(sheetRows(workbook, "VALES")),
  };
}

function mergedRows() {
  const draftDates = new Set(draft.fichas.map((row) => row.date));
  const draftValeDates = new Set(draft.vales.map((row) => row.date));
  const fichas = [
    ...existingRows.fichas.filter((row) => !draftDates.has(row.date)),
    ...draft.fichas,
  ].sort((left, right) => left.date.localeCompare(right.date));
  const vales = [
    ...existingRows.vales.filter((row) => !draftValeDates.has(row.date)),
    ...draft.vales,
  ].sort((left, right) => left.date.localeCompare(right.date));
  return { fichas, vales };
}

function aoaFromRows(rows) {
  return [
    mainHeaders,
    ...rows.map((row) => [
      displayDate(row.date),
      row.ecf,
      row.acerto,
      row.dinheiro,
      row.chAv,
      row.cartao,
      row.vales,
      row.total1,
      row.pre,
      row.vendas,
      row.receb,
      row.total2,
      row.descon,
      row.juros,
      row.total3,
      row.chPEmi,
    ]),
  ];
}

function valesAoaFromRows(rows) {
  return [
    valeHeaders,
    ...rows.map((row) => [displayDate(row.date), row.descricao, row.valor, row.obs]),
  ];
}

function appendSheet(workbook, name, rows) {
  const worksheet = window.XLSX.utils.aoa_to_sheet(rows);
  const columnCount = Math.max(...rows.map((row) => row.length));
  worksheet["!cols"] = Array.from({ length: columnCount }, (_, index) => {
    const width = Math.max(12, ...rows.map((row) => String(row[index] ?? "").length + 2));
    return { wch: Math.min(width, 28) };
  });
  window.XLSX.utils.book_append_sheet(workbook, worksheet, name);
}

function buildWorkbookBytes() {
  const rows = mergedRows();
  const workbook = window.XLSX.utils.book_new();
  appendSheet(workbook, "FICHADIARIAGERAL", aoaFromRows(rows.fichas));
  appendSheet(workbook, "VALES", valesAoaFromRows(rows.vales));
  return window.XLSX.write(workbook, { bookType: "xlsx", type: "array" });
}

function readDailyForm() {
  const base = {
    date: document.querySelector("#dailyDate").value,
    ecf: document.querySelector("#dailyEcf").value.trim(),
    acerto: parseAmount(document.querySelector("#dailyAcerto").value),
    dinheiro: parseAmount(document.querySelector("#dailyDinheiro").value),
    chAv: parseAmount(document.querySelector("#dailyChAv").value),
    cartao: parseAmount(document.querySelector("#dailyCartao").value),
    vales: parseAmount(document.querySelector("#dailyVales").value),
    pre: parseAmount(document.querySelector("#dailyPre").value),
    vendas: parseAmount(document.querySelector("#dailyVendas").value),
    receb: parseAmount(document.querySelector("#dailyReceb").value),
    descon: parseAmount(document.querySelector("#dailyDescon").value),
    juros: parseAmount(document.querySelector("#dailyJuros").value),
    chPEmi: parseAmount(document.querySelector("#dailyChPEmi").value),
  };
  return { ...base, ...calculateTotals(base) };
}

function updateDailyTotalPreview() {
  const row = readDailyForm();
  document.querySelector("#dailyTotalPreview").textContent =
    `Total 1: ${money.format(row.total1)} | Total 2: ${money.format(row.total2)} | Total 3: ${money.format(row.total3)}`;
}

function addOrReplaceDraftFicha(row) {
  draft.fichas = draft.fichas.filter((item) => item.date !== row.date);
  draft.fichas.push(row);
  saveDraft();
}

function renderPreview() {
  const rows = mergedRows().fichas;
  document.querySelector("#loadedDailyCount").textContent = String(existingRows.fichas.length);
  document.querySelector("#pendingDailyCount").textContent = String(draft.fichas.length + draft.vales.length);
  const body = document.querySelector("#dailyPreviewBody");
  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="8">Nenhuma ficha registrada.</td></tr>`;
    return;
  }
  body.innerHTML = rows.slice(-12).map((row) => `<tr>
    <td>${displayDate(row.date)}</td>
    <td>${money.format(row.acerto)}</td>
    <td>${money.format(row.dinheiro)}</td>
    <td>${money.format(row.chAv)}</td>
    <td>${money.format(row.cartao)}</td>
    <td>${money.format(row.total1)}</td>
    <td>${money.format(row.total2)}</td>
    <td>${money.format(row.total3)}</td>
  </tr>`).join("");
}

function renderDrafts() {
  const list = document.querySelector("#draftList");
  const total = draft.fichas.length + draft.vales.length;
  document.querySelector("#saveSummary").textContent = total
    ? `${draft.fichas.length} fichas e ${draft.vales.length} vales pendentes`
    : "Nenhum lançamento pendente";
  if (!total) {
    list.innerHTML = `<p class="empty-state">Nenhum lançamento pendente.</p>`;
    return;
  }
  const dailyRows = draft.fichas.map((row) => `<div class="return-row">
    <strong>${displayDate(row.date)}</strong>
    <span>Ficha diária - Total 1 ${money.format(row.total1)}</span>
    <strong>${money.format(row.total3)}</strong>
  </div>`);
  const valeRows = draft.vales.map((row) => `<div class="return-row">
    <strong>${displayDate(row.date)}</strong>
    <span>Vale - ${row.descricao}</span>
    <strong>${money.format(row.valor)}</strong>
  </div>`);
  list.innerHTML = [...dailyRows, ...valeRows].join("");
}

function render() {
  renderPreview();
  renderDrafts();
  updateDailyTotalPreview();
}

function clearForm(form) {
  form.reset();
  updateDailyTotalPreview();
}

async function saveWorkbook() {
  if (!serviceInfo.online) {
    window.alert("Serviço local offline. Execute o serviço local antes de salvar.");
    return;
  }
  const total = draft.fichas.length + draft.vales.length;
  if (!total) {
    window.alert("Não há lançamentos pendentes para salvar.");
    return;
  }
  const dates = [...new Set([...draft.fichas.map((row) => row.date), ...draft.vales.map((row) => row.date)])]
    .map(displayDate)
    .join(", ");
  const firstConfirm = window.confirm(
    `A planilha local será alterada.\n\nFichas: ${draft.fichas.length}\nVales: ${draft.vales.length}\nDatas afetadas: ${dates}\n\nDeseja continuar?`,
  );
  if (!firstConfirm) return;
  const secondConfirm = window.confirm(
    "Confirma novamente que deseja substituir a planilha atual? Um backup será criado antes da alteração.",
  );
  if (!secondConfirm) return;

  const bytes = buildWorkbookBytes();
  const response = await serviceFetch("/workbook", {
    method: "POST",
    headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
    body: bytes,
  }).then((item) => item.json());

  draft = { fichas: [], vales: [] };
  saveDraft();
  window.alert(`Planilha salva com sucesso.\n${response.path}`);
  await refreshWorkbook();
}

document.querySelector("#brandMenuButton").addEventListener("click", (event) => {
  event.stopPropagation();
  toggleBrandToolbar();
});

document.addEventListener("click", (event) => {
  if (!event.target.closest(".brand-menu")) closeBrandToolbar();
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeBrandToolbar();
});

document.querySelectorAll("[data-money-field]").forEach((input) => {
  input.addEventListener("input", updateDailyTotalPreview);
});

document.querySelector("#dailyForm").addEventListener("submit", (event) => {
  event.preventDefault();
  const row = readDailyForm();
  addOrReplaceDraftFicha(row);
  clearForm(event.target);
  render();
});

document.querySelector("#valeForm").addEventListener("submit", (event) => {
  event.preventDefault();
  draft.vales.push({
    id: `${Date.now()}`,
    date: document.querySelector("#valeDate").value,
    descricao: document.querySelector("#valeDescricao").value.trim(),
    valor: parseAmount(document.querySelector("#valeValor").value),
    obs: document.querySelector("#valeObs").value.trim(),
  });
  saveDraft();
  clearForm(event.target);
  render();
});

document.querySelector("#clearDraftBtn").addEventListener("click", () => {
  if (!window.confirm("Limpar todos os lançamentos pendentes?")) return;
  draft = { fichas: [], vales: [] };
  saveDraft();
  render();
});

document.querySelector("#refreshWorkbookBtn").addEventListener("click", refreshWorkbook);
document.querySelector("#saveWorkbookBtn").addEventListener("click", saveWorkbook);

refreshWorkbook();
