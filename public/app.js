import { parseSystemReport } from "./parse-cx-core.js";

const money = new Intl.NumberFormat("pt-BR", {
  style: "currency",
  currency: "BRL",
});

const state = {
  system: null,
  transactions: [],
  statuses: {},
  currentIndex: 0,
  adjustments: {},
  returnEntries: [],
  notes: {},
  step: "tickets",
};

const EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const desktopApi = window.noxxusDesktop || null;

const closeCashState = {
  existingRows: { fichas: [], vales: [] },
  draft: { fichas: [], vales: [] },
  editing: null,
  workbookPath: "",
  workbookFileName: "",
  workbookBaseDirectory: "",
  askDirectoryEverySave: false,
};

const closeMainHeaders = [
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

const closeValeHeaders = ["DATA", "DISCRIMINAÇÃO", "VALOR", "OBS DIA"];

const closeExcelBorder = {
  top: { style: "thin", color: { rgb: "666666" } },
  right: { style: "thin", color: { rgb: "666666" } },
  bottom: { style: "thin", color: { rgb: "666666" } },
  left: { style: "thin", color: { rgb: "666666" } },
};

const closeExcelColors = {
  white: "FFFFFF",
  mainGreen: "C6E0B4",
  mainPeach: "F8CBAD",
  mainAqua: "CCFFFF",
  mainYellow: "FFF2CC",
  valeGreen: "92D050",
  totalYellow: "FFFF00",
};

const closeMainColumnFills = [
  closeExcelColors.mainAqua,
  closeExcelColors.mainGreen,
  closeExcelColors.mainGreen,
  closeExcelColors.mainPeach,
  closeExcelColors.mainPeach,
  closeExcelColors.mainPeach,
  closeExcelColors.mainPeach,
  closeExcelColors.mainAqua,
  closeExcelColors.mainYellow,
  closeExcelColors.mainPeach,
  closeExcelColors.mainPeach,
  closeExcelColors.mainGreen,
  closeExcelColors.mainPeach,
  closeExcelColors.mainPeach,
  closeExcelColors.mainGreen,
  closeExcelColors.mainPeach,
];

const closeValeColumnFills = [
  closeExcelColors.valeGreen,
  closeExcelColors.white,
  closeExcelColors.valeGreen,
  closeExcelColors.white,
];

const closeMoneyColumnIndexes = new Set([2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]);
const closeValeMoneyColumnIndexes = new Set([2]);

const RETURN_TYPES = {
  baixa_conta: "A) Baixa na conta",
  sem_entrada: "B) Devolução sem entrada",
  nova_venda: "C) Devolução + nova venda",
};

const SALE_VALUE_MODES = {
  total_nota: "Valor total da nova nota",
  pago_cliente: "Apenas o que foi pago pelo cliente",
};

function parseAmount(value) {
  if (typeof value === "number") return value;
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

function diff(a, b) {
  return rounded(a - b);
}

function diffStatus(value) {
  return Math.abs(value) < 0.01 ? "Batendo" : "Divergente";
}

function setStatus(kind, text) {
  const pill = document.querySelector("#statusPill");
  pill.className = `status-pill ${kind || ""}`.trim();
  pill.textContent = text;
}

function formatMethods(item) {
  return item.paymentMethods.length ? item.paymentMethods.join(" + ") : "Sem forma informada";
}

function cleanText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function closeNormalizeText(value) {
  return cleanText(value);
}

function currentIsoDate() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
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

function amountInputValue(value) {
  return Number(value || 0) ? String(rounded(value)).replace(".", ",") : "";
}

function setCloseMoneyInput(selector, value) {
  document.querySelector(selector).value = amountInputValue(value);
}

function draftId(type) {
  return `${type}-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function closeCalculateTotals(row) {
  const total1 = rounded(row.dinheiro + row.chAv + row.cartao + row.vales);
  const total2 = rounded(total1 + row.pre + row.vendas - row.receb);
  const total3 = rounded(total2 - row.descon + row.juros);
  return { total1, total2, total3 };
}

function closeSheetRows(workbook, expectedName) {
  const sheetName = workbook.SheetNames.find((name) => closeNormalizeText(name) === closeNormalizeText(expectedName));
  if (!sheetName) return [];
  return window.XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: "",
    raw: true,
  });
}

function closeFindHeaderRow(rows, requiredHeaders) {
  return rows.findIndex((row) => {
    const labels = row.map(closeNormalizeText);
    return requiredHeaders.every((header) => labels.includes(closeNormalizeText(header)));
  });
}

function closeColumnMap(headerRow) {
  const map = new Map();
  headerRow.forEach((cell, index) => map.set(closeNormalizeText(cell), index));
  return map;
}

function closeValueAt(row, map, label) {
  return row[map.get(closeNormalizeText(label))] ?? "";
}

function closeParseMainRows(rows) {
  const headerIndex = closeFindHeaderRow(rows, ["DATA", "DINHEIRO", "TOTAL1"]);
  if (headerIndex === -1) return [];
  const map = closeColumnMap(rows[headerIndex]);
  return rows.slice(headerIndex + 1)
    .map((row) => {
      const date = toIsoDate(closeValueAt(row, map, "DATA"));
      if (!date) return null;
      const base = {
        date,
        ecf: String(closeValueAt(row, map, "ECF") || ""),
        acerto: parseAmount(closeValueAt(row, map, "ACERTO")),
        dinheiro: parseAmount(closeValueAt(row, map, "DINHEIRO")),
        chAv: parseAmount(closeValueAt(row, map, "CH AV")),
        cartao: parseAmount(closeValueAt(row, map, "CARTÃO")),
        vales: parseAmount(closeValueAt(row, map, "VALES")),
        pre: parseAmount(closeValueAt(row, map, "PRÉ")),
        vendas: parseAmount(closeValueAt(row, map, "VENDAS")),
        receb: parseAmount(closeValueAt(row, map, "RECEB")),
        descon: parseAmount(closeValueAt(row, map, "DESCON")),
        juros: parseAmount(closeValueAt(row, map, "JUROS")),
        chPEmi: parseAmount(closeValueAt(row, map, "CH P/ EMI")),
      };
      return { ...base, ...closeCalculateTotals(base) };
    })
    .filter(Boolean);
}

function closeParseValeRows(rows) {
  const headerIndex = closeFindHeaderRow(rows, ["DATA", "DISCRIMINAÇÃO", "VALOR"]);
  if (headerIndex === -1) return [];
  const map = closeColumnMap(rows[headerIndex]);
  return rows.slice(headerIndex + 1)
    .map((row) => {
      const date = toIsoDate(closeValueAt(row, map, "DATA"));
      if (!date) return null;
      return {
        id: draftId("existing-vale"),
        date,
        descricao: String(closeValueAt(row, map, "DISCRIMINAÇÃO") || ""),
        valor: parseAmount(closeValueAt(row, map, "VALOR")),
        obs: String(closeValueAt(row, map, "OBS DIA") || ""),
      };
    })
    .filter(Boolean);
}

function closeParseWorkbook(buffer) {
  if (!window.XLSX) throw new Error("Leitor de Excel não carregou.");
  const workbook = window.XLSX.read(buffer, { type: "array", raw: true, cellDates: false });
  return {
    fichas: closeParseMainRows(closeSheetRows(workbook, "FICHADIARIAGERAL")),
    vales: closeParseValeRows(closeSheetRows(workbook, "VALES")),
  };
}

function ticketItems() {
  return state.transactions.filter((item) => item.needsPhysicalCheck);
}

function returnItems() {
  return state.transactions.filter((item) => item.isReturn);
}

function neutralizedItems() {
  return state.transactions.filter((item) => item.neutralized || item.internalMovement);
}

function counts() {
  const items = ticketItems();
  const confirmed = items.filter((item) => state.statuses[item.id]?.status === "confirmed").length;
  const issues = items.filter((item) => state.statuses[item.id]?.status === "issue").length;
  const pending = items.length - confirmed - issues;
  return { total: items.length, confirmed, issues, pending };
}

function sumConfirmedPayments() {
  const base = {
    contaCliente: 0,
    dinheiro: 0,
    cheques: 0,
    pre: 0,
    cartaoDebito: 0,
    cartaoCredito: 0,
    vales: 0,
    parcelado: 0,
    outros: 0,
    totalCaixa: 0,
  };

  for (const item of state.transactions) {
    const isConfirmed = state.statuses[item.id]?.status === "confirmed";
    if (!isConfirmed && !item.autoConfirmed) continue;
    for (const key of Object.keys(base)) {
      if (key in item.payments) base[key] += item.payments[key];
    }
    base.totalCaixa += item.cashAmount;
  }

  for (const key of Object.keys(base)) base[key] = rounded(base[key]);
  return base;
}

function adjustment(name) {
  return parseAmount(state.adjustments[name]);
}

function returnEntriesTotal() {
  return rounded(state.returnEntries.reduce((total, entry) => total + entry.amount, 0));
}

function returnEntriesByType(type) {
  return rounded(
    state.returnEntries
      .filter((entry) => entry.type === type)
      .reduce((total, entry) => total + entry.amount, 0),
  );
}

function itemStatus(item) {
  if (item.autoConfirmed) return "Automático";
  const status = state.statuses[item.id]?.status || "pending";
  if (status === "confirmed") return "Confirmado";
  if (status === "issue") return "Com problema";
  return "Pendente";
}

function isCustomerAccountReceipt(item) {
  const text = cleanText(`${item.description} ${item.details?.join(" ") || ""}`);
  return (
    item.type === "REC" ||
    (text.includes("receb") && text.includes("cliente")) ||
    Math.abs(item.payments.contaCliente) > 0.009 ||
    Math.abs(item.payments.parcelado) > 0.009
  );
}

function customerName(item) {
  const detail = item.details?.find((line) => cleanText(line).length > 2);
  if (detail) return detail;
  return item.description.replace(/receb\.?\s*de\s*parcela\s*-\s*cliente/iu, "").trim();
}

function renderSystem() {
  const totals = state.system?.totals;
  const list = document.querySelector("#systemList");

  const rows = totals
    ? [
        ["Dinheiro", totals.dinheiro],
        ["Cheques/Pix", totals.cheques],
        ["Débito", totals.cartaoDebito],
        ["Crédito", totals.cartaoCredito],
        ["Pré", totals.pre],
        ["Vales", totals.vales],
        ["Parcelado", totals.parcelado],
        ["Total", totals.totalCaixa],
      ]
    : [
        ["Dinheiro", 0],
        ["Cheques/Pix", 0],
        ["Débito", 0],
        ["Crédito", 0],
        ["Pré", 0],
      ];

  list.innerHTML = rows
    .map(([label, value]) => `<div><span>${label}</span><strong>${money.format(value)}</strong></div>`)
    .join("");
  document.querySelector("#systemTotal").textContent = money.format(totals?.totalCaixa || 0);
}

function renderSummary() {
  const count = counts();
  document.querySelector("#checkedCount").textContent = `${count.confirmed}/${count.total}`;
  document.querySelector("#issueCount").textContent = String(count.issues);

  if (!state.system) {
    setStatus("", "Aguardando CX");
  } else if (count.issues) {
    setStatus("bad", "Com problemas");
  } else if (count.total && count.pending === 0) {
    setStatus("ok", "Tickets OK");
  } else {
    setStatus("", "Conferindo");
  }
}

function setStep(step) {
  state.step = step;
  document.querySelectorAll(".step").forEach((button) => {
    button.classList.toggle("active", button.dataset.step === step);
  });
  document.querySelectorAll(".tool-step").forEach((section) => {
    section.classList.toggle("active", section.id === `${step}Step`);
  });
}

function currentTicket() {
  const items = ticketItems();
  if (!items.length) return null;
  state.currentIndex = Math.max(0, Math.min(state.currentIndex, items.length - 1));
  return items[state.currentIndex];
}

function renderTicketCard() {
  const card = document.querySelector("#ticketCard");
  const item = currentTicket();
  if (!item) {
    card.innerHTML = `<p class="empty-state">Importe o arquivo CX para montar a fila de conferência.</p>`;
    return;
  }

  const status = state.statuses[item.id]?.status || "pending";
  const note = state.statuses[item.id]?.note || "";
  const statusText =
    status === "confirmed" ? "Confirmado" : status === "issue" ? "Com problema" : "Pendente";

  card.innerHTML = `
    <div class="ticket-main">
      <div>
        <span class="ticket-sequence">${item.sequence || "Sem sequência"}</span>
        <p class="empty-state">${item.description}</p>
        ${
          item.details?.length
            ? `<p class="empty-state">${item.details.join(" ")}</p>`
            : ""
        }
      </div>
      <div class="ticket-amount">
        <span>Valor do movimento</span>
        <strong>${money.format(item.amount)}</strong>
      </div>
    </div>

    <div class="ticket-meta">
      <div><span>Forma</span><strong>${formatMethods(item)}</strong></div>
      <div><span>Tipo</span><strong>${item.type}</strong></div>
      <div><span>Funcionário</span><strong>${item.employee || "-"}</strong></div>
      <div><span>Status</span><strong>${statusText}</strong></div>
    </div>

    ${note ? `<p class="empty-state">Obs.: ${note}</p>` : ""}

    <div class="ticket-actions">
      <button id="confirmTicket" type="button">Confirmar ticket físico</button>
      <button class="danger" id="issueTicket" type="button">Marcar problema</button>
    </div>
  `;

  document.querySelector("#confirmTicket").addEventListener("click", () => markTicket("confirmed"));
  document.querySelector("#issueTicket").addEventListener("click", () => markTicket("issue"));
}

function renderTicketList() {
  const items = ticketItems();
  const count = counts();
  document.querySelector("#ticketListHint").textContent =
    `${count.pending} pendentes, ${count.issues} com problema`;

  document.querySelector("#ticketList").innerHTML = items
    .map((item, index) => {
      const status = state.statuses[item.id]?.status || "pending";
      const statusText =
        status === "confirmed" ? "OK" : status === "issue" ? "Problema" : "Pendente";
      return `<div class="ticket-row ${status}">
        <strong>${item.sequence || "-"}</strong>
        <span>${formatMethods(item)}${item.details?.length ? ` - ${item.details.join(" ")}` : ""}</span>
        <strong>${money.format(item.amount)}</strong>
        <button data-jump="${index}" type="button">${statusText}</button>
      </div>`;
    })
    .join("");

  document.querySelectorAll("[data-jump]").forEach((button) => {
    button.addEventListener("click", () => {
      state.currentIndex = Number(button.dataset.jump);
      render();
    });
  });
}

function renderReturns() {
  const returns = returnItems();
  const list = document.querySelector("#returnList");
  if (!returns.length) {
    list.innerHTML = `<p class="empty-state">Nenhuma devolução foi identificada automaticamente no CX.</p>`;
    return;
  }

  list.innerHTML = returns
    .map(
      (item) => `<div class="return-row">
        <strong>${item.sequence || item.type}</strong>
        <span>${item.description} - ${formatMethods(item)}</span>
        <strong>${money.format(item.amount)}</strong>
      </div>`,
    )
    .join("");
}

function renderNeutralized() {
  const items = neutralizedItems();
  const list = document.querySelector("#neutralizedList");
  if (!items.length) {
    list.innerHTML = `<p class="empty-state">Nenhum movimento interno foi ajustado automaticamente.</p>`;
    return;
  }

  list.innerHTML = items
    .map((item) => {
      const detail = item.internalMovement
        ? "movimento interno entre formas"
        : `anula com ${item.neutralizedWith}`;
      return `<div class="return-row">
        <strong>${item.sequence}</strong>
        <span>${item.description} - ${detail}</span>
        <strong>${money.format(item.amount)}</strong>
      </div>`;
    })
    .join("");
}

function renderManualReturns() {
  const list = document.querySelector("#manualReturnsList");
  if (!list) return;
  if (!state.returnEntries.length) {
    list.innerHTML = `<p class="empty-state">Nenhuma devolução lançada manualmente.</p>`;
    return;
  }

  list.innerHTML = state.returnEntries
    .map((entry) => {
      const saleInfo = entry.type === "nova_venda"
        ? ` - nova nota: ${money.format(entry.saleValue)} (${SALE_VALUE_MODES[entry.saleMode]})`
        : "";
      return `<div class="return-row manual-return-row">
        <span>${RETURN_TYPES[entry.type]}${saleInfo}${entry.note ? ` - ${entry.note}` : ""}</span>
        <strong>${money.format(entry.amount)}</strong>
        <strong>${entry.createdAt}</strong>
        <button data-remove-return="${entry.id}" type="button">Remover</button>
      </div>`;
    })
    .join("");

  document.querySelectorAll("[data-remove-return]").forEach((button) => {
    button.addEventListener("click", () => {
      state.returnEntries = state.returnEntries.filter((entry) => entry.id !== button.dataset.removeReturn);
      renderManualReturns();
      renderChecks();
    });
  });
}

function buildCheckData() {
  const totals = state.system?.totals;
  const count = counts();
  const confirmed = sumConfirmedPayments();
  const manualVales = adjustment("valesManuais");
  const dinheiroContado = adjustment("dinheiroContado");
  const pixCnpj = adjustment("pixCnpj");
  const dinheiroFisicoAjustado = dinheiroContado + pixCnpj + manualVales;
  const hasDinheiroContado = String(state.adjustments.dinheiroContado || "").trim() !== "";
  const devolucoes = returnEntriesTotal();

  if (!totals) return [];

  const rows = [
    {
      label: "Tickets pendentes",
      left: count.pending,
      right: 0,
      difference: count.pending,
      status: count.pending === 0 ? "Batendo" : "Pendente",
      level: count.pending === 0 ? "ok" : "warn",
      detail: String(count.pending),
    },
    {
      label: "Tickets com problema",
      left: count.issues,
      right: 0,
      difference: count.issues,
      status: count.issues === 0 ? "Batendo" : "Divergente",
      level: count.issues === 0 ? "ok" : "bad",
      detail: String(count.issues),
    },
  ];

  const crossedRows = [
    ["Total confirmado x Total do sistema", confirmed.totalCaixa, totals.totalCaixa],
    ["Dinheiro confirmado x Dinheiro do sistema", confirmed.dinheiro, totals.dinheiro],
    ["Cheques/Pix confirmados x Cheques/Pix do sistema", confirmed.cheques, totals.cheques],
    [
      "Cartões confirmados x Débito + Crédito do sistema",
      confirmed.cartaoDebito + confirmed.cartaoCredito,
      totals.cartaoDebito + totals.cartaoCredito,
    ],
    ["Pré confirmado x Pré do sistema", confirmed.pre, totals.pre],
  ];

  if (hasDinheiroContado) {
    crossedRows.splice(2, 0, [
      "Dinheiro contado + Pix CNPJ + Vales x Dinheiro do sistema",
      dinheiroFisicoAjustado,
      totals.dinheiro,
    ]);
  } else {
    rows.push({
      label: "Dinheiro contado fisicamente",
      left: "",
      right: totals.dinheiro,
      difference: "",
      status: "Pendente",
      level: "warn",
      detail: "Informe para validar sobra ou falta",
    });
  }

  for (const [label, left, right] of crossedRows) {
    const difference = diff(left, right);
    const ok = Math.abs(difference) < 0.01;
    rows.push({
      label,
      left: rounded(left),
      right: rounded(right),
      difference,
      status: diffStatus(difference),
      level: ok ? "ok" : "bad",
      detail: ok ? "Batendo" : `Diferença: ${money.format(difference)}`,
    });
  }

  if (hasDinheiroContado) {
    rows.push(
      {
        label: "Pix CNPJ somado à contagem física",
        left: pixCnpj,
        right: "",
        difference: "",
        status: "Informativo",
        level: "warn",
        detail: money.format(pixCnpj),
      },
      {
        label: "Vales recompostos na contagem física",
        left: manualVales,
        right: "",
        difference: "",
        status: "Informativo",
        level: "warn",
        detail: money.format(manualVales),
      },
    );
  }

  rows.push({
    label: "Devoluções lançadas para revisão",
    left: devolucoes,
    right: "",
    difference: "",
    status: "Informativo",
    level: "warn",
    detail: money.format(devolucoes),
  });

  return rows;
}

function renderChecks() {
  const totals = state.system?.totals;
  const count = counts();

  if (!totals) {
    document.querySelector("#checks").innerHTML =
      `<div class="check-row warn"><span>Arquivo CX</span><strong>Importe o relatório para iniciar</strong></div>`;
    return;
  }

  const checks = buildCheckData();
  const allOk = count.pending === 0 && count.issues === 0 && checks.every((row) => row.level !== "bad");
  document.querySelector("#checks").innerHTML = checks
    .map(
      (row) => `<div class="check-row ${row.level}">
        <span>${row.label}</span>
        <strong>${row.detail}</strong>
      </div>`,
    )
    .join("");
  document.querySelector("#resultHint").textContent = allOk
    ? "Conferência final sem divergências"
    : "Revise os pontos destacados";
}

function sheetValue(value) {
  return value === null || value === undefined ? "" : value;
}

function appendSheet(workbook, name, rows) {
  const worksheet = window.XLSX.utils.aoa_to_sheet(rows);
  const columnCount = Math.max(...rows.map((row) => row.length));
  worksheet["!cols"] = Array.from({ length: columnCount }, (_, index) => {
    const width = Math.max(
      12,
      ...rows.map((row) => String(sheetValue(row[index])).length + 2),
    );
    return { wch: Math.min(width, 48) };
  });
  window.XLSX.utils.book_append_sheet(workbook, worksheet, name.slice(0, 31));
}

function closeFormulaCell(formula, value) {
  return { t: "n", f: formula, v: rounded(value) };
}

function closeAoaFromRows(rows) {
  return [
    closeMainHeaders,
    ...rows.map((row, index) => {
      const excelRow = index + 2;
      return [
        displayDate(row.date),
        row.ecf,
        row.acerto,
        row.dinheiro,
        row.chAv,
        row.cartao,
        row.vales,
        closeFormulaCell(`D${excelRow}+E${excelRow}+F${excelRow}+G${excelRow}`, row.total1),
        row.pre,
        row.vendas,
        row.receb,
        closeFormulaCell(`H${excelRow}+I${excelRow}+J${excelRow}-K${excelRow}`, row.total2),
        row.descon,
        row.juros,
        closeFormulaCell(`L${excelRow}-M${excelRow}+N${excelRow}`, row.total3),
        row.chPEmi,
      ];
    }),
  ];
}

function closeValesSheetData(rows) {
  const sheetRows = [closeValeHeaders];
  const totalRowIndexes = [];

  for (let index = 0; index < rows.length;) {
    const date = rows[index].date;
    const firstExcelRow = sheetRows.length + 1;
    let total = 0;

    while (index < rows.length && rows[index].date === date) {
      const row = rows[index];
      total += row.valor;
      sheetRows.push([displayDate(row.date), row.descricao, row.valor, row.obs]);
      index += 1;
    }

    const lastExcelRow = sheetRows.length;
    totalRowIndexes.push(sheetRows.length);
    sheetRows.push(["", "", closeFormulaCell(`SUM(C${firstExcelRow}:C${lastExcelRow})`, total), ""]);
  }

  return { rows: sheetRows, totalRowIndexes };
}

function closeCellDisplayValue(cell) {
  return cell && typeof cell === "object" && "v" in cell ? cell.v : cell;
}

function closeApplyCellStyle(worksheet, rowIndex, columnIndex, style, numberFormat = "") {
  const cellRef = window.XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
  worksheet[cellRef] = worksheet[cellRef] || { t: "s", v: "" };
  const currentStyle = worksheet[cellRef].s || {};
  worksheet[cellRef].s = {
    ...currentStyle,
    ...style,
    font: { ...(currentStyle.font || {}), ...(style.font || {}) },
    fill: style.fill || currentStyle.fill,
    alignment: { ...(currentStyle.alignment || {}), ...(style.alignment || {}) },
    border: style.border || currentStyle.border,
  };
  if (numberFormat) worksheet[cellRef].z = numberFormat;
}

function closeColumnStyle(fillColor, options = {}) {
  return {
    fill: { fgColor: { rgb: fillColor }, patternType: "solid" },
    font: {
      bold: Boolean(options.bold),
      color: { rgb: "000000" },
    },
    alignment: {
      horizontal: options.horizontal || "center",
      vertical: "center",
    },
    border: closeExcelBorder,
  };
}

function closeStyleMainSheet(worksheet, rows) {
  const columnCount = closeMainHeaders.length;
  for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
    closeApplyCellStyle(
      worksheet,
      0,
      columnIndex,
      closeColumnStyle(closeMainColumnFills[columnIndex], { bold: true }),
    );
  }

  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const isMoney = closeMoneyColumnIndexes.has(columnIndex);
      closeApplyCellStyle(
        worksheet,
        rowIndex,
        columnIndex,
        closeColumnStyle(closeMainColumnFills[columnIndex], { horizontal: isMoney ? "right" : "center" }),
        isMoney ? "#,##0.00" : "",
      );
    }
  }
}

function closeStyleValesSheet(worksheet, rows, totalRowIndexes = []) {
  const columnCount = closeValeHeaders.length;
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const isHeader = rowIndex === 0;
      const isMoney = closeValeMoneyColumnIndexes.has(columnIndex);
      closeApplyCellStyle(
        worksheet,
        rowIndex,
        columnIndex,
        closeColumnStyle(closeValeColumnFills[columnIndex], {
          bold: isHeader,
          horizontal: isMoney && !isHeader ? "right" : "center",
        }),
        isMoney && !isHeader ? "#,##0.00" : "",
      );
    }
  }

  totalRowIndexes.forEach((rowIndex) => {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      closeApplyCellStyle(
        worksheet,
        rowIndex,
        columnIndex,
        closeColumnStyle(closeExcelColors.totalYellow, {
          bold: true,
          horizontal: columnIndex === 2 ? "right" : "center",
        }),
        columnIndex === 2 ? "#,##0.00" : "",
      );
    }
  });
}

function closeAppendWorkbookSheet(workbook, name, rows, options = {}) {
  const worksheet = window.XLSX.utils.aoa_to_sheet(rows);
  const columnCount = Math.max(...rows.map((row) => row.length));
  worksheet["!cols"] = Array.from({ length: columnCount }, (_, index) => {
    const width = Math.max(12, ...rows.map((row) => String(closeCellDisplayValue(row[index]) ?? "").length + 2));
    return { wch: Math.min(width, 28) };
  });
  if (options.kind === "main") closeStyleMainSheet(worksheet, rows);
  if (options.kind === "vales") closeStyleValesSheet(worksheet, rows, options.totalRowIndexes);
  window.XLSX.utils.book_append_sheet(workbook, worksheet, name);
}

function closeBuildWorkbookBytes(rows) {
  const workbook = window.XLSX.utils.book_new();
  const valesSheet = closeValesSheetData(rows.vales);
  closeAppendWorkbookSheet(workbook, "FICHADIARIAGERAL", closeAoaFromRows(rows.fichas), { kind: "main" });
  closeAppendWorkbookSheet(workbook, "VALES", valesSheet.rows, {
    kind: "vales",
    totalRowIndexes: valesSheet.totalRowIndexes,
  });
  return window.XLSX.write(workbook, { bookType: "xlsx", type: "array", cellStyles: true });
}

function paymentColumns(item) {
  return [
    item.payments.dinheiro,
    item.payments.cheques,
    item.payments.pre,
    item.payments.cartaoDebito,
    item.payments.cartaoCredito,
    item.payments.vales,
    item.payments.parcelado,
    item.payments.contaCliente,
    item.payments.outros,
  ].map(rounded);
}

function transactionRows(items) {
  return items.map((item) => [
    item.row,
    item.sequence,
    item.description,
    item.details?.join(" "),
    item.employee,
    item.type,
    formatMethods(item),
    itemStatus(item),
    state.statuses[item.id]?.note || "",
    rounded(item.amount),
    rounded(item.cashAmount),
    ...paymentColumns(item),
  ]);
}

function transactionHeader() {
  return [
    "Linha CX",
    "Sequência",
    "Descrição",
    "Detalhes",
    "Funcionário",
    "Tipo",
    "Formas",
    "Status",
    "Observação",
    "Valor movimento",
    "Total caixa",
    "Dinheiro",
    "Cheques/Pix",
    "Pré",
    "Débito",
    "Crédito",
    "Vales",
    "Parcelado",
    "Conta cliente",
    "Outros",
  ];
}

function adjustmentRows() {
  const manualVales = adjustment("valesManuais");
  const dinheiroContado = adjustment("dinheiroContado");
  const pixCnpj = adjustment("pixCnpj");

  return [
    ["Total de vales", manualVales, state.notes.vales || ""],
    ["Dinheiro contado", dinheiroContado, state.notes.dinheiro || ""],
    ["Pix CNPJ", pixCnpj, "Somado à contagem física do dinheiro"],
    ["Devolução A - baixa na conta", returnEntriesByType("baixa_conta"), "Total dos lançamentos individuais"],
    ["Devolução B - sem entrada", returnEntriesByType("sem_entrada"), "Total dos lançamentos individuais"],
    ["Devolução C - nova venda", returnEntriesByType("nova_venda"), "Total dos lançamentos individuais"],
    ["Total de devoluções lançadas", returnEntriesTotal(), ""],
  ];
}

function manualReturnRows() {
  return state.returnEntries.map((entry, index) => [
    index + 1,
    RETURN_TYPES[entry.type],
    entry.amount,
    entry.type === "nova_venda" ? entry.saleValue : "",
    entry.type === "nova_venda" ? SALE_VALUE_MODES[entry.saleMode] : "",
    entry.note,
    entry.createdAt,
  ]);
}

function customerAccountRows() {
  return state.transactions
    .filter(isCustomerAccountReceipt)
    .map((item) => [
      item.row,
      item.sequence,
      customerName(item),
      item.description,
      item.details?.join(" "),
      item.employee,
      item.type,
      itemStatus(item),
      rounded(item.amount),
      rounded(item.payments.contaCliente),
      rounded(item.payments.parcelado),
      rounded(item.payments.dinheiro),
      rounded(item.payments.cheques),
      rounded(item.payments.cartaoDebito),
      rounded(item.payments.cartaoCredito),
      formatMethods(item),
    ]);
}

function issueRows() {
  return ticketItems()
    .filter((item) => {
      const status = state.statuses[item.id]?.status || "pending";
      return status !== "confirmed";
    })
    .map((item) => [
      item.row,
      item.sequence,
      item.description,
      item.details?.join(" "),
      item.employee,
      formatMethods(item),
      itemStatus(item),
      state.statuses[item.id]?.note || "",
      rounded(item.amount),
    ]);
}

function exportReport() {
  if (!state.system) {
    window.alert("Importe o arquivo CX antes de exportar o relatório.");
    return;
  }
  if (!window.XLSX) {
    window.alert("Exportador de Excel não carregou. Atualize a página e tente novamente.");
    return;
  }

  const workbook = window.XLSX.utils.book_new();
  const totals = state.system.totals;
  const confirmed = sumConfirmedPayments();
  const count = counts();
  const metadata = state.system.metadata || {};
  const stamp = new Date();
  const fileStamp = stamp.toISOString().slice(0, 19).replace(/[:T]/g, "-");
  const checkRows = buildCheckData();
  const automaticItems = neutralizedItems();
  const customerRows = customerAccountRows();
  const issues = issueRows();
  const returnTotal = returnItems().reduce((total, item) => total + item.amount, 0);
  const automaticTotal = automaticItems.reduce((total, item) => total + item.amount, 0);
  const manualAdjustments = adjustmentRows();
  const manualReturns = manualReturnRows();
  const finalStatus =
    count.pending === 0 && count.issues === 0 && checkRows.every((row) => row.level !== "bad")
      ? "Conferência sem divergências"
      : "Conferência com pontos para revisar";

  appendSheet(workbook, "Resumo", [
    ["Relatório de conferência de caixa"],
    ["Gerado em", stamp.toLocaleString("pt-BR")],
    ["Empresa", metadata.empresa || ""],
    ["Caixa", metadata.caixa || ""],
    ["Linhas lidas", metadata.linhasLidas || ""],
    ["Status final", finalStatus],
    [],
    ["Indicador", "Valor"],
    ["Tickets físicos", count.total],
    ["Tickets confirmados", count.confirmed],
    ["Tickets pendentes", count.pending],
    ["Tickets com problema", count.issues],
    ["Movimentos automáticos", automaticItems.length],
    ["Contas de clientes recebidas", customerRows.length],
    ["Devoluções encontradas no CX", returnItems().length],
    ["Devoluções lançadas manualmente", manualReturns.length],
    [],
    ["Caixa conferido", "Sistema", "Confirmado/Calculado", "Diferença"],
    ["Total caixa", totals.totalCaixa, confirmed.totalCaixa, diff(confirmed.totalCaixa, totals.totalCaixa)],
    ["Dinheiro", totals.dinheiro, confirmed.dinheiro, diff(confirmed.dinheiro, totals.dinheiro)],
    ["Cheques/Pix", totals.cheques, confirmed.cheques, diff(confirmed.cheques, totals.cheques)],
    ["Débito", totals.cartaoDebito, confirmed.cartaoDebito, diff(confirmed.cartaoDebito, totals.cartaoDebito)],
    ["Crédito", totals.cartaoCredito, confirmed.cartaoCredito, diff(confirmed.cartaoCredito, totals.cartaoCredito)],
    ["Pré", totals.pre, confirmed.pre, diff(confirmed.pre, totals.pre)],
    [],
    ["Valores fora da conferência física", "Valor", "Observação"],
    ["Vales do sistema", totals.vales, "Informativo"],
    ["Parcelado do sistema", totals.parcelado, "Não entra como ticket físico"],
    ["Contas de clientes no sistema", totals.contaCliente, "Recebimentos/baixas identificados no CX"],
    ["Devoluções encontradas no CX", rounded(returnTotal), "Detalhes na aba Devoluções"],
    ["Movimentos automáticos", rounded(automaticTotal), "Detalhes na aba Movimentos automáticos"],
    [],
    ["Ajustes manuais informados", "Valor", "Observação"],
    ...manualAdjustments,
  ]);

  appendSheet(workbook, "Checks", [
    ["Check", "Valor conferido", "Valor sistema", "Diferença", "Status", "Detalhe"],
    ...checkRows.map((row) => [
      row.label,
      row.left,
      row.right,
      row.difference,
      row.status,
      row.detail,
    ]),
  ]);

  appendSheet(workbook, "Divergencias", [
    ["Check", "Valor conferido", "Valor sistema", "Diferença", "Status", "Detalhe"],
    ...checkRows
      .filter((row) => row.level !== "ok")
      .map((row) => [row.label, row.left, row.right, row.difference, row.status, row.detail]),
    [],
    ["Tickets pendentes ou com problema"],
    ["Linha CX", "Sequência", "Descrição", "Detalhes", "Funcionário", "Formas", "Status", "Observação", "Valor"],
    ...issues,
  ]);

  appendSheet(workbook, "Tickets", [
    transactionHeader(),
    ...transactionRows(ticketItems()),
  ]);

  appendSheet(workbook, "Movimentos automáticos", [
    transactionHeader(),
    ...transactionRows(neutralizedItems()),
  ]);

  appendSheet(workbook, "Devoluções", [
    transactionHeader(),
    ...transactionRows(returnItems()),
  ]);

  appendSheet(workbook, "Devoluções lançadas", [
    ["Lançamento", "Tipo", "Valor devolvido", "Valor nova nota", "Tipo valor nova nota", "Observação", "Criado em"],
    ...manualReturns,
  ]);

  appendSheet(workbook, "Contas clientes", [
    [
      "Linha CX",
      "Sequência",
      "Cliente",
      "Descrição",
      "Detalhes",
      "Funcionário",
      "Tipo",
      "Status",
      "Valor movimento",
      "Conta cliente",
      "Parcelado",
      "Dinheiro",
      "Cheques/Pix",
      "Débito",
      "Crédito",
      "Formas",
    ],
    ...customerRows,
  ]);

  appendSheet(workbook, "Todos movimentos CX", [
    transactionHeader(),
    ...transactionRows(state.transactions),
  ]);

  appendSheet(workbook, "Ajustes manuais", [
    ["Campo", "Valor", "Observação"],
    ...manualAdjustments,
  ]);

  window.XLSX.writeFile(workbook, `relatório-conferência-caixa-${fileStamp}.xlsx`);
}

function markTicket(status) {
  const item = currentTicket();
  if (!item) return;
  let note = "";
  if (status === "issue") {
    note = window.prompt("Qual foi o problema encontrado neste ticket?", "") || "Problema marcado";
  }
  state.statuses[item.id] = { status, note };
  goNextPending();
}

function goNextPending() {
  const items = ticketItems();
  const next = items.findIndex((item, index) => {
    return index > state.currentIndex && !state.statuses[item.id];
  });
  if (next >= 0) {
    state.currentIndex = next;
  } else {
    const first = items.findIndex((item) => !state.statuses[item.id]);
    if (first >= 0) state.currentIndex = first;
  }
  render();
}

function toggleReturnSaleFields() {
  const isNewSale = document.querySelector("#returnType").value === "nova_venda";
  document.querySelectorAll(".new-sale-field").forEach((field) => {
    field.hidden = !isNewSale;
  });
}

function addReturnEntry() {
  const type = document.querySelector("#returnType").value;
  const amountInput = document.querySelector("#returnAmount");
  const saleValueInput = document.querySelector("#returnSaleValue");
  const saleModeInput = document.querySelector("#returnSaleMode");
  const noteInput = document.querySelector("#returnNote");
  const amount = parseAmount(amountInput.value);
  const saleValue = type === "nova_venda" ? parseAmount(saleValueInput.value) : 0;

  if (amount <= 0) {
    window.alert("Informe o valor devolvido para adicionar o lançamento.");
    return;
  }

  state.returnEntries.push({
    id: `${Date.now()}-${state.returnEntries.length}`,
    type,
    amount: rounded(amount),
    saleValue: rounded(saleValue),
    saleMode: saleModeInput.value,
    note: noteInput.value.trim(),
    createdAt: new Date().toLocaleString("pt-BR"),
  });

  amountInput.value = "";
  saleValueInput.value = "";
  noteInput.value = "";
  renderManualReturns();
  renderChecks();
}

function closeCashTotals() {
  const totals = state.system?.totals || {};
  return {
    dinheiro: rounded(totals.dinheiro || 0),
    chAv: rounded(totals.cheques || 0),
    cartao: rounded((totals.cartaoDebito || 0) + (totals.cartaoCredito || 0)),
  };
}

function closeApplyWorkbookStatus(status) {
  closeCashState.workbookPath = status.path || "";
  closeCashState.workbookFileName = status.exists ? status.fileName : "";
  closeCashState.workbookBaseDirectory = status.baseDirectory || "";
  closeCashState.askDirectoryEverySave = Boolean(status.askDirectoryEverySave);
  document.querySelector("#closeWorkbookPath").textContent = closeCashState.workbookPath
    ? closeCashState.workbookPath
    : "O app buscará a planilha automaticamente nos arquivos.";
}

async function closeRefreshDesktopWorkbook() {
  if (!desktopApi?.readDailyWorkbook) {
    closeCashState.existingRows = { fichas: [], vales: [] };
    document.querySelector("#closeWorkbookPath").textContent =
      "Salvamento automático disponível apenas no aplicativo instalado.";
    return;
  }

  const payload = await desktopApi.readDailyWorkbook();
  closeApplyWorkbookStatus(payload);
  if (payload.exists && payload.bytes?.length) {
    const bytes = new Uint8Array(payload.bytes);
    closeCashState.existingRows = closeParseWorkbook(bytes.buffer);
  } else {
    closeCashState.existingRows = { fichas: [], vales: [] };
  }
}

async function closeChooseWorkbookFolder() {
  if (!desktopApi?.chooseDailyWorkbookDirectory) {
    window.alert("A escolha de pasta está disponível apenas no aplicativo instalado.");
    return;
  }

  try {
    const payload = await desktopApi.chooseDailyWorkbookDirectory();
    if (payload.canceled) return;
    closeApplyWorkbookStatus(payload);
    await closeRefreshDesktopWorkbook();
    closeRender();
  } catch (error) {
    window.alert(`Não consegui configurar a pasta.\n${error.message || error}`);
  }
}

function closeMergedRows() {
  const draftDates = new Set(closeCashState.draft.fichas.map((row) => row.date));
  const fichas = [
    ...closeCashState.existingRows.fichas.filter((row) => !draftDates.has(row.date)),
    ...closeCashState.draft.fichas,
  ].sort((left, right) => left.date.localeCompare(right.date));

  const vales = [
    ...closeCashState.existingRows.vales,
    ...closeCashState.draft.vales,
  ].sort((left, right) => left.date.localeCompare(right.date));

  return { fichas, vales };
}

function readCloseDailyForm() {
  const base = {
    date: document.querySelector("#closeDailyDate").value,
    ecf: document.querySelector("#closeDailyEcf").value.trim(),
    acerto: parseAmount(document.querySelector("#closeDailyAcerto").value),
    dinheiro: parseAmount(document.querySelector("#closeDailyDinheiro").value),
    chAv: parseAmount(document.querySelector("#closeDailyChAv").value),
    cartao: parseAmount(document.querySelector("#closeDailyCartao").value),
    vales: parseAmount(document.querySelector("#closeDailyVales").value),
    pre: parseAmount(document.querySelector("#closeDailyPre").value),
    vendas: parseAmount(document.querySelector("#closeDailyVendas").value),
    receb: parseAmount(document.querySelector("#closeDailyReceb").value),
    descon: parseAmount(document.querySelector("#closeDailyDescon").value),
    juros: parseAmount(document.querySelector("#closeDailyJuros").value),
    chPEmi: parseAmount(document.querySelector("#closeDailyChPEmi").value),
  };
  return { ...base, ...closeCalculateTotals(base) };
}

function fillCloseDailyForm(row) {
  document.querySelector("#closeDailyDate").value = row.date || currentIsoDate();
  document.querySelector("#closeDailyEcf").value = row.ecf || "";
  setCloseMoneyInput("#closeDailyAcerto", row.acerto);
  setCloseMoneyInput("#closeDailyDinheiro", row.dinheiro);
  setCloseMoneyInput("#closeDailyChAv", row.chAv);
  setCloseMoneyInput("#closeDailyCartao", row.cartao);
  setCloseMoneyInput("#closeDailyVales", row.vales);
  setCloseMoneyInput("#closeDailyPre", row.pre);
  setCloseMoneyInput("#closeDailyVendas", row.vendas);
  setCloseMoneyInput("#closeDailyReceb", row.receb);
  setCloseMoneyInput("#closeDailyDescon", row.descon);
  setCloseMoneyInput("#closeDailyJuros", row.juros);
  setCloseMoneyInput("#closeDailyChPEmi", row.chPEmi);
  updateCloseDailyTotalPreview();
}

function fillCloseValeForm(row) {
  document.querySelector("#closeValeDate").value = row.date || currentIsoDate();
  document.querySelector("#closeValeDescricao").value = row.descricao || "";
  setCloseMoneyInput("#closeValeValor", row.valor);
  document.querySelector("#closeValeObs").value = row.obs || "";
}

function resetCloseAutomaticFields() {
  const today = currentIsoDate();
  const totals = closeCashTotals();
  document.querySelector("#closeDailyDate").value = today;
  document.querySelector("#closeValeDate").value = today;
  document.querySelector("#closeCashDateLabel").textContent = displayDate(today);
  setCloseMoneyInput("#closeDailyDinheiro", totals.dinheiro);
  setCloseMoneyInput("#closeDailyChAv", totals.chAv);
  setCloseMoneyInput("#closeDailyCartao", totals.cartao);
  updateCloseDailyTotalPreview();
}

function updateCloseDailyTotalPreview() {
  const row = readCloseDailyForm();
  document.querySelector("#closeDailyTotalPreview").textContent =
    `Total 1: ${money.format(row.total1)} | Total 2: ${money.format(row.total2)} | Total 3: ${money.format(row.total3)}`;
}

function updateCloseSubmitButtons() {
  document.querySelector("#closeDailySubmitBtn").textContent =
    closeCashState.editing?.type === "ficha" ? "Atualizar ficha" : "Adicionar ficha";
  document.querySelector("#closeValeSubmitBtn").textContent =
    closeCashState.editing?.type === "vale" ? "Atualizar vale" : "Adicionar vale";
}

function addOrReplaceCloseDraftFicha(row) {
  const editingDate = closeCashState.editing?.type === "ficha" ? closeCashState.editing.id : "";
  closeCashState.draft.fichas = closeCashState.draft.fichas
    .filter((item) => item.date !== row.date && item.date !== editingDate);
  closeCashState.draft.fichas.push(row);
  closeCashState.editing = null;
}

function editCloseDraftEntry(type, id) {
  if (type === "ficha") {
    const row = closeCashState.draft.fichas.find((item) => item.date === id);
    if (!row) return;
    closeCashState.editing = { type, id };
    fillCloseDailyForm(row);
    document.querySelector("#closeDailyEcf").focus();
  } else {
    const row = closeCashState.draft.vales.find((item) => item.id === id);
    if (!row) return;
    closeCashState.editing = { type, id };
    fillCloseValeForm(row);
    document.querySelector("#closeValeDescricao").focus();
  }
  updateCloseSubmitButtons();
}

function removeCloseDraftEntry(type, id) {
  if (!window.confirm("Remover este lançamento pendente?")) return;
  if (type === "ficha") {
    closeCashState.draft.fichas = closeCashState.draft.fichas.filter((item) => item.date !== id);
  } else {
    closeCashState.draft.vales = closeCashState.draft.vales.filter((item) => item.id !== id);
  }
  if (closeCashState.editing?.type === type && closeCashState.editing.id === id) {
    closeCashState.editing = null;
  }
  closeRender();
}

function closeClearForm(form, preserveSelector = "") {
  const preserved = preserveSelector ? document.querySelector(preserveSelector).value : "";
  form.reset();
  if (preserveSelector) document.querySelector(preserveSelector).value = preserved;
  resetCloseAutomaticFields();
}

function closeRenderPreview() {
  const rows = closeMergedRows().fichas;
  const body = document.querySelector("#closeDailyPreviewBody");
  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="8">Nenhuma ficha registrada.</td></tr>`;
    return;
  }

  body.innerHTML = rows.slice(-5).map((row) => `<tr>
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

function closeRenderDrafts() {
  const list = document.querySelector("#closeDraftList");
  const total = closeCashState.draft.fichas.length + closeCashState.draft.vales.length;
  document.querySelector("#closeCashFichaCount").textContent = String(closeCashState.draft.fichas.length);
  document.querySelector("#closeCashValeCount").textContent = String(closeCashState.draft.vales.length);
  document.querySelector("#closeSaveSummary").textContent = total
    ? `${closeCashState.draft.fichas.length} fichas e ${closeCashState.draft.vales.length} vales pendentes`
    : "Nenhum lançamento pendente";

  if (!total) {
    list.innerHTML = `<p class="empty-state">Nenhum lançamento pendente.</p>`;
    return;
  }

  const dailyRows = closeCashState.draft.fichas.map((row) => `<div class="return-row">
    <strong>${displayDate(row.date)}</strong>
    <span>Ficha diária - Total 1 ${money.format(row.total1)}</span>
    <strong>${money.format(row.total3)}</strong>
  </div>`);
  const valeRows = closeCashState.draft.vales.map((row) => `<div class="return-row">
    <strong>${displayDate(row.date)}</strong>
    <span>Vale - ${row.descricao}</span>
    <strong>${money.format(row.valor)}</strong>
  </div>`);

  list.innerHTML = [...dailyRows, ...valeRows].join("");
  list.querySelectorAll(".return-row").forEach((rowElement, index) => {
    const isFicha = index < closeCashState.draft.fichas.length;
    const type = isFicha ? "ficha" : "vale";
    const entry = isFicha
      ? closeCashState.draft.fichas[index]
      : closeCashState.draft.vales[index - closeCashState.draft.fichas.length];
    const id = isFicha ? entry.date : entry.id;
    rowElement.classList.add("draft-row");
    rowElement.insertAdjacentHTML("beforeend", `
      <div class="draft-actions">
        <button class="ghost" data-close-edit-type="${type}" data-close-edit-id="${id}" type="button">Editar</button>
        <button class="danger" data-close-remove-type="${type}" data-close-remove-id="${id}" type="button">Remover</button>
      </div>
    `);
  });

  document.querySelectorAll("[data-close-edit-type]").forEach((button) => {
    button.addEventListener("click", () => editCloseDraftEntry(button.dataset.closeEditType, button.dataset.closeEditId));
  });
  document.querySelectorAll("[data-close-remove-type]").forEach((button) => {
    button.addEventListener("click", () => removeCloseDraftEntry(button.dataset.closeRemoveType, button.dataset.closeRemoveId));
  });
}

function closeRender() {
  closeRenderPreview();
  closeRenderDrafts();
  updateCloseDailyTotalPreview();
  updateCloseSubmitButtons();
}

function closeDownloadBytes(bytes, fileName) {
  const blob = new Blob([bytes], { type: EXCEL_MIME });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function closeWriteWorkbook(bytes) {
  if (desktopApi?.saveDailyWorkbook) {
    const response = await desktopApi.saveDailyWorkbook(Array.from(new Uint8Array(bytes)));
    closeApplyWorkbookStatus(response);
    return { mode: "desktop", ...response };
  }

  closeDownloadBytes(bytes, "fichas-diarias.xlsx");
  return { mode: "download" };
}

async function saveCloseCashWorkbook() {
  const total = closeCashState.draft.fichas.length + closeCashState.draft.vales.length;
  if (!total) {
    window.alert("Não há lançamentos pendentes para salvar.");
    return;
  }

  const dates = [...new Set([
    ...closeCashState.draft.fichas.map((row) => row.date),
    ...closeCashState.draft.vales.map((row) => row.date),
  ])].map(displayDate).join(", ");
  const isUpdating = Boolean(
    closeCashState.workbookFileName ||
    closeCashState.existingRows.fichas.length ||
    closeCashState.existingRows.vales.length,
  );
  const targetText = desktopApi
    ? closeCashState.askDirectoryEverySave
      ? `O app pedirá a pasta antes de salvar.\nPasta atual:\n${closeCashState.workbookBaseDirectory || closeCashState.workbookPath || "não definida"}`
      : `A planilha será salva em:\n${closeCashState.workbookPath || closeCashState.workbookBaseDirectory || "pasta padrão do Noxxus System"}`
    : "Uma planilha será baixada pelo navegador.";
  const confirmed = window.confirm(
    `${isUpdating ? "Gostaria de atualizar a planilha?" : "Gostaria de adicionar a planilha?"}\n\n${targetText}\n\nFichas: ${closeCashState.draft.fichas.length}\nVales: ${closeCashState.draft.vales.length}\nDatas afetadas: ${dates}\n\nSe a planilha já existe, um backup será criado antes da alteração.`,
  );
  if (!confirmed) return;

  try {
    const rowsToSave = closeMergedRows();
    const bytes = closeBuildWorkbookBytes(rowsToSave);
    const result = await closeWriteWorkbook(bytes);
    closeCashState.existingRows = rowsToSave;
    closeCashState.draft = { fichas: [], vales: [] };
    closeCashState.editing = null;
    const backupText = result.backupPath ? `\nBackup criado em:\n${result.backupPath}` : "";
    window.alert(
      result.mode === "download"
        ? "Planilha gerada e baixada."
        : `Planilha salva com sucesso em:\n${result.path || closeCashState.workbookFileName}${backupText}`,
    );
    closeRender();
  } catch (error) {
    if (error.message !== "Escolha de pasta cancelada.") {
      window.alert(`Não consegui salvar a planilha.\n${error.message || error}`);
    }
  }
}

async function openCloseCashModal() {
  if (!state.system?.totals) {
    window.alert("Importe o arquivo CX antes de fechar o caixa.");
    return;
  }

  const modal = document.querySelector("#closeCashModal");
  modal.hidden = false;
  resetCloseAutomaticFields();

  try {
    await closeRefreshDesktopWorkbook();
  } catch (error) {
    document.querySelector("#closeWorkbookPath").textContent = error.message || String(error);
  }

  closeRender();
}

function closeCloseCashModal() {
  document.querySelector("#closeCashModal").hidden = true;
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

function render() {
  renderSystem();
  renderSummary();
  renderTicketCard();
  renderTicketList();
  renderReturns();
  renderNeutralized();
  renderManualReturns();
  renderChecks();
}

function resetConference() {
  if (!window.confirm("Limpar tudo, inclusive o arquivo CX carregado?")) return;
  state.system = null;
  state.transactions = [];
  state.statuses = {};
  state.currentIndex = 0;
  state.adjustments = {};
  state.returnEntries = [];
  state.notes = {};
  document.querySelector("#cxFile").value = "";
  document.querySelector("#fileName").textContent = "Nenhum arquivo";
  document.querySelector("#importError").hidden = true;
  document.querySelector("#importError").textContent = "";
  document.querySelectorAll("[data-adjustment]").forEach((input) => {
    input.value = "";
  });
  document.querySelectorAll("[data-note]").forEach((input) => {
    input.value = "";
  });
  document.querySelector("#returnType").value = "baixa_conta";
  document.querySelector("#returnSaleMode").value = "total_nota";
  document.querySelector("#returnAmount").value = "";
  document.querySelector("#returnSaleValue").value = "";
  document.querySelector("#returnNote").value = "";
  toggleReturnSaleFields();
  setStep("tickets");
  closeCashState.draft = { fichas: [], vales: [] };
  closeCashState.editing = null;
  render();
}

async function parseCx(file) {
  if (!window.XLSX) {
    throw new Error("Leitor de Excel não carregou. Atualize a página e tente novamente.");
  }

  const workbook = window.XLSX.read(await file.arrayBuffer(), {
    type: "array",
    cellDates: false,
    cellText: false,
    raw: true,
  });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("O arquivo enviado não possui abas.");
  }
  const rows = window.XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {
    header: 1,
    blankrows: true,
    defval: null,
    raw: true,
  });
  return parseSystemReport(rows);
}

document.querySelector("#cxFile").addEventListener("change", async (event) => {
  const [file] = event.target.files;
  if (!file) return;
  document.querySelector("#fileName").textContent = file.name;
  document.querySelector("#importError").hidden = true;
  document.querySelector("#importError").textContent = "";
  setStatus("", "Lendo CX");

  try {
    state.system = await parseCx(file);
    state.transactions = state.system.transactions || [];
    state.statuses = {};
    state.currentIndex = 0;
    closeCashState.draft = { fichas: [], vales: [] };
    closeCashState.editing = null;
    setStep("tickets");
    render();
  } catch (error) {
    setStatus("bad", "Erro no CX");
    document.querySelector("#fileName").textContent = "Arquivo inválido";
    document.querySelector("#importError").textContent = error.message;
    document.querySelector("#importError").hidden = false;
    event.target.value = "";
    render();
  }
});

document.querySelectorAll(".step").forEach((button) => {
  button.addEventListener("click", () => setStep(button.dataset.step));
});

document.querySelector("#previousTicket").addEventListener("click", () => {
  state.currentIndex = Math.max(0, state.currentIndex - 1);
  render();
});

document.querySelector("#nextPending").addEventListener("click", goNextPending);
document.querySelector("#finishAdjustments").addEventListener("click", () => setStep("result"));
document.querySelector("#closeCashBtn").addEventListener("click", openCloseCashModal);
document.querySelector("#closeCashModalBtn").addEventListener("click", closeCloseCashModal);
document.querySelector("#closeRefreshWorkbookBtn").addEventListener("click", async () => {
  try {
    await closeRefreshDesktopWorkbook();
    closeRender();
  } catch (error) {
    window.alert(`Não consegui atualizar a planilha.\n${error.message || error}`);
  }
});
document.querySelector("#closeChooseWorkbookFolderBtn").addEventListener("click", closeChooseWorkbookFolder);
document.querySelector("#closeSaveWorkbookBtn").addEventListener("click", saveCloseCashWorkbook);
document.querySelector("#closeClearDraftBtn").addEventListener("click", () => {
  if (!window.confirm("Limpar todos os lançamentos pendentes?")) return;
  closeCashState.draft = { fichas: [], vales: [] };
  closeCashState.editing = null;
  closeRender();
});
document.querySelector("#closeDailyForm").addEventListener("submit", (event) => {
  event.preventDefault();
  const row = readCloseDailyForm();
  addOrReplaceCloseDraftFicha(row);
  closeClearForm(event.target, "#closeDailyDate");
  closeRender();
});
document.querySelector("#closeValeForm").addEventListener("submit", (event) => {
  event.preventDefault();
  const row = {
    id: closeCashState.editing?.type === "vale" ? closeCashState.editing.id : draftId("vale"),
    date: document.querySelector("#closeValeDate").value,
    descricao: document.querySelector("#closeValeDescricao").value.trim(),
    valor: parseAmount(document.querySelector("#closeValeValor").value),
    obs: document.querySelector("#closeValeObs").value.trim(),
  };

  if (closeCashState.editing?.type === "vale") {
    closeCashState.draft.vales = closeCashState.draft.vales
      .map((item) => (item.id === closeCashState.editing.id ? row : item));
    closeCashState.editing = null;
  } else {
    closeCashState.draft.vales.push(row);
  }

  closeClearForm(event.target, "#closeValeDate");
  closeRender();
});
document.querySelectorAll("[data-close-money-field]").forEach((input) => {
  input.addEventListener("input", updateCloseDailyTotalPreview);
});
document.querySelector("#exportReportBtn").addEventListener("click", exportReport);
document.querySelector("#returnType").addEventListener("change", toggleReturnSaleFields);
document.querySelector("#addReturnEntry").addEventListener("click", addReturnEntry);
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

document.querySelectorAll("[data-adjustment]").forEach((input) => {
  input.addEventListener("input", () => {
    state.adjustments[input.dataset.adjustment] = input.value;
    renderChecks();
  });
});

document.querySelectorAll("[data-note]").forEach((input) => {
  input.addEventListener("input", () => {
    state.notes[input.dataset.note] = input.value;
  });
});

document.querySelector("#clearAllBtn").addEventListener("click", resetConference);
document.querySelector("#clearBtn").addEventListener("click", resetConference);

if (!desktopApi) {
  document.querySelector("#closeChooseWorkbookFolderBtn").disabled = true;
  document.querySelector("#closeChooseWorkbookFolderBtn").title = "Disponível apenas no aplicativo instalado";
}

toggleReturnSaleFields();
setStep("tickets");
render();
