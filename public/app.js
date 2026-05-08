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
  notes: {},
  step: "tickets",
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

function itemStatus(item) {
  if (item.autoConfirmed) return "Automatico";
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
        ["Debito", totals.cartaoDebito],
        ["Credito", totals.cartaoCredito],
        ["Pre", totals.pre],
        ["Vales", totals.vales],
        ["Parcelado", totals.parcelado],
        ["Total", totals.totalCaixa],
      ]
    : [
        ["Dinheiro", 0],
        ["Cheques/Pix", 0],
        ["Debito", 0],
        ["Credito", 0],
        ["Pre", 0],
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
    card.innerHTML = `<p class="empty-state">Importe o arquivo CX para montar a fila de conferencia.</p>`;
    return;
  }

  const status = state.statuses[item.id]?.status || "pending";
  const note = state.statuses[item.id]?.note || "";
  const statusText =
    status === "confirmed" ? "Confirmado" : status === "issue" ? "Com problema" : "Pendente";

  card.innerHTML = `
    <div class="ticket-main">
      <div>
        <span class="ticket-sequence">${item.sequence || "Sem seq."}</span>
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
      <div><span>Funcionario</span><strong>${item.employee || "-"}</strong></div>
      <div><span>Status</span><strong>${statusText}</strong></div>
    </div>

    ${note ? `<p class="empty-state">Obs.: ${note}</p>` : ""}

    <div class="ticket-actions">
      <button id="confirmTicket" type="button">Confirmar ticket fisico</button>
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
    list.innerHTML = `<p class="empty-state">Nenhuma devolucao foi identificada automaticamente no CX.</p>`;
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

function buildCheckData() {
  const totals = state.system?.totals;
  const count = counts();
  const confirmed = sumConfirmedPayments();
  const manualVales = adjustment("valesManuais");
  const dinheiroContado = adjustment("dinheiroContado");
  const pixCnpj = adjustment("pixCnpj");
  const dinheiroFisicoAjustado = dinheiroContado + pixCnpj + manualVales;
  const hasDinheiroContado = String(state.adjustments.dinheiroContado || "").trim() !== "";
  const devolucoes =
    adjustment("devolucaoBaixa") +
    adjustment("devolucaoSemEntrada") +
    adjustment("devolucaoNovaVenda");

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
      "Cartoes confirmados x Debito + Credito do sistema",
      confirmed.cartaoDebito + confirmed.cartaoCredito,
      totals.cartaoDebito + totals.cartaoCredito,
    ],
    ["Pre confirmado x Pre do sistema", confirmed.pre, totals.pre],
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
      detail: ok ? "Batendo" : `Diferenca: ${money.format(difference)}`,
    });
  }

  if (hasDinheiroContado) {
    rows.push(
      {
        label: "Pix CNPJ somado a contagem fisica",
        left: pixCnpj,
        right: "",
        difference: "",
        status: "Informativo",
        level: "warn",
        detail: money.format(pixCnpj),
      },
      {
        label: "Vales recompostos na contagem fisica",
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
    label: "Devolucoes informadas para revisao",
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
      `<div class="check-row warn"><span>Arquivo CX</span><strong>Importe o relatorio para iniciar</strong></div>`;
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
    ? "Conferencia final sem divergencias"
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
    "Sequencia",
    "Descricao",
    "Detalhes",
    "Funcionario",
    "Tipo",
    "Formas",
    "Status",
    "Observacao",
    "Valor movimento",
    "Total caixa",
    "Dinheiro",
    "Cheques/Pix",
    "Pre",
    "Debito",
    "Credito",
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
  const devolucaoBaixa = adjustment("devolucaoBaixa");
  const devolucaoSemEntrada = adjustment("devolucaoSemEntrada");
  const devolucaoNovaVenda = adjustment("devolucaoNovaVenda");
  const diferencaNovaVenda = adjustment("diferencaNovaVenda");

  return [
    ["Total de vales", manualVales, state.notes.vales || ""],
    ["Dinheiro contado", dinheiroContado, state.notes.dinheiro || ""],
    ["Pix CNPJ", pixCnpj, "Somado a contagem fisica do dinheiro"],
    ["Devolucao A - baixa na conta", devolucaoBaixa, "Sai e entra novamente no sistema"],
    ["Devolucao B - sem entrada", devolucaoSemEntrada, "Sai do caixa e deve ser revisada"],
    ["Devolucao C - nova venda", devolucaoNovaVenda, "Valor devolvido usado em nova venda"],
    ["Diferenca paga pelo cliente", diferencaNovaVenda, "Complemento recebido na nova venda"],
    ["Total devolucoes informadas", devolucaoBaixa + devolucaoSemEntrada + devolucaoNovaVenda, ""],
  ];
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
    window.alert("Importe o arquivo CX antes de exportar o relatorio.");
    return;
  }
  if (!window.XLSX) {
    window.alert("Exportador de Excel nao carregou. Atualize a pagina e tente novamente.");
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
  const finalStatus =
    count.pending === 0 && count.issues === 0 && checkRows.every((row) => row.level !== "bad")
      ? "Conferencia sem divergencias"
      : "Conferencia com pontos para revisar";

  appendSheet(workbook, "Resumo", [
    ["Relatorio de conferencia de caixa"],
    ["Gerado em", stamp.toLocaleString("pt-BR")],
    ["Empresa", metadata.empresa || ""],
    ["Caixa", metadata.caixa || ""],
    ["Linhas lidas", metadata.linhasLidas || ""],
    ["Status final", finalStatus],
    [],
    ["Indicador", "Valor"],
    ["Tickets fisicos", count.total],
    ["Tickets confirmados", count.confirmed],
    ["Tickets pendentes", count.pending],
    ["Tickets com problema", count.issues],
    ["Movimentos automaticos", automaticItems.length],
    ["Contas de clientes recebidas", customerRows.length],
    ["Devolucoes encontradas no CX", returnItems().length],
    [],
    ["Caixa conferido", "Sistema", "Confirmado/Calculado", "Diferenca"],
    ["Total caixa", totals.totalCaixa, confirmed.totalCaixa, diff(confirmed.totalCaixa, totals.totalCaixa)],
    ["Dinheiro", totals.dinheiro, confirmed.dinheiro, diff(confirmed.dinheiro, totals.dinheiro)],
    ["Cheques/Pix", totals.cheques, confirmed.cheques, diff(confirmed.cheques, totals.cheques)],
    ["Debito", totals.cartaoDebito, confirmed.cartaoDebito, diff(confirmed.cartaoDebito, totals.cartaoDebito)],
    ["Credito", totals.cartaoCredito, confirmed.cartaoCredito, diff(confirmed.cartaoCredito, totals.cartaoCredito)],
    ["Pre", totals.pre, confirmed.pre, diff(confirmed.pre, totals.pre)],
    [],
    ["Valores fora da conferencia fisica", "Valor", "Observacao"],
    ["Vales do sistema", totals.vales, "Informativo"],
    ["Parcelado do sistema", totals.parcelado, "Nao entra como ticket fisico"],
    ["Contas de clientes no sistema", totals.contaCliente, "Recebimentos/baixas identificados no CX"],
    ["Devolucoes encontradas no CX", rounded(returnTotal), "Detalhes na aba Devolucoes"],
    ["Movimentos automaticos", rounded(automaticTotal), "Detalhes na aba Movimentos automaticos"],
    [],
    ["Ajustes manuais informados", "Valor", "Observacao"],
    ...manualAdjustments,
  ]);

  appendSheet(workbook, "Checks", [
    ["Check", "Valor conferido", "Valor sistema", "Diferenca", "Status", "Detalhe"],
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
    ["Check", "Valor conferido", "Valor sistema", "Diferenca", "Status", "Detalhe"],
    ...checkRows
      .filter((row) => row.level !== "ok")
      .map((row) => [row.label, row.left, row.right, row.difference, row.status, row.detail]),
    [],
    ["Tickets pendentes ou com problema"],
    ["Linha CX", "Sequencia", "Descricao", "Detalhes", "Funcionario", "Formas", "Status", "Observacao", "Valor"],
    ...issues,
  ]);

  appendSheet(workbook, "Tickets", [
    transactionHeader(),
    ...transactionRows(ticketItems()),
  ]);

  appendSheet(workbook, "Movimentos automaticos", [
    transactionHeader(),
    ...transactionRows(neutralizedItems()),
  ]);

  appendSheet(workbook, "Devolucoes", [
    transactionHeader(),
    ...transactionRows(returnItems()),
  ]);

  appendSheet(workbook, "Contas clientes", [
    [
      "Linha CX",
      "Sequencia",
      "Cliente",
      "Descricao",
      "Detalhes",
      "Funcionario",
      "Tipo",
      "Status",
      "Valor movimento",
      "Conta cliente",
      "Parcelado",
      "Dinheiro",
      "Cheques/Pix",
      "Debito",
      "Credito",
      "Formas",
    ],
    ...customerRows,
  ]);

  appendSheet(workbook, "Todos movimentos CX", [
    transactionHeader(),
    ...transactionRows(state.transactions),
  ]);

  appendSheet(workbook, "Ajustes manuais", [
    ["Campo", "Valor", "Observacao"],
    ...manualAdjustments,
  ]);

  window.XLSX.writeFile(workbook, `relatorio-conferencia-caixa-${fileStamp}.xlsx`);
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

function render() {
  renderSystem();
  renderSummary();
  renderTicketCard();
  renderTicketList();
  renderReturns();
  renderNeutralized();
  renderChecks();
}

async function parseCx(file) {
  if (!window.XLSX) {
    throw new Error("Leitor de Excel nao carregou. Atualize a pagina e tente novamente.");
  }

  const workbook = window.XLSX.read(await file.arrayBuffer(), {
    type: "array",
    cellDates: false,
    cellText: false,
    raw: true,
  });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("O arquivo enviado nao possui abas.");
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
    setStep("tickets");
    render();
  } catch (error) {
    setStatus("bad", "Erro no CX");
    document.querySelector("#fileName").textContent = "Arquivo invalido";
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
document.querySelector("#exportReportBtn").addEventListener("click", exportReport);

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

document.querySelector("#clearBtn").addEventListener("click", () => {
  state.statuses = {};
  state.adjustments = {};
  document.querySelectorAll("[data-adjustment]").forEach((input) => {
    input.value = "";
  });
  state.currentIndex = 0;
  setStep("tickets");
  render();
});

setStep("tickets");
render();
