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

function setStatus(kind, text) {
  const pill = document.querySelector("#statusPill");
  pill.className = `status-pill ${kind || ""}`.trim();
  pill.textContent = text;
}

function formatMethods(item) {
  return item.paymentMethods.length ? item.paymentMethods.join(" + ") : "Sem forma informada";
}

function ticketItems() {
  return state.transactions.filter((item) => item.needsPhysicalCheck);
}

function returnItems() {
  return state.transactions.filter((item) => item.isReturn);
}

function neutralizedItems() {
  return state.transactions.filter((item) => item.neutralized);
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

  for (const item of ticketItems()) {
    if (state.statuses[item.id]?.status !== "confirmed") continue;
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
    list.innerHTML = `<p class="empty-state">Nenhum par de entrada/saida foi anulado automaticamente.</p>`;
    return;
  }

  list.innerHTML = items
    .map(
      (item) => `<div class="return-row">
        <strong>${item.sequence}</strong>
        <span>${item.description} - anula com ${item.neutralizedWith}</span>
        <strong>${money.format(item.amount)}</strong>
      </div>`,
    )
    .join("");
}

function checkRow(label, left, right, okText = "Batendo") {
  const difference = diff(left, right);
  const ok = Math.abs(difference) < 0.01;
  const className = ok ? "ok" : "bad";
  const detail = ok ? okText : `Diferenca: ${money.format(difference)}`;
  return `<div class="check-row ${className}">
    <span>${label}</span>
    <strong>${detail}</strong>
  </div>`;
}

function renderChecks() {
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
  const checks = [];

  if (!totals) {
    document.querySelector("#checks").innerHTML =
      `<div class="check-row warn"><span>Arquivo CX</span><strong>Importe o relatorio para iniciar</strong></div>`;
    return;
  }

  checks.push(
    `<div class="check-row ${count.pending === 0 ? "ok" : "warn"}">
      <span>Tickets pendentes</span>
      <strong>${count.pending}</strong>
    </div>`,
  );
  checks.push(
    `<div class="check-row ${count.issues === 0 ? "ok" : "bad"}">
      <span>Tickets com problema</span>
      <strong>${count.issues}</strong>
    </div>`,
  );
  checks.push(checkRow("Total confirmado x Total do sistema", confirmed.totalCaixa, totals.totalCaixa));
  checks.push(checkRow("Dinheiro confirmado x Dinheiro do sistema", confirmed.dinheiro, totals.dinheiro));
  if (hasDinheiroContado) {
    checks.push(checkRow("Dinheiro contado + Pix CNPJ + Vales x Dinheiro do sistema", dinheiroFisicoAjustado, totals.dinheiro));
    checks.push(
      `<div class="check-row warn">
        <span>Pix CNPJ somado a contagem fisica</span>
        <strong>${money.format(pixCnpj)}</strong>
      </div>`,
    );
    checks.push(
      `<div class="check-row warn">
        <span>Vales recompostos na contagem fisica</span>
        <strong>${money.format(manualVales)}</strong>
      </div>`,
    );
  } else {
    checks.push(
      `<div class="check-row warn">
        <span>Dinheiro contado fisicamente</span>
        <strong>Informe para validar sobra ou falta</strong>
      </div>`,
    );
  }
  checks.push(checkRow("Cheques/Pix confirmados x Cheques/Pix do sistema", confirmed.cheques, totals.cheques));
  checks.push(checkRow("Cartoes confirmados x Debito + Credito do sistema", confirmed.cartaoDebito + confirmed.cartaoCredito, totals.cartaoDebito + totals.cartaoCredito));
  checks.push(checkRow("Pre confirmado x Pre do sistema", confirmed.pre, totals.pre));
  checks.push(
    `<div class="check-row warn">
      <span>Devolucoes informadas para revisao</span>
      <strong>${money.format(devolucoes)}</strong>
    </div>`,
  );

  const allOk = count.pending === 0 && count.issues === 0 && checks.every((row) => !row.includes("check-row bad"));
  document.querySelector("#checks").innerHTML = checks.join("");
  document.querySelector("#resultHint").textContent = allOk
    ? "Conferencia final sem divergencias"
    : "Revise os pontos destacados";
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
