const state = {
  analyses: [],
  manualAdjustments: {},
  taxRates: {},
  platform: null,
  marketplace: null,
  rows: [],
  mlOnly: [],
  filteredClosing: [],
  filteredDiffs: [],
  activeTab: "closing",
  expandedProducts: new Set(),
  charts: {},
};

const AUTH_ENDPOINT = "/api/auth";
const SESSION_ENDPOINT = "/api/session";
const SESSION_DATE_FIELDS = ["date", "updatedAt", "platformDate", "marketplaceDate"];

const BR_MONTHS = new Map([
  ["janeiro", 0],
  ["fevereiro", 1],
  ["marco", 2],
  ["março", 2],
  ["abril", 3],
  ["maio", 4],
  ["junho", 5],
  ["julho", 6],
  ["agosto", 7],
  ["setembro", 8],
  ["outubro", 9],
  ["novembro", 10],
  ["dezembro", 11],
]);

const MONTH_NAMES = [
  "Janeiro",
  "Fevereiro",
  "Março",
  "Abril",
  "Maio",
  "Junho",
  "Julho",
  "Agosto",
  "Setembro",
  "Outubro",
  "Novembro",
  "Dezembro",
];

let refs = {};
let sessionSaveTimer = 0;

document.addEventListener("DOMContentLoaded", () => {
  refs = {
    platformFile: document.querySelector("#platformFile"),
    marketplaceFile: document.querySelector("#marketplaceFile"),
    platformFileName: document.querySelector("#platformFileName"),
    marketplaceFileName: document.querySelector("#marketplaceFileName"),
    analysisName: document.querySelector("#analysisName"),
    newAnalysis: document.querySelector("#newAnalysis"),
    saveAnalysis: document.querySelector("#saveAnalysis"),
    analysisCount: document.querySelector("#analysisCount"),
    healthPanel: document.querySelector("#healthPanel"),
    loginScreen: document.querySelector("#loginScreen"),
    loginForm: document.querySelector("#loginForm"),
    loginUser: document.querySelector("#loginUser"),
    loginPassword: document.querySelector("#loginPassword"),
    loginButton: document.querySelector("#loginButton"),
    loginError: document.querySelector("#loginError"),
    accountStatus: document.querySelector("#accountStatus"),
    logoutButton: document.querySelector("#logoutButton"),
    exportXlsx: document.querySelector("#exportXlsx"),
    exportCsv: document.querySelector("#exportCsv"),
    periodPreset: document.querySelector("#periodPreset"),
    analysisFilter: document.querySelector("#analysisFilter"),
    yearField: document.querySelector("#yearField"),
    yearFilter: document.querySelector("#yearFilter"),
    monthField: document.querySelector("#monthField"),
    monthFilter: document.querySelector("#monthFilter"),
    startField: document.querySelector("#startField"),
    startDate: document.querySelector("#startDate"),
    endField: document.querySelector("#endField"),
    endDate: document.querySelector("#endDate"),
    statusFilter: document.querySelector("#statusFilter"),
    productSearch: document.querySelector("#productSearch"),
    onlyMatched: document.querySelector("#onlyMatched"),
    ignoreCanceled: document.querySelector("#ignoreCanceled"),
    kpiNet: document.querySelector("#kpiNet"),
    kpiNetSub: document.querySelector("#kpiNetSub"),
    kpiCost: document.querySelector("#kpiCost"),
    kpiCostSub: document.querySelector("#kpiCostSub"),
    kpiProfit: document.querySelector("#kpiProfit"),
    kpiProfitSub: document.querySelector("#kpiProfitSub"),
    kpiRoi: document.querySelector("#kpiRoi"),
    kpiTicket: document.querySelector("#kpiTicket"),
    dailyHint: document.querySelector("#dailyHint"),
    matchRate: document.querySelector("#matchRate"),
    matchedCount: document.querySelector("#matchedCount"),
    platformMissingCount: document.querySelector("#platformMissingCount"),
    marketMissingCount: document.querySelector("#marketMissingCount"),
    cancelledCount: document.querySelector("#cancelledCount"),
    grossProductTotal: document.querySelector("#grossProductTotal"),
    priceIncreaseTotal: document.querySelector("#priceIncreaseTotal"),
    saleFeeTotal: document.querySelector("#saleFeeTotal"),
    installmentFeeTotal: document.querySelector("#installmentFeeTotal"),
    shippingIncomeTotal: document.querySelector("#shippingIncomeTotal"),
    shippingCostsTotal: document.querySelector("#shippingCostsTotal"),
    declaredShippingTotal: document.querySelector("#declaredShippingTotal"),
    shippingDifferenceTotal: document.querySelector("#shippingDifferenceTotal"),
    discountTotal: document.querySelector("#discountTotal"),
    refundTotal: document.querySelector("#refundTotal"),
    netComponentTotal: document.querySelector("#netComponentTotal"),
    checkDiffTotal: document.querySelector("#checkDiffTotal"),
    tableCount: document.querySelector("#tableCount"),
    productSummaryCount: document.querySelector("#productSummaryCount"),
    productSummaryBody: document.querySelector("#productSummaryBody"),
    tableTitle: document.querySelector("#tableTitle"),
    tabClosing: document.querySelector("#tabClosing"),
    tabDiff: document.querySelector("#tabDiff"),
    tabTax: document.querySelector("#tabTax"),
    resultHead: document.querySelector("#resultHead"),
    resultBody: document.querySelector("#resultBody"),
  };

  refs.platformFile.addEventListener("change", (event) => handleFile("platform", event));
  refs.marketplaceFile.addEventListener("change", (event) => handleFile("marketplace", event));
  refs.newAnalysis.addEventListener("click", clearDraftAnalysis);
  refs.saveAnalysis.addEventListener("click", saveCurrentAnalysis);
  refs.loginForm.addEventListener("submit", handleLogin);
  refs.logoutButton.addEventListener("click", handleLogout);
  refs.exportXlsx.addEventListener("click", exportWorkbook);
  refs.exportCsv.addEventListener("click", exportCsv);
  refs.periodPreset.addEventListener("change", () => {
    syncFilterVisibility();
    applyFilters();
  });
  [
    refs.analysisFilter,
    refs.yearFilter,
    refs.monthFilter,
    refs.startDate,
    refs.endDate,
    refs.statusFilter,
    refs.productSearch,
    refs.onlyMatched,
    refs.ignoreCanceled,
  ].forEach((element) => {
    element.addEventListener("input", applyFilters);
    element.addEventListener("change", applyFilters);
  });
  refs.tabClosing.addEventListener("click", () => setTab("closing"));
  refs.tabDiff.addEventListener("click", () => setTab("diff"));
  refs.tabTax.addEventListener("click", () => setTab("tax"));
  refs.resultBody.addEventListener("input", handleDiffAdjustmentInput);
  refs.resultBody.addEventListener("input", handleTaxRateInput);
  refs.resultBody.addEventListener("blur", handleDiffAdjustmentBlur, true);
  refs.resultBody.addEventListener("blur", handleTaxRateBlur, true);
  refs.productSummaryBody.addEventListener("click", handleProductSummaryClick);

  refs.periodPreset.value = "all";
  refs.onlyMatched.checked = true;
  refs.ignoreCanceled.checked = true;
  syncFilterVisibility();
  setupCharts();
  updateFilterOptions();
  renderAll();
  checkAuthAndLoad();

  if (window.lucide) {
    window.lucide.createIcons();
  }
});

async function handleFile(type, event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const isPlatform = type === "platform";
  const label = isPlatform ? refs.platformFileName : refs.marketplaceFileName;
  label.textContent = file.name;
  updateHealth("Processando arquivo...");

  try {
    if (!window.XLSX) {
      throw new Error("A biblioteca de Excel não carregou. Verifique a conexão com a internet.");
    }

    const workbook = await readWorkbook(file);
    const parsed = isPlatform ? parsePlatformWorkbook(workbook, file.name) : parseMarketplaceWorkbook(workbook, file.name);
    state[type] = parsed;
    syncDraftState();
    updateHealth(state.platform && state.marketplace
      ? "Planilhas prontas para salvar"
      : "Planilha carregada");
  } catch (error) {
    console.error(error);
    state[type] = null;
    label.textContent = `Erro: ${file.name}`;
    syncDraftState();
    updateHealth(error.message, "warning");
    renderAll();
  }
}

function buildSessionPayload() {
  return {
    version: 2,
    savedAt: new Date().toISOString(),
    analyses: state.analyses.map(packAnalysis),
    manualAdjustments: packManualAdjustments(),
    taxRates: packTaxRates(),
  };
}

function packAnalysis(analysis) {
  return {
    id: analysis.id,
    name: analysis.name,
    createdAt: analysis.createdAt,
    platform: packDataset(analysis.platform),
    marketplace: packDataset(analysis.marketplace),
  };
}

function packDataset(dataset) {
  if (!dataset) return null;
  return {
    fileName: dataset.fileName,
    sheetName: dataset.sheetName,
    headerRow: dataset.headerRow,
    rows: dataset.rows || [],
  };
}

function unpackDataset(dataset, type) {
  if (!dataset) return null;
  const idField = type === "platform" ? "orderId" : "saleId";
  const rows = (dataset.rows || []).map(restoreRowDates).filter((row) => row[idField]);

  return {
    fileName: dataset.fileName || "Dados salvos",
    sheetName: dataset.sheetName || "",
    headerRow: dataset.headerRow || 1,
    rows,
    byId: new Map(rows.map((row) => [row[idField], row])),
  };
}

function unpackAnalysis(analysis, index = 0) {
  if (!analysis) return null;
  const platform = unpackDataset(analysis.platform, "platform");
  const marketplace = unpackDataset(analysis.marketplace, "marketplace");
  if (!platform || !marketplace) return null;

  return {
    id: analysis.id || createAnalysisId(),
    name: analysis.name || inferAnalysisName(platform, marketplace, index),
    createdAt: analysis.createdAt || new Date().toISOString(),
    platform,
    marketplace,
  };
}

function restoreRowDates(row) {
  const restored = { ...row };
  SESSION_DATE_FIELDS.forEach((field) => {
    if (restored[field]) {
      restored[field] = parseDate(restored[field]);
    }
  });
  return restored;
}

function packManualAdjustments() {
  return normalizeManualAdjustments(state.manualAdjustments);
}

function normalizeManualAdjustments(source = {}) {
  return Object.entries(source || {}).reduce((adjustments, [key, value]) => {
    const item = {};
    ["netReceived", "cost"].forEach((field) => {
      const number = Number(value?.[field]);
      if (Number.isFinite(number)) item[field] = round2(number);
    });

    if (Object.keys(item).length) adjustments[key] = item;
    return adjustments;
  }, {});
}

function packTaxRates() {
  return normalizeTaxRates(state.taxRates);
}

function normalizeTaxRates(source = {}, legacyRate = null) {
  const rates = Object.entries(source || {}).reduce((items, [key, value]) => {
    const rate = normalizeTaxRate(value);
    if (key === "default") {
      if (rate > 0) items[key] = rate;
    } else {
      items[key] = rate;
    }
    return items;
  }, {});

  const fallback = normalizeTaxRate(legacyRate);
  if (fallback > 0 && !Object.keys(rates).length) rates.default = fallback;
  return rates;
}

async function saveCurrentAnalysis() {
  if (!state.platform || !state.marketplace) {
    updateHealth("Selecione as duas planilhas antes de salvar", "warning");
    return;
  }

  refs.saveAnalysis.disabled = true;
  const analysis = {
    id: createAnalysisId(),
    name: text(refs.analysisName?.value) || inferAnalysisName(state.platform, state.marketplace, state.analyses.length),
    createdAt: new Date().toISOString(),
    platform: state.platform,
    marketplace: state.marketplace,
  };

  state.analyses.push(analysis);
  clearDraftAnalysis({ keepHealth: true });
  reconcile();
  await saveSessionToDatabase("Analise salva no banco");
}

function clearDraftAnalysis(options = {}) {
  state.platform = null;
  state.marketplace = null;
  if (refs.platformFile) refs.platformFile.value = "";
  if (refs.marketplaceFile) refs.marketplaceFile.value = "";
  if (refs.platformFileName) refs.platformFileName.textContent = "Nenhum arquivo selecionado";
  if (refs.marketplaceFileName) refs.marketplaceFileName.textContent = "Nenhum arquivo selecionado";
  if (refs.analysisName) refs.analysisName.value = "";
  syncDraftState();
  if (!options.keepHealth) updateHealth();
}

function syncDraftState() {
  const readyToSave = Boolean(state.platform && state.marketplace);
  if (refs.saveAnalysis) refs.saveAnalysis.disabled = !readyToSave;
  if (refs.analysisCount) {
    const total = state.analyses.length;
    refs.analysisCount.textContent = `${formatInteger(total)} ${total === 1 ? "analise salva" : "analises salvas"}`;
  }
}

function applySessionPayload(payload) {
  state.manualAdjustments = normalizeManualAdjustments(payload?.manualAdjustments);
  state.taxRates = normalizeTaxRates(payload?.taxRates, payload?.taxRate);

  const savedAnalyses = Array.isArray(payload?.analyses)
    ? payload.analyses
    : payload?.platform && payload?.marketplace
      ? [{
        id: createAnalysisId(),
        name: inferAnalysisName(
          unpackDataset(payload.platform, "platform"),
          unpackDataset(payload.marketplace, "marketplace"),
          0
        ),
        createdAt: payload.savedAt || new Date().toISOString(),
        platform: payload.platform,
        marketplace: payload.marketplace,
      }]
      : [];

  state.analyses = savedAnalyses
    .map((analysis, index) => unpackAnalysis(analysis, index))
    .filter(Boolean);

  clearDraftAnalysis({ keepHealth: true });
  reconcile();
  syncDraftState();
}

function hasImportedData() {
  return state.analyses.length > 0;
}

async function checkAuthAndLoad() {
  setLoggedOutUi();
  try {
    const auth = await authRequest("GET");
    setLoggedInUi(auth.user);
    await loadSessionFromDatabase();
  } catch {
    resetDashboardState();
    setLoggedOutUi();
  }
}

async function handleLogin(event) {
  event.preventDefault();
  refs.loginError.textContent = "";
  refs.loginButton.disabled = true;

  try {
    const data = await authRequest("POST", {
      username: refs.loginUser.value,
      password: refs.loginPassword.value,
    });
    refs.loginPassword.value = "";
    setLoggedInUi(data.user);
    await loadSessionFromDatabase();
  } catch (error) {
    refs.loginError.textContent = error.message || "Usuario ou senha invalidos.";
  } finally {
    refs.loginButton.disabled = false;
  }
}

async function handleLogout() {
  try {
    await fetch(`${AUTH_ENDPOINT}?action=logout`, { method: "POST" });
  } finally {
    resetDashboardState();
    setLoggedOutUi();
  }
}

async function authRequest(method, body = null) {
  const options = {
    method,
    headers: { "Content-Type": "application/json" },
    credentials: "same-origin",
  };
  let url = AUTH_ENDPOINT;
  if (body) options.body = JSON.stringify(body);

  const response = await fetch(url, options);
  const data = await response.json().catch(() => ({}));
  if (!response.ok || data.ok === false) {
    throw new Error(data.error || "Falha no acesso.");
  }
  return data;
}

function setLoggedInUi(user) {
  refs.loginScreen.classList.add("hidden");
  refs.accountStatus.textContent = user || "Conectado";
  refs.accountStatus.classList.add("connected");
  refs.logoutButton.disabled = false;
  refs.loginError.textContent = "";
}

function setLoggedOutUi() {
  refs.loginScreen.classList.remove("hidden");
  refs.accountStatus.textContent = "Desconectado";
  refs.accountStatus.classList.remove("connected");
  refs.logoutButton.disabled = true;
}

function resetDashboardState() {
  state.analyses = [];
  state.manualAdjustments = {};
  state.taxRates = {};
  state.expandedProducts.clear();
  state.platform = null;
  state.marketplace = null;
  state.rows = [];
  state.mlOnly = [];
  state.filteredClosing = [];
  state.filteredDiffs = [];
  refs.platformFile.value = "";
  refs.marketplaceFile.value = "";
  refs.platformFileName.textContent = "Nenhum arquivo selecionado";
  refs.marketplaceFileName.textContent = "Nenhum arquivo selecionado";
  refs.analysisName.value = "";
  refs.periodPreset.value = "all";
  refs.analysisFilter.value = "all";
  refs.statusFilter.value = "all";
  refs.productSearch.value = "";
  refs.onlyMatched.checked = true;
  refs.ignoreCanceled.checked = true;
  syncFilterVisibility();
  syncDraftState();
  updateFilterOptions();
  renderAll();
  updateHealth();
}

async function loadSessionFromDatabase() {
  try {
    const data = await sessionRequest("GET");
    if (data.payload) {
      applySessionPayload(data.payload);
      updateHealth("Dados carregados do banco", "ready");
    } else {
      resetDashboardState();
      updateHealth("Nenhuma planilha salva no banco", "warning");
    }
  } catch (error) {
    resetDashboardState();
    updateHealth(error.message || "Nao consegui carregar o banco", "warning");
  }
}

async function saveSessionToDatabase(successMessage = "Dados salvos no banco") {
  if (!hasImportedData()) return;

  try {
    await sessionRequest("POST", { payload: buildSessionPayload() });
    updateHealth(successMessage, "ready");
  } catch (error) {
    console.error(error);
    updateHealth(error.message || "Nao consegui salvar no banco", "warning");
  }
}

function scheduleSessionSave(successMessage = "Dados salvos no banco") {
  window.clearTimeout(sessionSaveTimer);
  sessionSaveTimer = window.setTimeout(() => {
    saveSessionToDatabase(successMessage);
  }, 700);
}

async function sessionRequest(method, body = null) {
  const options = {
    method,
    headers: { "Content-Type": "application/json" },
    credentials: "same-origin",
  };
  if (body) options.body = JSON.stringify(body);

  const response = await fetch(SESSION_ENDPOINT, options);
  const data = await response.json().catch(() => ({}));
  if (!response.ok || data.ok === false) {
    throw new Error(data.error || "Falha ao acessar os dados salvos.");
  }
  return data;
}
async function readWorkbook(file) {
  const data = await file.arrayBuffer();
  return window.XLSX.read(data, {
    type: "array",
    cellDates: true,
    cellNF: false,
    cellText: true,
  });
}

function parsePlatformWorkbook(workbook, fileName) {
  for (const sheetName of workbook.SheetNames) {
    const rows = getSheetRows(workbook, sheetName);
    const headerRow = rows.findIndex((row) => {
      const keys = row.map(compactKey);
      return keys.includes("numero do pedido no canal") && keys.includes("custo total");
    });

    if (headerRow < 0) continue;

    const headers = rows[headerRow];
    const find = makeColumnFinder(headers);
    const cols = {
      orderId: find.one(["Número do Pedido no Canal", "Numero do Pedido no Canal", "Pedido no Canal"]),
      internalOrderId: find.one(["Número do Pedido", "Numero do Pedido"]),
      customer: find.one(["Cliente"]),
      channel: find.one(["Canal"]),
      company: find.one(["Empresa"]),
      status: find.one(["Status"]),
      orderValue: find.one(["Valor Total"]),
      cost: find.required(["Custo Total"], "Custo Total"),
      platformProfit: find.one(["Lucro Total"]),
      platformNet: find.one(["Líquido Total", "Liquido Total"]),
      logistics: find.one(["Modo de Logística", "Modo de Logistica"]),
      product: find.one(["Produtos", "Produto"]),
      createdAt: find.one(["Data de Criação", "Data de Criacao"]),
      updatedAt: find.one(["Data de Atualização", "Data de Atualizacao"]),
    };

    const idColumn = cols.orderId >= 0 ? cols.orderId : cols.internalOrderId;
    if (idColumn < 0) {
      throw new Error("Não encontrei a coluna do número do pedido na planilha de pedidos.");
    }

    const parsedRows = rows
      .slice(headerRow + 1)
      .filter((row) => !rowIsEmpty(row))
      .map((row, index) => {
        const product = cell(row, cols.product);
        return {
          source: "platform",
          sourceRow: headerRow + index + 2,
          orderId: cleanId(cell(row, idColumn)),
          channelOrderId: cleanId(cell(row, cols.orderId)),
          internalOrderId: cleanId(cell(row, cols.internalOrderId)),
          customer: text(cell(row, cols.customer)),
          channel: text(cell(row, cols.channel)),
          company: text(cell(row, cols.company)),
          status: text(cell(row, cols.status)) || "Sem status",
          orderValue: parseMoney(cell(row, cols.orderValue)),
          cost: parseMoney(cell(row, cols.cost)),
          platformProfit: parseMoney(cell(row, cols.platformProfit)),
          platformNet: parseMoney(cell(row, cols.platformNet)),
          logistics: text(cell(row, cols.logistics)),
          product: text(product),
          sku: extractSku(product),
          date: parseDate(cell(row, cols.createdAt)),
          updatedAt: parseDate(cell(row, cols.updatedAt)),
        };
      })
      .filter((row) => row.orderId);

    return {
      fileName,
      sheetName,
      headerRow: headerRow + 1,
      rows: parsedRows,
      byId: new Map(parsedRows.map((row) => [row.orderId, row])),
    };
  }

  throw new Error("Não encontrei as colunas de pedidos e custo na planilha de pedidos.");
}

function parseMarketplaceWorkbook(workbook, fileName) {
  for (const sheetName of workbook.SheetNames) {
    const rows = getSheetRows(workbook, sheetName);
    const headerRow = rows.findIndex((row) => {
      const keys = row.map(compactKey);
      return keys.some((key) => key === "n de venda" || key.endsWith("de venda")) && keys.includes("total brl");
    });

    if (headerRow < 0) continue;

    const headers = rows[headerRow];
    const find = makeColumnFinder(headers);
    const cols = {
      saleId: find.required(["N.º de venda", "N de venda", "Número de venda", "Numero de venda"], "N.º de venda"),
      saleDate: find.required(["Data da venda"], "Data da venda"),
      status: find.one(["Estado"]),
      statusDescription: find.one(["Descrição do status", "Descricao do status"]),
      units: find.one(["Unidades"]),
      grossProduct: find.one(["Receita por produtos (BRL)"]),
      priceIncrease: find.one(["Receita por acréscimo no preço (pago pelo comprador)", "Receita por acrescimo no preco"]),
      installmentFee: find.one(["Taxa de parcelamento equivalente ao acréscimo", "Taxa de parcelamento equivalente ao acrescimo"]),
      saleFee: find.one(["Tarifa de venda e impostos (BRL)"]),
      shippingIncome: find.one(["Receita por envio (BRL)"]),
      shippingFee: find.one(["Tarifas de envio (BRL)"]),
      declaredShippingCost: find.one(["Custo de envio com base nas medidas e peso declarados"]),
      shippingDifference: find.one(["Custo por diferenças nas medidas e no peso do pacote", "Custo por diferencas nas medidas e no peso do pacote"]),
      discounts: find.one(["Descontos e bônus", "Descontos e bonus"]),
      refunds: find.one(["Cancelamentos e reembolsos (BRL)"]),
      netTotal: find.required(["Total (BRL)"], "Total (BRL)"),
      billingMonth: find.one(["Mês de faturamento das suas tarifas", "Mes de faturamento das suas tarifas"]),
      sku: find.one(["SKU"]),
      listingId: find.one(["# de anúncio", "# de anuncio"]),
      title: find.one(["Título do anúncio", "Titulo do anuncio"]),
      variation: find.one(["Variação", "Variacao"]),
      unitPrice: find.one(["Preço unitário de venda do anúncio (BRL)", "Preco unitario de venda do anuncio"]),
      adType: find.one(["Tipo de anúncio", "Tipo de anuncio"]),
      buyer: find.one(["Comprador"]),
      released: find.one(["Dinheiro liberado"]),
    };

    const groups = new Map();

    rows
      .slice(headerRow + 1)
      .filter((row) => !rowIsEmpty(row))
      .forEach((row, index) => {
        const saleId = cleanId(cell(row, cols.saleId));
        if (!saleId) return;

        if (!groups.has(saleId)) {
          groups.set(saleId, createMarketplaceGroup(saleId, headerRow + index + 2));
        }

        const group = groups.get(saleId);
        const saleDate = parseDate(cell(row, cols.saleDate));
        const sku = text(cell(row, cols.sku));
        const title = text(cell(row, cols.title));
        const status = text(cell(row, cols.status));

        group.rowCount += 1;
        group.firstRow = Math.min(group.firstRow, headerRow + index + 2);
        group.date = pickEarliestDate(group.date, saleDate);
        group.dateText = group.dateText || text(cell(row, cols.saleDate));
        group.statuses.add(status || "Sem status");
        group.statusDescriptions.add(text(cell(row, cols.statusDescription)));
        group.skus.add(sku);
        group.titles.add(title);
        group.variations.add(text(cell(row, cols.variation)));
        group.buyers.add(text(cell(row, cols.buyer)));
        group.billingMonths.add(text(cell(row, cols.billingMonth)));
        group.listingIds.add(text(cell(row, cols.listingId)));
        group.adTypes.add(text(cell(row, cols.adType)));
        group.released = group.released || text(cell(row, cols.released));

        group.units += parseMoney(cell(row, cols.units));
        group.grossProduct += parseMoney(cell(row, cols.grossProduct));
        group.priceIncrease += parseMoney(cell(row, cols.priceIncrease));
        group.installmentFee += parseMoney(cell(row, cols.installmentFee));
        group.saleFee += parseMoney(cell(row, cols.saleFee));
        group.shippingIncome += parseMoney(cell(row, cols.shippingIncome));
        group.shippingFee += parseMoney(cell(row, cols.shippingFee));
        group.declaredShippingCost += parseMoney(cell(row, cols.declaredShippingCost));
        group.shippingDifference += parseMoney(cell(row, cols.shippingDifference));
        group.discounts += parseMoney(cell(row, cols.discounts));
        group.refunds += parseMoney(cell(row, cols.refunds));
        group.netReceived += parseMoney(cell(row, cols.netTotal));
        group.unitPrice += parseMoney(cell(row, cols.unitPrice));
      });

    const parsedRows = [...groups.values()].map(finalizeMarketplaceGroup);

    return {
      fileName,
      sheetName,
      headerRow: headerRow + 1,
      rows: parsedRows,
      byId: new Map(parsedRows.map((row) => [row.saleId, row])),
    };
  }

  throw new Error("Não encontrei o cabeçalho de vendas do Mercado Livre.");
}

function createMarketplaceGroup(saleId, rowNumber) {
  return {
    source: "marketplace",
    saleId,
    firstRow: rowNumber,
    rowCount: 0,
    date: null,
    dateText: "",
    statuses: new Set(),
    statusDescriptions: new Set(),
    skus: new Set(),
    titles: new Set(),
    variations: new Set(),
    buyers: new Set(),
    billingMonths: new Set(),
    listingIds: new Set(),
    adTypes: new Set(),
    released: "",
    units: 0,
    grossProduct: 0,
    priceIncrease: 0,
    installmentFee: 0,
    saleFee: 0,
    shippingIncome: 0,
    shippingFee: 0,
    declaredShippingCost: 0,
    shippingDifference: 0,
    discounts: 0,
    refunds: 0,
    netReceived: 0,
    unitPrice: 0,
  };
}

function finalizeMarketplaceGroup(group) {
  const status = joinSet(group.statuses) || "Sem status";
  const sku = joinSet(group.skus, " + ");
  const title = firstSetValue(group.titles);
  const fees = group.saleFee + group.installmentFee;
  const shippingDetail = group.declaredShippingCost + group.shippingDifference;
  const shippingCosts = Math.abs(group.shippingFee) > 0.009 ? group.shippingFee : shippingDetail;
  const shippingNet = group.shippingIncome + shippingCosts;
  const componentNet =
    group.grossProduct +
    group.priceIncrease +
    group.installmentFee +
    group.saleFee +
    group.shippingIncome +
    shippingCosts +
    group.discounts +
    group.refunds;

  return {
    source: "marketplace",
    saleId: group.saleId,
    firstRow: group.firstRow,
    rowCount: group.rowCount,
    date: group.date,
    dateText: group.dateText,
    status,
    statusDescription: joinSet(group.statusDescriptions),
    sku,
    title,
    variation: joinSet(group.variations),
    buyer: joinSet(group.buyers),
    billingMonth: joinSet(group.billingMonths),
    listingId: joinSet(group.listingIds),
    adType: joinSet(group.adTypes),
    released: group.released,
    units: group.units,
    grossProduct: group.grossProduct,
    priceIncrease: group.priceIncrease,
    grossRevenue: group.grossProduct + group.priceIncrease,
    fees,
    saleFee: group.saleFee,
    installmentFee: group.installmentFee,
    shippingIncome: group.shippingIncome,
    shippingFee: group.shippingFee,
    declaredShippingCost: group.declaredShippingCost,
    shippingDifference: group.shippingDifference,
    shippingDetail,
    shippingCosts,
    shippingNet,
    discounts: group.discounts,
    refunds: group.refunds,
    netReceived: group.netReceived,
    componentNet,
    checkDiff: group.netReceived - componentNet,
    unitPrice: group.unitPrice,
  };
}

function reconcile() {
  state.rows = [];
  state.mlOnly = [];

  state.analyses.forEach((analysis) => {
    const reconciled = reconcileAnalysis(analysis);
    state.rows.push(...reconciled.rows);
    state.mlOnly.push(...reconciled.mlOnly);
  });

  updateFilterOptions();
  applyFilters();
  updateHealth();
}

function reconcileAnalysis(analysis) {
  const mlById = analysis.marketplace.byId;
  const platformIds = new Set(analysis.platform.rows.map((row) => row.orderId));
  const analysisFields = {
    analysisId: analysis.id,
    analysisName: analysis.name,
  };

  const rows = analysis.platform.rows.map((platformRow) => {
    const marketplaceRow = mlById.get(platformRow.orderId);
    const matched = Boolean(marketplaceRow);
    const netReceived = marketplaceRow?.netReceived ?? 0;
    const cost = platformRow.cost;
    const profit = netReceived - cost;
    const product = platformRow.product || marketplaceRow?.title || "";
    const sku = platformRow.sku || marketplaceRow?.sku || "";
    const status = marketplaceRow?.status || platformRow.status || "Sem status";
    const statusDescription = marketplaceRow?.statusDescription || "";

    return applyManualAdjustment({
      ...analysisFields,
      adjustmentKey: manualAdjustmentKey(analysis.id, "platform", platformRow.orderId),
      source: "platform",
      matched,
      missingType: matched ? "" : "Pedido sem Mercado Livre",
      orderId: platformRow.orderId,
      platformOrderId: platformRow.internalOrderId,
      marketplaceOrderId: marketplaceRow?.saleId || "",
      date: marketplaceRow?.date || platformRow.date,
      platformDate: platformRow.date,
      marketplaceDate: marketplaceRow?.date || null,
      status,
      statusDescription,
      platformStatus: platformRow.status,
      marketplaceStatus: marketplaceRow?.status || "",
      product,
      sku,
      customer: platformRow.customer,
      buyer: marketplaceRow?.buyer || "",
      channel: platformRow.channel,
      logistics: platformRow.logistics,
      units: marketplaceRow?.units ?? 0,
      orderValue: platformRow.orderValue,
      baseNetReceived: netReceived,
      baseCost: cost,
      cost,
      netReceived,
      grossRevenue: marketplaceRow?.grossRevenue ?? 0,
      grossProduct: marketplaceRow?.grossProduct ?? 0,
      priceIncrease: marketplaceRow?.priceIncrease ?? 0,
      fees: marketplaceRow?.fees ?? 0,
      saleFee: marketplaceRow?.saleFee ?? 0,
      installmentFee: marketplaceRow?.installmentFee ?? 0,
      shippingIncome: marketplaceRow?.shippingIncome ?? 0,
      shippingFee: marketplaceRow?.shippingFee ?? 0,
      declaredShippingCost: marketplaceRow?.declaredShippingCost ?? 0,
      shippingDifference: marketplaceRow?.shippingDifference ?? 0,
      shippingDetail: marketplaceRow?.shippingDetail ?? 0,
      shippingCosts: marketplaceRow?.shippingCosts ?? 0,
      shippingNet: marketplaceRow?.shippingNet ?? 0,
      discounts: marketplaceRow?.discounts ?? 0,
      refunds: marketplaceRow?.refunds ?? 0,
      componentNet: marketplaceRow?.componentNet ?? 0,
      checkDiff: marketplaceRow?.checkDiff ?? 0,
      profit,
      margin: netReceived ? profit / netReceived : 0,
      roi: cost ? profit / cost : 0,
      sourceRow: platformRow.sourceRow,
      marketplaceRow: marketplaceRow?.firstRow || "",
    });
  });

  const mlOnly = analysis.marketplace.rows
    .filter((row) => !platformIds.has(row.saleId))
    .map((row) => applyManualAdjustment({
      ...analysisFields,
      adjustmentKey: manualAdjustmentKey(analysis.id, "marketplace", row.saleId),
      source: "marketplace",
      matched: false,
      missingType: "Mercado Livre sem pedido",
      orderId: row.saleId,
      platformOrderId: "",
      marketplaceOrderId: row.saleId,
      date: row.date,
      platformDate: null,
      marketplaceDate: row.date,
      status: row.status || "Sem status",
      statusDescription: row.statusDescription || "",
      platformStatus: "",
      marketplaceStatus: row.status || "Sem status",
      product: row.title,
      sku: row.sku,
      customer: "",
      buyer: row.buyer,
      channel: "Mercado Livre",
      logistics: "",
      units: row.units,
      orderValue: 0,
      baseNetReceived: row.netReceived,
      baseCost: 0,
      cost: 0,
      netReceived: row.netReceived,
      grossRevenue: row.grossRevenue,
      grossProduct: row.grossProduct,
      priceIncrease: row.priceIncrease,
      fees: row.fees,
      saleFee: row.saleFee,
      installmentFee: row.installmentFee,
      shippingIncome: row.shippingIncome,
      shippingFee: row.shippingFee,
      declaredShippingCost: row.declaredShippingCost,
      shippingDifference: row.shippingDifference,
      shippingDetail: row.shippingDetail,
      shippingCosts: row.shippingCosts,
      shippingNet: row.shippingNet,
      discounts: row.discounts,
      refunds: row.refunds,
      componentNet: row.componentNet,
      checkDiff: row.checkDiff,
      profit: row.netReceived,
      margin: row.netReceived ? 1 : 0,
      roi: 0,
      sourceRow: "",
      marketplaceRow: row.firstRow,
    }));

  return { rows, mlOnly };
}

function applyFilters(options = {}) {
  syncFilterVisibility();
  const baseRows = [...state.rows, ...state.mlOnly];
  state.filteredClosing = baseRows.filter((row) => matchesFilters(row, true, true));
  state.filteredDiffs = baseRows.filter((row) => !row.matched && matchesFilters(row, false, true));
  renderAll(options);
}

function matchesFilters(row, respectOnlyMatched, respectIgnored = true) {
  if (respectOnlyMatched && refs.onlyMatched?.checked && !row.matched && !row.manualAdjusted) {
    return false;
  }

  const selectedAnalysis = refs.analysisFilter?.value || "all";
  if (selectedAnalysis !== "all" && row.analysisId !== selectedAnalysis) {
    return false;
  }

  if (respectIgnored && refs.ignoreCanceled?.checked && isCancelled(row)) {
    return false;
  }

  const preset = refs.periodPreset?.value || "all";
  if (preset === "year") {
    const selectedYear = refs.yearFilter?.value || "";
    if (selectedYear && yearKey(row.date) !== selectedYear) return false;
  }

  if (preset === "month") {
    const selectedMonth = refs.monthFilter?.value || "";
    if (selectedMonth && monthKey(row.date) !== selectedMonth) return false;
  }

  if (preset === "custom") {
    const start = refs.startDate?.value ? parseDate(`${refs.startDate.value} 00:00`) : null;
    const end = refs.endDate?.value ? parseDate(`${refs.endDate.value} 23:59`) : null;
    if (start && (!row.date || row.date < start)) return false;
    if (end && (!row.date || row.date > end)) return false;
  }

  const status = refs.statusFilter?.value || "all";
  if (status !== "all" && compactKey(row.status) !== status) return false;

  const query = normalizeText(refs.productSearch?.value || "");
  if (query) {
    const haystack = normalizeText([
      row.orderId,
      row.platformOrderId,
      row.marketplaceOrderId,
      row.product,
      row.sku,
      row.customer,
      row.buyer,
      row.status,
      row.statusDescription,
    ].join(" "));
    if (!haystack.includes(query)) return false;
  }

  return true;
}

function updateFilterOptions() {
  const rows = [...state.rows, ...state.mlOnly];
  const previousAnalysis = refs.analysisFilter?.value || "all";
  const previousYear = refs.yearFilter?.value || "";
  const previousMonth = refs.monthFilter?.value || "";
  const previousStatus = refs.statusFilter?.value || "all";

  refs.analysisFilter.innerHTML = [
    `<option value="all">Todas</option>`,
    ...state.analyses.map((analysis) => (
      `<option value="${escapeAttribute(analysis.id)}">${escapeHtml(analysis.name)}</option>`
    )),
  ].join("");
  if (previousAnalysis === "all" || state.analyses.some((analysis) => analysis.id === previousAnalysis)) {
    refs.analysisFilter.value = previousAnalysis;
  }

  const years = [...new Set(rows.map((row) => yearKey(row.date)).filter(Boolean))].sort();
  refs.yearFilter.innerHTML = years.length
    ? years.map((key) => `<option value="${key}">${key}</option>`).join("")
    : `<option value="">Sem ano</option>`;
  if (years.includes(previousYear)) refs.yearFilter.value = previousYear;

  const months = [...new Set(rows.map((row) => monthKey(row.date)).filter(Boolean))].sort();
  refs.monthFilter.innerHTML = months.length
    ? months.map((key) => `<option value="${key}">${monthLabel(key)}</option>`).join("")
    : `<option value="">Sem mês</option>`;
  if (months.includes(previousMonth)) refs.monthFilter.value = previousMonth;

  const statuses = [...new Set(rows.map((row) => row.status).filter(Boolean))]
    .map((status) => ({ label: status, value: compactKey(status) }))
    .sort((a, b) => a.label.localeCompare(b.label, "pt-BR"));
  refs.statusFilter.innerHTML = [
    `<option value="all">Todos</option>`,
    ...statuses.map((item) => `<option value="${escapeAttribute(item.value)}">${escapeHtml(item.label)}</option>`),
  ].join("");
  if (statuses.some((item) => item.value === previousStatus)) refs.statusFilter.value = previousStatus;

  syncDraftState();
}

function syncFilterVisibility() {
  const preset = refs.periodPreset?.value || "all";
  refs.yearField.classList.toggle("hidden", preset !== "year");
  refs.monthField.classList.toggle("hidden", preset !== "month");
  refs.startField.classList.toggle("hidden", preset !== "custom");
  refs.endField.classList.toggle("hidden", preset !== "custom");
}

function renderAll(options = {}) {
  renderKpis();
  renderReconciliation();
  renderBreakdown();
  updateCharts();
  renderProductSummary();
  if (options.renderTable !== false) renderTable();
  const hasData = state.rows.length > 0 || state.mlOnly.length > 0;
  refs.exportXlsx.disabled = !hasData;
  refs.exportCsv.disabled = !hasData;
}

function renderKpis() {
  const rows = financialRows();
  const net = sum(rows, "netReceived");
  const cost = sum(rows, "cost");
  const profit = net - cost;
  const margin = net ? profit / net : 0;
  const roi = cost ? profit / cost : 0;
  const saleCount = rows.length;
  const ticket = saleCount ? net / saleCount : 0;

  refs.kpiNet.textContent = formatMoney(net);
  refs.kpiNetSub.textContent = `${saleCount} vendas`;
  refs.kpiCost.textContent = formatMoney(cost);
  refs.kpiCostSub.textContent = financialBaseLabel();
  refs.kpiProfit.textContent = formatMoney(profit);
  refs.kpiProfit.classList.toggle("positive", profit >= 0);
  refs.kpiProfit.classList.toggle("negative", profit < 0);
  refs.kpiProfitSub.textContent = `${formatPercent(margin)} margem`;
  refs.kpiRoi.textContent = formatPercent(roi);
  refs.kpiRoi.classList.toggle("positive", roi >= 0);
  refs.kpiRoi.classList.toggle("negative", roi < 0);
  refs.kpiTicket.textContent = `Ticket médio ${formatMoney(ticket)}`;
}

function renderReconciliation() {
  const platformScopeAll = state.rows.filter((row) => matchesFilters(row, false, false));
  const mlOnlyScopeAll = state.mlOnly.filter((row) => matchesFilters(row, false, false));
  const platformScope = platformScopeAll.filter((row) => !refs.ignoreCanceled?.checked || !isCancelled(row));
  const mlOnlyScope = mlOnlyScopeAll.filter((row) => !refs.ignoreCanceled?.checked || !isCancelled(row));
  const matched = platformScope.filter((row) => row.matched).length;
  const platformMissing = platformScope.filter((row) => !row.matched).length;
  const marketMissing = mlOnlyScope.length;
  const totalPlatform = platformScope.length;
  const rate = totalPlatform ? matched / totalPlatform : 0;
  const cancelled = [...platformScopeAll, ...mlOnlyScopeAll].filter(isCancelled).length;

  refs.matchRate.textContent = `${formatPercent(rate)} localizado`;
  refs.matchedCount.textContent = formatInteger(matched);
  refs.platformMissingCount.textContent = formatInteger(platformMissing);
  refs.marketMissingCount.textContent = formatInteger(marketMissing);
  refs.cancelledCount.textContent = formatInteger(cancelled);
}

function financialBaseLabel() {
  const parts = [];
  if (refs.onlyMatched?.checked) parts.push("conciliada");
  if (refs.ignoreCanceled?.checked) parts.push("sem canceladas");
  parts.push("sem devoluções");
  return parts.length ? `base ${parts.join(" · ")}` : "base filtrada";
}

function renderBreakdown() {
  const rows = financialRows();
  refs.grossProductTotal.textContent = formatMoney(sum(rows, "grossProduct"));
  refs.priceIncreaseTotal.textContent = formatMoney(sum(rows, "priceIncrease"));
  refs.saleFeeTotal.textContent = formatMoney(sum(rows, "saleFee"));
  refs.installmentFeeTotal.textContent = formatMoney(sum(rows, "installmentFee"));
  refs.shippingIncomeTotal.textContent = formatMoney(sum(rows, "shippingIncome"));
  refs.shippingCostsTotal.textContent = formatMoney(sum(rows, "shippingCosts"));
  refs.declaredShippingTotal.textContent = formatMoney(sum(rows, "declaredShippingCost"));
  refs.shippingDifferenceTotal.textContent = formatMoney(sum(rows, "shippingDifference"));
  refs.discountTotal.textContent = formatMoney(sum(rows, "discounts"));
  refs.refundTotal.textContent = formatMoney(sum(rows, "refunds"));
  refs.netComponentTotal.textContent = formatMoney(sum(rows, "netReceived"));
  refs.checkDiffTotal.textContent = formatMoney(sum(rows, "checkDiff"));
}

function financialRows(rows = state.filteredClosing) {
  return rows.filter((row) => !isReturned(row));
}

function taxSummaryRows(rows = financialRows()) {
  const map = new Map();
  rows.forEach((row) => {
    if (isReturned(row)) return;

    const key = monthKey(row.date) || "sem-data";
    if (!map.has(key)) {
      map.set(key, {
        key,
        label: key === "sem-data" ? "Sem mês" : monthLabel(key),
        sales: 0,
        gross: 0,
        netReceived: 0,
        tax: 0,
        afterTax: 0,
      });
    }

    const item = map.get(key);
    item.sales += 1;
    item.gross += taxGrossValue(row);
    item.netReceived += Number(row.netReceived) || 0;
  });

  return [...map.values()]
    .map((item) => {
      const rate = taxRateForMonth(item.key);
      const tax = item.gross * (rate / 100);
      return {
        ...item,
        rate,
        tax,
        afterTax: item.gross - tax,
      };
    })
    .sort((a, b) => a.key.localeCompare(b.key));
}

function taxSummaryTotal(rows = taxSummaryRows()) {
  const total = rows.reduce((acc, row) => {
    acc.sales += row.sales;
    acc.gross += row.gross;
    acc.netReceived += row.netReceived;
    acc.tax += row.tax;
    acc.afterTax += row.afterTax;
    return acc;
  }, {
    sales: 0,
    gross: 0,
    netReceived: 0,
    tax: 0,
    afterTax: 0,
    effectiveRate: 0,
  });

  total.effectiveRate = total.gross ? (total.tax / total.gross) * 100 : 0;
  return total;
}

function taxGrossValue(row) {
  if (isReturned(row)) return 0;
  const gross = (Number(row.grossProduct) || 0)
    + (Number(row.priceIncrease) || 0)
    + (Number(row.shippingIncome) || 0);
  return gross || Number(row.netReceived) || 0;
}

function taxRateForMonth(key) {
  return normalizeTaxRate(state.taxRates[key] ?? state.taxRates.default ?? 0);
}

function handleTaxRateInput(event) {
  const input = event.target.closest("[data-tax-rate-key]");
  if (!input) return;

  const key = input.dataset.taxRateKey;
  setTaxRateForMonth(key, parseMoney(input.value));
  updateRenderedTaxAmounts(key);
  scheduleSessionSave("Alíquota salva no banco");
}

function handleTaxRateBlur(event) {
  const input = event.target.closest("[data-tax-rate-key]");
  if (!input) return;
  input.value = formatEditableMoney(taxRateForMonth(input.dataset.taxRateKey));
  updateRenderedTaxAmounts(input.dataset.taxRateKey);
}

function setTaxRateForMonth(key, value) {
  const rate = normalizeTaxRate(value);
  state.taxRates[key] = rate;
}

function updateRenderedTaxAmounts(key) {
  const rows = taxSummaryRows();
  const row = rows.find((item) => item.key === key);
  const totals = taxSummaryTotal(rows);
  const rowElement = refs.resultBody.querySelector(`[data-tax-row-key="${key}"]`);

  if (rowElement && row) {
    setTaxCell(rowElement, "tax", formatMoney(row.tax));
    setTaxCell(rowElement, "afterTax", formatMoney(row.afterTax));
  }

  const totalElement = refs.resultBody.querySelector("[data-tax-total]");
  if (totalElement) {
    setTaxCell(totalElement, "sales", formatInteger(totals.sales));
    setTaxCell(totalElement, "gross", formatMoney(totals.gross));
    setTaxCell(totalElement, "netReceived", formatMoney(totals.netReceived));
    setTaxCell(totalElement, "effectiveRate", formatTaxRate(totals.effectiveRate));
    setTaxCell(totalElement, "tax", formatMoney(totals.tax));
    setTaxCell(totalElement, "afterTax", formatMoney(totals.afterTax));
  }
}

function setTaxCell(rowElement, name, value) {
  const cell = rowElement.querySelector(`[data-tax-cell="${name}"]`);
  if (cell) cell.textContent = value;
}

function renderProductSummary() {
  if (!refs.productSummaryBody) return;

  const products = productSummaryRows(financialRows());
  const visibleProducts = products.slice(0, 60);
  refs.productSummaryCount.textContent = `${formatInteger(products.length)} ${products.length === 1 ? "produto" : "produtos"}`;

  if (!products.length) {
    refs.productSummaryBody.innerHTML = `<tr><td colspan="6" class="empty-state">Sem produtos para os filtros atuais.</td></tr>`;
    return;
  }

  refs.productSummaryBody.innerHTML = visibleProducts.map(renderProductSummaryRow).join("");
}

function renderProductSummaryRow(item) {
  const expanded = state.expandedProducts.has(item.key);
  const toggleLabel = expanded ? "Ocultar pedidos" : "Mostrar pedidos";

  return `
    <tr class="${item.profit >= 0 ? "profit-row" : "loss-row"}">
      <td>
        <strong>${escapeHtml(item.label)}</strong>
        <small>${escapeHtml(`${formatInteger(item.sales)} vendas`)}</small>
      </td>
      <td class="money">
        <div class="sales-toggle">
          <span>${formatInteger(item.sales)}</span>
          <button
            class="summary-toggle"
            type="button"
            aria-label="${escapeAttribute(`${toggleLabel} de ${item.label}`)}"
            aria-expanded="${expanded ? "true" : "false"}"
            data-product-key="${escapeAttribute(item.key)}"
          >${expanded ? "-" : "+"}</button>
        </div>
      </td>
      <td class="money">${formatMoney(item.netReceived)}</td>
      <td class="money">${formatMoney(item.cost)}</td>
      <td class="money ${item.profit >= 0 ? "positive" : "negative"}">${formatMoney(item.profit)}</td>
      <td class="money ${item.margin >= 0 ? "positive" : "negative"}">${formatPercent(item.margin)}</td>
    </tr>
    ${expanded ? renderProductOrderDetails(item) : ""}
  `;
}

function renderProductOrderDetails(item) {
  return `
    <tr class="product-detail-row">
      <td colspan="6">
        <div class="product-order-list">
          <table>
            <thead>
              <tr>
                <th>Pedido</th>
                <th>Data</th>
                <th class="money">Valor pago</th>
                <th class="money">Valor recebido</th>
                <th class="money">Lucro</th>
              </tr>
            </thead>
            <tbody>
              ${item.orders.map(renderProductOrderRow).join("")}
            </tbody>
          </table>
        </div>
      </td>
    </tr>
  `;
}

function renderProductOrderRow(row) {
  const order = row.marketplaceOrderId || row.orderId || row.platformOrderId || "-";
  const detail = row.platformOrderId && row.platformOrderId !== order ? `Pedido ${row.platformOrderId}` : "";

  return `
    <tr>
      <td>
        <strong>${escapeHtml(order)}</strong>
        <small>${escapeHtml(detail)}</small>
      </td>
      <td>${escapeHtml(formatDate(row.date))}</td>
      <td class="money">${formatMoney(row.cost)}</td>
      <td class="money">${formatMoney(row.netReceived)}</td>
      <td class="money ${row.profit >= 0 ? "positive" : "negative"}">${formatMoney(row.profit)}</td>
    </tr>
  `;
}

function handleProductSummaryClick(event) {
  const button = event.target.closest("[data-product-key]");
  if (!button) return;

  const key = button.dataset.productKey;
  if (state.expandedProducts.has(key)) {
    state.expandedProducts.delete(key);
  } else {
    state.expandedProducts.add(key);
  }

  renderProductSummary();
}

function productSummaryRows(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const key = productKey(row);
    if (!map.has(key)) {
      map.set(key, {
        key,
        label: key,
        sales: 0,
        netReceived: 0,
        cost: 0,
        profit: 0,
        margin: 0,
        orders: [],
      });
    }

    const item = map.get(key);
    item.sales += 1;
    item.netReceived += Number(row.netReceived) || 0;
    item.cost += Number(row.cost) || 0;
    item.profit += Number(row.profit) || 0;
    item.orders.push(row);
  });

  return [...map.values()]
    .map((item) => ({
      ...item,
      margin: item.netReceived ? item.profit / item.netReceived : 0,
      orders: [...item.orders].sort(sortByDateDesc),
    }))
    .sort((a, b) => b.profit - a.profit || b.netReceived - a.netReceived);
}

function renderTable() {
  if (state.activeTab === "closing") return renderClosingTable();
  if (state.activeTab === "diff") return renderDiffTable();
  return renderTaxTable();
}

function renderClosingTable() {
  const rows = [...state.filteredClosing].sort(sortByDateDesc);
  refs.tableTitle.textContent = "Fechamento por venda";
  refs.tableCount.textContent = `${formatInteger(rows.length)} registros`;
  refs.resultHead.innerHTML = `
    <tr>
      <th>Data</th>
      <th>Analise</th>
      <th>Venda</th>
      <th>Status</th>
      <th>Produto</th>
      <th class="money">Produto ML</th>
      <th class="money">Acréscimo</th>
      <th class="money">Tarifas</th>
      <th class="money">Envio líquido</th>
      <th class="money">Desc./bônus</th>
      <th class="money">Canc./reemb.</th>
      <th class="money">Recebido</th>
      <th class="money">Custo</th>
      <th class="money">Lucro</th>
      <th class="money">Margem</th>
    </tr>
  `;

  if (!rows.length) {
    refs.resultBody.innerHTML = `<tr><td colspan="15" class="empty-state">Sem registros para os filtros atuais.</td></tr>`;
    return;
  }

  refs.resultBody.innerHTML = rows.map((row) => `
    <tr>
      <td>${escapeHtml(formatDate(row.date))}</td>
      <td>${escapeHtml(row.analysisName || "-")}</td>
      <td>
        <strong>${escapeHtml(row.orderId || "-")}</strong>
        <small>${escapeHtml(row.platformOrderId ? `Pedido ${row.platformOrderId}` : row.missingType || "")}</small>
      </td>
      <td>
        <span class="status-pill ${statusClass(row)}">${escapeHtml(row.status || "-")}</span>
        <small>${escapeHtml(row.statusDescription || "")}</small>
      </td>
      <td>
        <strong>${escapeHtml(row.product || "Sem produto")}</strong>
        <small>${escapeHtml(row.sku ? `SKU ${row.sku}` : row.buyer || "")}</small>
      </td>
      <td class="money">${formatMoney(row.grossProduct)}</td>
      <td class="money">${formatMoney(row.priceIncrease)}</td>
      <td class="money">${formatMoney(row.fees)}</td>
      <td class="money">
        ${formatMoney(row.shippingNet)}
        <small>${escapeHtml(`receita ${formatMoney(row.shippingIncome)} · tarifa ${formatMoney(row.shippingCosts)}`)}</small>
      </td>
      <td class="money">${formatMoney(row.discounts)}</td>
      <td class="money">${formatMoney(row.refunds)}</td>
      <td class="money">${formatMoney(row.netReceived)}</td>
      <td class="money">${formatMoney(row.cost)}</td>
      <td class="money ${row.profit >= 0 ? "positive" : "negative"}">${formatMoney(row.profit)}</td>
      <td class="money ${row.margin >= 0 ? "positive" : "negative"}">${formatPercent(row.margin)}</td>
    </tr>
  `).join("");
}

function renderDiffTable() {
  const rows = [...state.filteredDiffs].sort(sortByDateDesc);
  refs.tableTitle.textContent = "Divergências";
  refs.tableCount.textContent = `${formatInteger(rows.length)} divergências`;
  refs.resultHead.innerHTML = `
    <tr>
      <th>Tipo</th>
      <th>Data</th>
      <th>Analise</th>
      <th>ID</th>
      <th>Status</th>
      <th>Produto</th>
      <th class="money">Recebido</th>
      <th class="money">Custo</th>
      <th>Observação</th>
    </tr>
  `;

  if (!rows.length) {
    refs.resultBody.innerHTML = `<tr><td colspan="9" class="empty-state">Sem divergencias para os filtros atuais.</td></tr>`;
    return;
  }

  refs.resultBody.innerHTML = rows.map((row) => {
    const note = row.source === "platform"
      ? "Pedido encontrado na plataforma, mas não localizado no relatório do Mercado Livre."
      : "Venda localizada no Mercado Livre, mas não encontrada na planilha de pedidos.";

    return `
      <tr>
        <td>${escapeHtml(row.missingType)}</td>
        <td>${escapeHtml(formatDate(row.date))}</td>
        <td>${escapeHtml(row.analysisName || "-")}</td>
        <td><strong>${escapeHtml(row.orderId || "-")}</strong></td>
        <td><span class="status-pill ${statusClass(row)}">${escapeHtml(row.status || "-")}</span></td>
        <td>
          <strong>${escapeHtml(row.product || "Sem produto")}</strong>
          <small>${escapeHtml(row.sku ? `SKU ${row.sku}` : "")}</small>
        </td>
        <td class="money editable-money-cell">${renderMoneyEditor(row, "netReceived")}</td>
        <td class="money editable-money-cell">${renderMoneyEditor(row, "cost")}</td>
        <td>${escapeHtml(note)}</td>
      </tr>
    `;
  }).join("");
}

function renderTaxTable() {
  const rows = taxSummaryRows();
  const totals = taxSummaryTotal(rows);
  refs.tableTitle.textContent = "Impostos";
  refs.tableCount.textContent = `${formatInteger(rows.length)} ${rows.length === 1 ? "mês" : "meses"}`;
  refs.resultHead.innerHTML = `
    <tr>
      <th>Mês</th>
      <th class="money">Vendas</th>
      <th class="money">Valor bruto</th>
      <th class="money">Recebido líquido</th>
      <th class="money">Alíquota</th>
      <th class="money">Imposto governo</th>
      <th class="money">Bruto após imposto</th>
    </tr>
  `;

  if (!rows.length) {
    refs.resultBody.innerHTML = `<tr><td colspan="7" class="empty-state">Sem valores para os filtros atuais.</td></tr>`;
    return;
  }

  refs.resultBody.innerHTML = [
    ...rows.map((row) => `
      <tr data-tax-row-key="${escapeAttribute(row.key)}">
        <td><strong>${escapeHtml(row.label)}</strong></td>
        <td class="money">${formatInteger(row.sales)}</td>
        <td class="money">${formatMoney(row.gross)}</td>
        <td class="money">${formatMoney(row.netReceived)}</td>
        <td class="money editable-tax-cell">${renderTaxRateEditor(row)}</td>
        <td class="money negative" data-tax-cell="tax">${formatMoney(row.tax)}</td>
        <td class="money" data-tax-cell="afterTax">${formatMoney(row.afterTax)}</td>
      </tr>
    `),
    `
      <tr class="total-row" data-tax-total>
        <td><strong>Total filtrado</strong></td>
        <td class="money" data-tax-cell="sales">${formatInteger(totals.sales)}</td>
        <td class="money" data-tax-cell="gross">${formatMoney(totals.gross)}</td>
        <td class="money" data-tax-cell="netReceived">${formatMoney(totals.netReceived)}</td>
        <td class="money" data-tax-cell="effectiveRate">${formatTaxRate(totals.effectiveRate)}</td>
        <td class="money negative" data-tax-cell="tax">${formatMoney(totals.tax)}</td>
        <td class="money" data-tax-cell="afterTax">${formatMoney(totals.afterTax)}</td>
      </tr>
    `,
  ].join("");
}

function renderTaxRateEditor(row) {
  return `
    <input
      class="money-editor tax-rate-editor"
      type="text"
      inputmode="decimal"
      autocomplete="off"
      aria-label="${escapeAttribute(`Alíquota de ${row.label}`)}"
      data-tax-rate-key="${escapeAttribute(row.key)}"
      value="${escapeAttribute(formatEditableMoney(row.rate))}"
    />
  `;
}

function renderMoneyEditor(row, field) {
  const adjusted = hasManualAdjustmentField(row.adjustmentKey, field);
  const label = field === "netReceived" ? "Recebido" : "Custo";
  return `
    <input
      class="money-editor ${adjusted ? "adjusted" : ""}"
      type="text"
      inputmode="decimal"
      autocomplete="off"
      aria-label="${escapeAttribute(`${label} ${row.orderId || ""}`)}"
      data-adjustment-key="${escapeAttribute(row.adjustmentKey)}"
      data-adjustment-field="${escapeAttribute(field)}"
      value="${escapeAttribute(formatEditableMoney(row[field]))}"
    />
  `;
}

function handleDiffAdjustmentInput(event) {
  const input = event.target.closest("[data-adjustment-field]");
  if (!input) return;

  const row = findRowByAdjustmentKey(input.dataset.adjustmentKey);
  if (!row) return;

  setManualAdjustment(row, input.dataset.adjustmentField, parseMoney(input.value));
  input.classList.toggle("adjusted", hasManualAdjustmentField(row.adjustmentKey, input.dataset.adjustmentField));
  applyFilters({ renderTable: false });
  scheduleSessionSave("Ajuste salvo no banco");
}

function handleDiffAdjustmentBlur(event) {
  const input = event.target.closest("[data-adjustment-field]");
  if (!input) return;
  input.value = formatEditableMoney(parseMoney(input.value));
}

function findRowByAdjustmentKey(key) {
  return [...state.rows, ...state.mlOnly].find((row) => row.adjustmentKey === key);
}

function setManualAdjustment(row, field, value) {
  if (!["netReceived", "cost"].includes(field)) return;

  const baseField = field === "netReceived" ? "baseNetReceived" : "baseCost";
  const key = row.adjustmentKey;
  const nextValue = round2(value);
  const baseValue = round2(row[baseField] || 0);
  const current = { ...(state.manualAdjustments[key] || {}) };

  if (sameMoney(nextValue, baseValue)) {
    delete current[field];
  } else {
    current[field] = nextValue;
  }

  if (Object.keys(current).length) {
    state.manualAdjustments[key] = current;
  } else {
    delete state.manualAdjustments[key];
  }

  Object.assign(row, applyManualAdjustment(row));
}

function applyManualAdjustment(row) {
  const adjustment = state.manualAdjustments[row.adjustmentKey] || {};
  const hasNetAdjustment = hasManualAdjustmentField(row.adjustmentKey, "netReceived");
  const hasCostAdjustment = hasManualAdjustmentField(row.adjustmentKey, "cost");
  const netReceived = hasNetAdjustment ? adjustment.netReceived : row.baseNetReceived ?? row.netReceived ?? 0;
  const cost = hasCostAdjustment ? adjustment.cost : row.baseCost ?? row.cost ?? 0;
  const profit = netReceived - cost;
  const componentNet = Number(row.componentNet) || 0;

  return {
    ...row,
    netReceived,
    cost,
    profit,
    margin: netReceived ? profit / netReceived : 0,
    roi: cost ? profit / cost : 0,
    checkDiff: netReceived - componentNet,
    manualAdjusted: hasNetAdjustment || hasCostAdjustment,
  };
}

function manualAdjustmentKey(analysisId, source, orderId) {
  return [analysisId, source, cleanId(orderId)].map(text).join("::");
}

function hasManualAdjustmentField(key, field) {
  const value = state.manualAdjustments[key]?.[field];
  return typeof value === "number" && Number.isFinite(value);
}

function sameMoney(left, right) {
  return Math.abs(round2(left) - round2(right)) < 0.01;
}

function normalizeTaxRate(value) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0) return 0;
  return Math.min(round2(number), 100);
}

function formatTaxRate(value) {
  return `${formatEditableMoney(normalizeTaxRate(value))}%`;
}

function formatEditableMoney(value) {
  return round2(value).toLocaleString("pt-BR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function setTab(tab) {
  state.activeTab = tab;
  refs.tabClosing.classList.toggle("active", tab === "closing");
  refs.tabDiff.classList.toggle("active", tab === "diff");
  refs.tabTax.classList.toggle("active", tab === "tax");
  renderTable();
}

function setupCharts() {
  if (!window.Chart) return;

  const chartDefaults = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        labels: {
          boxWidth: 12,
          color: "#44534a",
          font: { size: 12, family: "Inter, system-ui, sans-serif" },
        },
      },
      tooltip: {
        callbacks: {
          label(context) {
            return `${context.dataset.label}: ${formatMoney(context.raw)}`;
          },
        },
      },
    },
  };

  state.charts.daily = new Chart(document.querySelector("#dailyChart"), {
    type: "line",
    data: {
      labels: [],
      datasets: [
        {
          label: "Recebido",
          data: [],
          borderColor: "#1c6aa5",
          backgroundColor: "rgba(28, 106, 165, 0.12)",
          tension: 0.25,
          fill: true,
        },
        {
          label: "Custo",
          data: [],
          borderColor: "#a96705",
          backgroundColor: "rgba(169, 103, 5, 0.10)",
          tension: 0.25,
        },
        {
          label: "Lucro",
          data: [],
          borderColor: "#16794c",
          backgroundColor: "rgba(22, 121, 76, 0.10)",
          tension: 0.25,
        },
      ],
    },
    options: {
      ...chartDefaults,
      scales: chartScales("daily"),
    },
  });

  state.charts.status = new Chart(document.querySelector("#statusChart"), {
    type: "bar",
    data: {
      labels: [],
      datasets: [{
        label: "Lucro",
        data: [],
        backgroundColor: "#1c6aa5",
        borderRadius: 5,
      }],
    },
    options: {
      ...chartDefaults,
      indexAxis: "y",
      scales: chartScales("horizontalMoney"),
    },
  });

  state.charts.product = new Chart(document.querySelector("#productChart"), {
    type: "bar",
    data: {
      labels: [],
      datasets: [{
        label: "Lucro",
        data: [],
        backgroundColor: "#16794c",
        borderRadius: 5,
      }],
    },
    options: {
      ...chartDefaults,
      indexAxis: "y",
      scales: chartScales("horizontalMoney"),
    },
  });

  state.charts.lossProduct = new Chart(document.querySelector("#lossProductChart"), {
    type: "bar",
    data: {
      labels: [],
      datasets: [{
        label: "Lucro",
        data: [],
        backgroundColor: "#ba3b3b",
        borderRadius: 5,
      }],
    },
    options: {
      ...chartDefaults,
      indexAxis: "y",
      scales: chartScales("horizontalMoney"),
    },
  });
}

function chartScales(mode) {
  const baseGrid = { color: "rgba(102, 115, 107, 0.14)" };
  const baseTick = { color: "#66736b", maxRotation: 0, autoSkip: true };

  if (mode === "horizontalMoney") {
    return {
      x: {
        grid: baseGrid,
        ticks: {
          ...baseTick,
          callback(value) {
            return compactMoney(Number(value));
          },
        },
      },
      y: {
        grid: baseGrid,
        ticks: {
          ...baseTick,
          callback(value) {
            return this.getLabelForValue(value);
          },
        },
      },
    };
  }

  return {
    x: {
      grid: baseGrid,
      ticks: baseTick,
    },
    y: {
      grid: baseGrid,
      ticks: {
        ...baseTick,
        callback(value) {
          return compactMoney(Number(value));
        },
      },
    },
  };
}

function updateCharts() {
  if (!window.Chart || !state.charts.daily) return;

  const rows = state.filteredClosing;
  const groupByMonth = ["all", "year"].includes(refs.periodPreset?.value || "all");
  const period = aggregate(
    rows.filter((row) => row.date),
    (row) => groupByMonth ? monthKey(row.date) : dateKey(row.date),
    ["netReceived", "cost", "profit"]
  );
  const sortedPeriod = [...period.entries()].sort(([a], [b]) => a.localeCompare(b));
  state.charts.daily.data.labels = sortedPeriod.map(([key]) => groupByMonth ? monthLabel(key) : formatShortDateKey(key));
  state.charts.daily.data.datasets[0].data = sortedPeriod.map(([, item]) => round2(item.netReceived));
  state.charts.daily.data.datasets[1].data = sortedPeriod.map(([, item]) => round2(item.cost));
  state.charts.daily.data.datasets[2].data = sortedPeriod.map(([, item]) => round2(item.profit));
  state.charts.daily.data.datasets[2].borderColor = sortedPeriod.some(([, item]) => item.profit < 0) ? "#ba3b3b" : "#16794c";
  state.charts.daily.update();
  refs.dailyHint.textContent = `${formatInteger(rows.length)} vendas filtradas`;

  const byStatus = aggregate(rows, (row) => row.status || "Sem status", ["profit"]);
  const statusRows = [...byStatus.entries()]
    .sort((a, b) => b[1].profit - a[1].profit)
    .slice(0, 8);
  state.charts.status.data.labels = statusRows.map(([label]) => truncate(label, 24));
  state.charts.status.data.datasets[0].data = statusRows.map(([, item]) => round2(item.profit));
  state.charts.status.data.datasets[0].backgroundColor = statusRows.map(([, item]) => item.profit >= 0 ? "#16794c" : "#ba3b3b");
  state.charts.status.update();

  const byProduct = aggregate(rows, productKey, ["profit"]);
  const productRows = [...byProduct.entries()]
    .sort((a, b) => b[1].profit - a[1].profit)
    .slice(0, 8);
  state.charts.product.data.labels = productRows.map(([label]) => truncate(label, 28));
  state.charts.product.data.datasets[0].data = productRows.map(([, item]) => round2(item.profit));
  state.charts.product.data.datasets[0].backgroundColor = productRows.map(([, item]) => item.profit >= 0 ? "#16794c" : "#ba3b3b");
  state.charts.product.update();

  const lossRows = [...byProduct.entries()]
    .sort((a, b) => a[1].profit - b[1].profit)
    .slice(0, 8);
  state.charts.lossProduct.data.labels = lossRows.map(([label]) => truncate(label, 28));
  state.charts.lossProduct.data.datasets[0].data = lossRows.map(([, item]) => round2(item.profit));
  state.charts.lossProduct.data.datasets[0].backgroundColor = lossRows.map(([, item]) => item.profit >= 0 ? "#16794c" : "#ba3b3b");
  state.charts.lossProduct.update();
}

function exportWorkbook() {
  if (!window.XLSX) return;

  const metrics = currentMetrics();
  const taxTotals = taxSummaryTotal();
  const returnsIgnored = state.filteredClosing.filter(isReturned).length;
  const selectedAnalysis = state.analyses.find((analysis) => analysis.id === refs.analysisFilter?.value);
  const cancelledIgnored = refs.ignoreCanceled?.checked
    ? [...state.rows, ...state.mlOnly].filter((row) => isCancelled(row) && matchesFilters(row, false, false)).length
    : 0;
  const summary = [
    ["Analises salvas", state.analyses.length],
    ["Filtro de analise", selectedAnalysis?.name || "Todas"],
    ["Arquivos pedidos", state.analyses.map((analysis) => analysis.platform.fileName).join(" | ")],
    ["Arquivos Mercado Livre", state.analyses.map((analysis) => analysis.marketplace.fileName).join(" | ")],
    ["Modo do fechamento", financialBaseLabel()],
    ["Registros filtrados", state.filteredClosing.length],
    ["Registros financeiros", financialRows().length],
    ["Devolucoes retiradas dos calculos", returnsIgnored],
    ["Canceladas ignoradas", cancelledIgnored],
    ["Recebido líquido", round2(metrics.net)],
    ["Custo de produtos", round2(metrics.cost)],
    ["Lucro líquido", round2(metrics.profit)],
    ["Margem", round2(metrics.margin)],
    ["ROI", round2(metrics.roi)],
    ["Valor bruto fiscal", round2(taxTotals.gross)],
    ["Aliquota efetiva imposto", formatTaxRate(taxTotals.effectiveRate)],
    ["Imposto governo", round2(taxTotals.tax)],
    ["Pedidos sem Mercado Livre", state.filteredDiffs.filter((row) => row.source === "platform").length],
    ["Mercado Livre sem pedido", state.filteredDiffs.filter((row) => row.source === "marketplace").length],
  ];

  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(summary), "Resumo");
  window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.json_to_sheet(taxSummaryRows().map(toTaxExportRow)), "Impostos");
  window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.json_to_sheet(productSummaryRows(financialRows()).map(toProductExportRow)), "Produtos");
  window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.json_to_sheet(state.filteredClosing.map(toExportRow)), "Fechamento");
  window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.json_to_sheet(state.filteredDiffs.map(toExportRow)), "Divergencias");
  window.XLSX.writeFile(wb, `fechamento_dropshipping_${dateStamp()}.xlsx`);
}

function exportCsv() {
  if (!window.XLSX) return;
  const rows = exportRowsForActiveTab();
  const sheet = window.XLSX.utils.json_to_sheet(rows);
  const csv = "\uFEFF" + window.XLSX.utils.sheet_to_csv(sheet);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${activeExportName()}_${dateStamp()}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}

function exportRowsForActiveTab() {
  if (state.activeTab === "tax") return taxSummaryRows().map(toTaxExportRow);
  const rows = state.activeTab === "closing" ? state.filteredClosing : state.filteredDiffs;
  return rows.map(toExportRow);
}

function activeExportName() {
  if (state.activeTab === "tax") return "impostos";
  return state.activeTab === "closing" ? "fechamento" : "divergencias";
}

function toExportRow(row) {
  return {
    "Analise": row.analysisName || "",
    "Tipo": row.missingType || "Conciliado",
    "Data": formatDate(row.date),
    "Venda Mercado Livre": row.marketplaceOrderId || row.orderId,
    "Pedido Plataforma": row.platformOrderId,
    "Status": row.status,
    "Descricao Status": row.statusDescription,
    "Produto": row.product,
    "SKU": row.sku,
    "Recebido Liquido": round2(row.netReceived),
    "Receita Produtos ML": round2(row.grossProduct),
    "Acrescimo Comprador ML": round2(row.priceIncrease),
    "Tarifa Venda Impostos ML": round2(row.saleFee),
    "Taxa Parcelamento ML": round2(row.installmentFee),
    "Tarifas Totais ML": round2(row.fees),
    "Receita Envio ML": round2(row.shippingIncome),
    "Tarifas Envio ML": round2(row.shippingCosts),
    "Envio Declarado ML": round2(row.declaredShippingCost),
    "Diferenca Medida Peso ML": round2(row.shippingDifference),
    "Envio Liquido ML": round2(row.shippingNet),
    "Descontos Bonus ML": round2(row.discounts),
    "Cancelamentos ML": round2(row.refunds),
    "Total Conferido Pelas Colunas": round2(row.componentNet),
    "Diferenca Conferencia": round2(row.checkDiff),
    "Custo Produto": round2(row.cost),
    "Lucro Liquido": round2(row.profit),
    "Margem": round2(row.margin),
    "ROI": round2(row.roi),
  };
}

function toProductExportRow(item) {
  return {
    "Produto": item.label,
    "Vendas": item.sales,
    "Recebido Liquido": round2(item.netReceived),
    "Custo Produto": round2(item.cost),
    "Lucro Liquido": round2(item.profit),
    "Margem": round2(item.margin),
  };
}

function toTaxExportRow(item) {
  return {
    "Mes": item.label,
    "Vendas": item.sales,
    "Valor Bruto": round2(item.gross),
    "Recebido Liquido": round2(item.netReceived),
    "Aliquota Imposto": formatTaxRate(item.rate),
    "Imposto Governo": round2(item.tax),
    "Bruto Apos Imposto": round2(item.afterTax),
  };
}

function currentMetrics() {
  const rows = financialRows();
  const net = sum(rows, "netReceived");
  const cost = sum(rows, "cost");
  const profit = net - cost;
  return {
    net,
    cost,
    profit,
    margin: net ? profit / net : 0,
    roi: cost ? profit / cost : 0,
  };
}

function updateHealth(message = "", forcedClass = "") {
  const panel = refs.healthPanel;
  if (!panel) return;

  panel.classList.remove("ready", "warning");
  if (forcedClass) panel.classList.add(forcedClass);

  if (message) {
    panel.innerHTML = `<strong>${escapeHtml(message)}</strong><span>${escapeHtml(statusLine())}</span>`;
    return;
  }

  if (state.rows.length || state.mlOnly.length) {
    const matched = state.rows.filter((row) => row.matched).length;
    const missingPlatform = state.rows.filter((row) => !row.matched).length;
    const missingMarketplace = state.mlOnly.length;
    const cancelled = [...state.rows, ...state.mlOnly].filter(isCancelled).length;
    panel.classList.add(missingPlatform || missingMarketplace ? "warning" : "ready");
    panel.innerHTML = `
      <strong>${formatInteger(matched)} vendas conciliadas</strong>
      <span>${formatInteger(missingPlatform)} pedidos sem ML · ${formatInteger(missingMarketplace)} vendas sem pedido · ${formatInteger(cancelled)} canceladas</span>
    `;
    return;
  }

  if (state.platform || state.marketplace) panel.classList.add("warning");
  panel.innerHTML = `<strong>Aguardando importacao</strong><span>${escapeHtml(statusLine())}</span>`;
}

function statusLine() {
  const saved = `${formatInteger(state.analyses.length)} ${state.analyses.length === 1 ? "analise salva" : "analises salvas"}`;
  const left = state.platform ? `Pedidos prontos: ${formatInteger(state.platform.rows.length)}` : "Pedidos pendente";
  const right = state.marketplace ? `Mercado Livre pronto: ${formatInteger(state.marketplace.rows.length)}` : "Mercado Livre pendente";
  return `${saved} | ${left} | ${right}`;
}

function createAnalysisId() {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID();
  return `analysis_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function inferAnalysisName(platform, marketplace, index = 0) {
  const rows = [...(platform?.rows || []), ...(marketplace?.rows || [])];
  const months = [...new Set(rows.map((row) => monthKey(row.date)).filter(Boolean))].sort();

  if (months.length === 1) return monthLabel(months[0]);
  if (months.length > 1) return `${monthLabel(months[0])} a ${monthLabel(months.at(-1))}`;

  const fileName = platform?.fileName || marketplace?.fileName;
  return fileName ? stripExtension(fileName) : `Analise ${index + 1}`;
}

function stripExtension(fileName) {
  return text(fileName).replace(/\.[a-z0-9]+$/i, "");
}

function getSheetRows(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  return window.XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    blankrows: false,
    raw: false,
  }).map((row) => row.map((value) => (typeof value === "string" ? value.trim() : value)));
}

function makeColumnFinder(headers) {
  const keys = headers.map(compactKey);

  function one(candidates) {
    for (const candidate of candidates) {
      const target = compactKey(candidate);
      const exact = keys.findIndex((key) => key === target);
      if (exact >= 0) return exact;
    }

    for (const candidate of candidates) {
      const target = compactKey(candidate);
      if (!target) continue;
      const contained = keys.findIndex((key) => key.includes(target) || target.includes(key));
      if (contained >= 0) return contained;
    }

    return -1;
  }

  return {
    one,
    required(candidates, label) {
      const index = one(candidates);
      if (index < 0) throw new Error(`Coluna obrigatória não encontrada: ${label}.`);
      return index;
    },
  };
}

function rowIsEmpty(row) {
  return !row.some((value) => text(value));
}

function cell(row, index) {
  if (index == null || index < 0) return "";
  return row[index] ?? "";
}

function text(value) {
  if (value == null) return "";
  return String(value).trim();
}

function normalizeText(value) {
  return text(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[º°ª]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function compactKey(value) {
  return normalizeText(value).replace(/[^a-z0-9]+/g, " ").trim();
}

function cleanId(value) {
  return text(value).replace(/\s+/g, "").replace(/\.0$/, "");
}

function extractSku(value) {
  const match = text(value).match(/SKU:\s*([A-Z0-9._-]+)/i);
  return match ? match[1] : "";
}

function parseMoney(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  let raw = text(value);
  if (!raw || raw === "-") return 0;

  raw = raw
    .replace(/\u00a0/g, "")
    .replace(/R\$/gi, "")
    .replace(/\s/g, "");

  const negative = raw.startsWith("(") && raw.endsWith(")");
  raw = raw.replace(/[()]/g, "");

  const lastComma = raw.lastIndexOf(",");
  const lastDot = raw.lastIndexOf(".");
  if (lastComma >= 0 && lastDot >= 0) {
    raw = lastComma > lastDot ? raw.replace(/\./g, "").replace(",", ".") : raw.replace(/,/g, "");
  } else if (lastComma >= 0) {
    raw = raw.replace(",", ".");
  } else {
    const dotCount = (raw.match(/\./g) || []).length;
    if (dotCount > 1) {
      const parts = raw.split(".");
      raw = `${parts.slice(0, -1).join("")}.${parts.at(-1)}`;
    }
  }

  const number = Number.parseFloat(raw.replace(/[^\d.-]/g, ""));
  if (!Number.isFinite(number)) return 0;
  return negative ? -number : number;
}

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && Number.isFinite(value)) return excelSerialDate(value);

  const raw = text(value);
  if (!raw) return null;

  const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (iso) {
    return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]), Number(iso[4] || 0), Number(iso[5] || 0));
  }

  const br = raw.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (br) {
    const year = br[3].length === 2 ? 2000 + Number(br[3]) : Number(br[3]);
    return new Date(year, Number(br[2]) - 1, Number(br[1]), Number(br[4] || 0), Number(br[5] || 0));
  }

  const normalized = normalizeText(raw);
  const longBr = normalized.match(/(\d{1,2}) de ([a-z]+) de (\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (longBr) {
    const month = BR_MONTHS.get(longBr[2]);
    if (month != null) {
      return new Date(Number(longBr[3]), month, Number(longBr[1]), Number(longBr[4] || 0), Number(longBr[5] || 0));
    }
  }

  const fallback = new Date(raw);
  return Number.isNaN(fallback.getTime()) ? null : fallback;
}

function excelSerialDate(serial) {
  const wholeDays = Math.floor(serial - 25569);
  const seconds = Math.round((serial - Math.floor(serial)) * 86400);
  const date = new Date(wholeDays * 86400 * 1000);
  date.setSeconds(date.getSeconds() + seconds);
  return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), date.getUTCHours(), date.getUTCMinutes());
}

function pickEarliestDate(current, next) {
  if (!current) return next || null;
  if (!next) return current;
  return next < current ? next : current;
}

function dateKey(date) {
  if (!date) return "";
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function monthKey(date) {
  if (!date) return "";
  return dateKey(date).slice(0, 7);
}

function yearKey(date) {
  if (!date) return "";
  return dateKey(date).slice(0, 4);
}

function monthLabel(key) {
  const [year, month] = key.split("-").map(Number);
  return `${MONTH_NAMES[month - 1]} ${year}`;
}

function formatShortDateKey(key) {
  const [, month, day] = key.split("-");
  return `${day}/${month}`;
}

function formatDate(date) {
  if (!date) return "-";
  return new Intl.DateTimeFormat("pt-BR").format(date);
}

function formatMoney(value) {
  return new Intl.NumberFormat("pt-BR", {
    style: "currency",
    currency: "BRL",
  }).format(value || 0);
}

function compactMoney(value) {
  return new Intl.NumberFormat("pt-BR", {
    notation: "compact",
    maximumFractionDigits: 1,
  }).format(value || 0);
}

function formatPercent(value) {
  return new Intl.NumberFormat("pt-BR", {
    style: "percent",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value || 0);
}

function formatInteger(value) {
  return new Intl.NumberFormat("pt-BR").format(value || 0);
}

function sum(rows, key) {
  return rows.reduce((total, row) => total + (Number(row[key]) || 0), 0);
}

function aggregate(rows, keyFn, fields) {
  const map = new Map();
  rows.forEach((row) => {
    const key = keyFn(row) || "Sem informação";
    if (!map.has(key)) {
      map.set(key, Object.fromEntries(fields.map((field) => [field, 0])));
    }
    const item = map.get(key);
    fields.forEach((field) => {
      item[field] += Number(row[field]) || 0;
    });
  });
  return map;
}

function productKey(row) {
  const sku = row.sku ? `${row.sku} · ` : "";
  return `${sku}${row.product || "Sem produto"}`;
}

function joinSet(set, separator = ", ") {
  return [...set].filter(Boolean).join(separator);
}

function firstSetValue(set) {
  return [...set].find(Boolean) || "";
}

function sortByDateDesc(a, b) {
  const da = a.date ? a.date.getTime() : 0;
  const db = b.date ? b.date.getTime() : 0;
  return db - da;
}

function isCancelled(row) {
  const haystack = normalizeText([
    row.status,
    row.statusDescription,
    row.marketplaceStatus,
    row.platformStatus,
  ].join(" "));
  return haystack.includes("cancel") || haystack.includes("nao despache");
}

function isReturned(row) {
  const haystack = normalizeText([
    row.status,
    row.statusDescription,
    row.marketplaceStatus,
    row.platformStatus,
  ].join(" "));
  return haystack.includes("devolu");
}

function statusClass(row) {
  if (!row.matched && row.source === "platform") return "bad";
  if (isCancelled(row)) return "bad";
  if (isReturned(row)) return "bad";
  if (row.matched) return "ok";
  return "";
}

function round2(value) {
  return Math.round((Number(value) || 0) * 100) / 100;
}

function truncate(value, maxLength) {
  const str = text(value);
  return str.length > maxLength ? `${str.slice(0, maxLength - 1)}…` : str;
}

function dateStamp() {
  const now = new Date();
  return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}_${String(now.getHours()).padStart(2, "0")}${String(now.getMinutes()).padStart(2, "0")}`;
}

function escapeHtml(value) {
  return text(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replace(/`/g, "&#096;");
}
