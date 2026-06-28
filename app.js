const state = {
  platform: null,
  marketplace: null,
  rows: [],
  mlOnly: [],
  filteredClosing: [],
  filteredDiffs: [],
  activeTab: "closing",
  charts: {},
};

const SESSION_STORAGE_KEY = "mlDropshipDashboardSession:v1";
const CLOUD_KEY_STORAGE_KEY = "mlDropshipDashboardCloudKey:v1";
const CLOUD_SESSION_ENDPOINT = "/api/session";
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

document.addEventListener("DOMContentLoaded", () => {
  refs = {
    platformFile: document.querySelector("#platformFile"),
    marketplaceFile: document.querySelector("#marketplaceFile"),
    platformFileName: document.querySelector("#platformFileName"),
    marketplaceFileName: document.querySelector("#marketplaceFileName"),
    healthPanel: document.querySelector("#healthPanel"),
    saveBackup: document.querySelector("#saveBackup"),
    loadBackup: document.querySelector("#loadBackup"),
    clearSaved: document.querySelector("#clearSaved"),
    backupFile: document.querySelector("#backupFile"),
    cloudKey: document.querySelector("#cloudKey"),
    cloudSave: document.querySelector("#cloudSave"),
    cloudLoad: document.querySelector("#cloudLoad"),
    cloudDelete: document.querySelector("#cloudDelete"),
    cloudStatus: document.querySelector("#cloudStatus"),
    exportXlsx: document.querySelector("#exportXlsx"),
    exportCsv: document.querySelector("#exportCsv"),
    periodPreset: document.querySelector("#periodPreset"),
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
    tabClosing: document.querySelector("#tabClosing"),
    tabDiff: document.querySelector("#tabDiff"),
    resultHead: document.querySelector("#resultHead"),
    resultBody: document.querySelector("#resultBody"),
  };

  refs.platformFile.addEventListener("change", (event) => handleFile("platform", event));
  refs.marketplaceFile.addEventListener("change", (event) => handleFile("marketplace", event));
  refs.saveBackup.addEventListener("click", exportSessionBackup);
  refs.loadBackup.addEventListener("click", () => refs.backupFile.click());
  refs.clearSaved.addEventListener("click", clearSavedSession);
  refs.backupFile.addEventListener("change", importSessionBackup);
  refs.cloudKey.addEventListener("input", handleCloudKeyInput);
  refs.cloudSave.addEventListener("click", saveSessionToCloud);
  refs.cloudLoad.addEventListener("click", loadSessionFromCloud);
  refs.cloudDelete.addEventListener("click", deleteCloudSession);
  refs.exportXlsx.addEventListener("click", exportWorkbook);
  refs.exportCsv.addEventListener("click", exportCsv);
  refs.periodPreset.addEventListener("change", () => {
    syncFilterVisibility();
    applyFilters();
  });
  [
    refs.monthFilter,
    refs.startDate,
    refs.endDate,
    refs.statusFilter,
    refs.productSearch,
    refs.onlyMatched,
    refs.ignoreCanceled,
  ].forEach((element) => element.addEventListener("input", applyFilters));
  refs.tabClosing.addEventListener("click", () => setTab("closing"));
  refs.tabDiff.addEventListener("click", () => setTab("diff"));

  refs.periodPreset.value = "all";
  refs.onlyMatched.checked = true;
  refs.ignoreCanceled.checked = true;
  refs.cloudKey.value = restoreCloudKey();
  syncFilterVisibility();
  setupCharts();
  const restored = restoreSavedSession();
  if (!restored) {
    updateFilterOptions();
    renderAll();
  }

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
    saveSessionToLocal();
    reconcile();
  } catch (error) {
    console.error(error);
    state[type] = null;
    label.textContent = `Erro: ${file.name}`;
    updateHealth(error.message, "warning");
    renderAll();
  }
}

function buildSessionPayload() {
  return {
    version: 1,
    savedAt: new Date().toISOString(),
    platform: packDataset(state.platform),
    marketplace: packDataset(state.marketplace),
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

function restoreRowDates(row) {
  const restored = { ...row };
  SESSION_DATE_FIELDS.forEach((field) => {
    if (restored[field]) {
      restored[field] = parseDate(restored[field]);
    }
  });
  return restored;
}

function hasImportedData() {
  return Boolean(state.platform || state.marketplace || state.rows.length || state.mlOnly.length);
}

function hasSavedSession() {
  try {
    return Boolean(window.localStorage?.getItem(SESSION_STORAGE_KEY));
  } catch {
    return false;
  }
}

function saveSessionToLocal() {
  if (!state.platform && !state.marketplace) return false;

  try {
    window.localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(buildSessionPayload()));
    return true;
  } catch (error) {
    console.warn("Não foi possível salvar a sessão localmente.", error);
    return false;
  }
}

function restoreSavedSession() {
  try {
    const raw = window.localStorage?.getItem(SESSION_STORAGE_KEY);
    if (!raw) return false;
    const payload = JSON.parse(raw);
    applySessionPayload(payload, { saveLocal: false });
    return true;
  } catch (error) {
    console.warn("Não foi possível restaurar a sessão salva.", error);
    return false;
  }
}

function applySessionPayload(payload, options = {}) {
  if (!payload || payload.version !== 1) {
    throw new Error("Arquivo de sessão inválido.");
  }

  state.platform = unpackDataset(payload.platform, "platform");
  state.marketplace = unpackDataset(payload.marketplace, "marketplace");

  refs.platformFileName.textContent = state.platform?.fileName || "Nenhum arquivo selecionado";
  refs.marketplaceFileName.textContent = state.marketplace?.fileName || "Nenhum arquivo selecionado";

  if (options.saveLocal !== false) {
    saveSessionToLocal();
  }

  if (state.platform && state.marketplace) {
    reconcile();
  } else {
    state.rows = [];
    state.mlOnly = [];
    state.filteredClosing = [];
    state.filteredDiffs = [];
    updateFilterOptions();
    renderAll();
    updateHealth("Sessão carregada parcialmente", "warning");
  }
}

function exportSessionBackup() {
  const payload = buildSessionPayload();
  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: "application/json;charset=utf-8",
  });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `backup_fechamento_dropshipping_${dateStamp()}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}

async function importSessionBackup(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    const payload = JSON.parse(await file.text());
    applySessionPayload(payload);
  } catch (error) {
    console.error(error);
    updateHealth("Não consegui carregar esse backup", "warning");
  } finally {
    event.target.value = "";
  }
}

function clearSavedSession() {
  try {
    window.localStorage?.removeItem(SESSION_STORAGE_KEY);
  } catch {
    // Ignora bloqueios do navegador; o estado da tela ainda será limpo.
  }

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
  refs.periodPreset.value = "all";
  refs.statusFilter.value = "all";
  refs.productSearch.value = "";
  refs.onlyMatched.checked = true;
  refs.ignoreCanceled.checked = true;
  syncFilterVisibility();
  updateFilterOptions();
  renderAll();
  updateHealth();
}

function restoreCloudKey() {
  try {
    return window.localStorage?.getItem(CLOUD_KEY_STORAGE_KEY) || "";
  } catch {
    return "";
  }
}

function getCloudKey() {
  return text(refs.cloudKey?.value);
}

function handleCloudKeyInput() {
  const key = getCloudKey();
  try {
    if (key) {
      window.localStorage?.setItem(CLOUD_KEY_STORAGE_KEY, key);
    } else {
      window.localStorage?.removeItem(CLOUD_KEY_STORAGE_KEY);
    }
  } catch {
    // A chave continua funcionando na sessão atual mesmo se o navegador bloquear localStorage.
  }
  updateCloudControls();
}

function updateCloudControls() {
  if (!refs.cloudKey) return;
  const hasKey = Boolean(getCloudKey());
  const hasSessionData = hasImportedData();
  refs.cloudSave.disabled = !hasKey || !hasSessionData;
  refs.cloudLoad.disabled = !hasKey;
  refs.cloudDelete.disabled = !hasKey;
}

function setCloudStatus(message, kind = "") {
  if (!refs.cloudStatus) return;
  refs.cloudStatus.textContent = message;
  refs.cloudStatus.classList.toggle("ok", kind === "ok");
  refs.cloudStatus.classList.toggle("warning", kind === "warning");
}

async function cloudRequest(method, body = null) {
  const key = getCloudKey();
  if (!key) {
    throw new Error("Digite uma chave da nuvem.");
  }

  const options = {
    method,
    headers: { "Content-Type": "application/json" },
  };
  let url = CLOUD_SESSION_ENDPOINT;

  if (method === "GET") {
    url += `?key=${encodeURIComponent(key)}`;
  } else {
    options.body = JSON.stringify({ key, ...body });
  }

  const response = await fetch(url, options);
  const data = await response.json().catch(() => ({}));
  if (!response.ok || data.ok === false) {
    throw new Error(data.error || "Falha na sincronização com a nuvem.");
  }
  return data;
}

async function saveSessionToCloud() {
  if (!hasImportedData()) {
    setCloudStatus("Importe as planilhas antes", "warning");
    return;
  }

  setCloudStatus("Salvando...");
  try {
    await cloudRequest("POST", { payload: buildSessionPayload() });
    setCloudStatus("Sessão salva na nuvem", "ok");
  } catch (error) {
    console.error(error);
    setCloudStatus(error.message, "warning");
  }
}

async function loadSessionFromCloud() {
  setCloudStatus("Carregando...");
  try {
    const data = await cloudRequest("GET");
    applySessionPayload(data.payload);
    setCloudStatus("Sessão carregada da nuvem", "ok");
  } catch (error) {
    console.error(error);
    setCloudStatus(error.message, "warning");
  }
}

async function deleteCloudSession() {
  setCloudStatus("Apagando...");
  try {
    await cloudRequest("DELETE");
    setCloudStatus("Sessão apagada da nuvem", "ok");
  } catch (error) {
    console.error(error);
    setCloudStatus(error.message, "warning");
  }
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
  if (!state.platform || !state.marketplace) {
    state.rows = [];
    state.mlOnly = [];
    state.filteredClosing = [];
    state.filteredDiffs = [];
    updateHealth();
    updateFilterOptions();
    renderAll();
    return;
  }

  const mlById = state.marketplace.byId;
  const platformIds = new Set(state.platform.rows.map((row) => row.orderId));

  state.rows = state.platform.rows.map((platformRow) => {
    const marketplaceRow = mlById.get(platformRow.orderId);
    const matched = Boolean(marketplaceRow);
    const netReceived = marketplaceRow?.netReceived ?? 0;
    const cost = platformRow.cost;
    const profit = netReceived - cost;
    const product = platformRow.product || marketplaceRow?.title || "";
    const sku = platformRow.sku || marketplaceRow?.sku || "";
    const status = marketplaceRow?.status || platformRow.status || "Sem status";
    const statusDescription = marketplaceRow?.statusDescription || "";

    return {
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
    };
  });

  state.mlOnly = state.marketplace.rows
    .filter((row) => !platformIds.has(row.saleId))
    .map((row) => ({
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

  updateFilterOptions();
  applyFilters();
  updateHealth();
}

function applyFilters() {
  syncFilterVisibility();
  const baseRows = [...state.rows, ...state.mlOnly];
  state.filteredClosing = baseRows.filter((row) => matchesFilters(row, true, true));
  state.filteredDiffs = baseRows.filter((row) => !row.matched && matchesFilters(row, false, true));
  renderAll();
}

function matchesFilters(row, respectOnlyMatched, respectIgnored = true) {
  if (respectOnlyMatched && refs.onlyMatched?.checked && !row.matched) {
    return false;
  }

  if (respectIgnored && refs.ignoreCanceled?.checked && isCancelled(row)) {
    return false;
  }

  const preset = refs.periodPreset?.value || "all";
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
  const previousMonth = refs.monthFilter?.value || "";
  const previousStatus = refs.statusFilter?.value || "all";

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
}

function syncFilterVisibility() {
  const preset = refs.periodPreset?.value || "all";
  refs.monthField.classList.toggle("hidden", preset !== "month");
  refs.startField.classList.toggle("hidden", preset !== "custom");
  refs.endField.classList.toggle("hidden", preset !== "custom");
}

function renderAll() {
  renderKpis();
  renderReconciliation();
  renderBreakdown();
  updateCharts();
  renderTable();
  const hasData = state.rows.length > 0 || state.mlOnly.length > 0;
  const hasSessionData = hasImportedData();
  refs.saveBackup.disabled = !hasSessionData;
  refs.clearSaved.disabled = !hasSessionData && !hasSavedSession();
  updateCloudControls();
  refs.exportXlsx.disabled = !hasData;
  refs.exportCsv.disabled = !hasData;
}

function renderKpis() {
  const rows = state.filteredClosing;
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
  return parts.length ? `base ${parts.join(" · ")}` : "base filtrada";
}

function renderBreakdown() {
  const rows = state.filteredClosing;
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

function renderTable() {
  if (state.activeTab === "closing") {
    renderClosingTable();
  } else {
    renderDiffTable();
  }
}

function renderClosingTable() {
  const rows = [...state.filteredClosing].sort(sortByDateDesc);
  refs.tableCount.textContent = `${formatInteger(rows.length)} registros`;
  refs.resultHead.innerHTML = `
    <tr>
      <th>Data</th>
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
    refs.resultBody.innerHTML = `<tr><td colspan="14" class="empty-state">Sem registros para os filtros atuais.</td></tr>`;
    return;
  }

  refs.resultBody.innerHTML = rows.map((row) => `
    <tr>
      <td>${escapeHtml(formatDate(row.date))}</td>
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
  refs.tableCount.textContent = `${formatInteger(rows.length)} divergências`;
  refs.resultHead.innerHTML = `
    <tr>
      <th>Tipo</th>
      <th>Data</th>
      <th>ID</th>
      <th>Status</th>
      <th>Produto</th>
      <th class="money">Recebido</th>
      <th class="money">Custo</th>
      <th>Observação</th>
    </tr>
  `;

  if (!rows.length) {
    refs.resultBody.innerHTML = `<tr><td colspan="8" class="empty-state">Sem divergências para os filtros atuais.</td></tr>`;
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
        <td><strong>${escapeHtml(row.orderId || "-")}</strong></td>
        <td><span class="status-pill ${statusClass(row)}">${escapeHtml(row.status || "-")}</span></td>
        <td>
          <strong>${escapeHtml(row.product || "Sem produto")}</strong>
          <small>${escapeHtml(row.sku ? `SKU ${row.sku}` : "")}</small>
        </td>
        <td class="money">${formatMoney(row.netReceived)}</td>
        <td class="money">${formatMoney(row.cost)}</td>
        <td>${escapeHtml(note)}</td>
      </tr>
    `;
  }).join("");
}

function setTab(tab) {
  state.activeTab = tab;
  refs.tabClosing.classList.toggle("active", tab === "closing");
  refs.tabDiff.classList.toggle("active", tab === "diff");
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
  const daily = aggregate(rows.filter((row) => row.date), (row) => dateKey(row.date), ["netReceived", "cost", "profit"]);
  const sortedDaily = [...daily.entries()].sort(([a], [b]) => a.localeCompare(b));
  state.charts.daily.data.labels = sortedDaily.map(([key]) => formatShortDateKey(key));
  state.charts.daily.data.datasets[0].data = sortedDaily.map(([, item]) => round2(item.netReceived));
  state.charts.daily.data.datasets[1].data = sortedDaily.map(([, item]) => round2(item.cost));
  state.charts.daily.data.datasets[2].data = sortedDaily.map(([, item]) => round2(item.profit));
  state.charts.daily.update();
  refs.dailyHint.textContent = `${formatInteger(rows.length)} vendas filtradas`;

  const byStatus = aggregate(rows, (row) => row.status || "Sem status", ["profit"]);
  const statusRows = [...byStatus.entries()]
    .sort((a, b) => b[1].profit - a[1].profit)
    .slice(0, 8);
  state.charts.status.data.labels = statusRows.map(([label]) => truncate(label, 24));
  state.charts.status.data.datasets[0].data = statusRows.map(([, item]) => round2(item.profit));
  state.charts.status.update();

  const byProduct = aggregate(rows, productKey, ["profit"]);
  const productRows = [...byProduct.entries()]
    .sort((a, b) => b[1].profit - a[1].profit)
    .slice(0, 8);
  state.charts.product.data.labels = productRows.map(([label]) => truncate(label, 28));
  state.charts.product.data.datasets[0].data = productRows.map(([, item]) => round2(item.profit));
  state.charts.product.update();
}

function exportWorkbook() {
  if (!window.XLSX) return;

  const metrics = currentMetrics();
  const cancelledIgnored = refs.ignoreCanceled?.checked
    ? [...state.rows, ...state.mlOnly].filter((row) => isCancelled(row) && matchesFilters(row, false, false)).length
    : 0;
  const summary = [
    ["Arquivo pedidos", state.platform?.fileName || ""],
    ["Arquivo Mercado Livre", state.marketplace?.fileName || ""],
    ["Modo do fechamento", financialBaseLabel()],
    ["Registros filtrados", state.filteredClosing.length],
    ["Canceladas ignoradas", cancelledIgnored],
    ["Recebido líquido", round2(metrics.net)],
    ["Custo de produtos", round2(metrics.cost)],
    ["Lucro líquido", round2(metrics.profit)],
    ["Margem", round2(metrics.margin)],
    ["ROI", round2(metrics.roi)],
    ["Pedidos sem Mercado Livre", state.filteredDiffs.filter((row) => row.source === "platform").length],
    ["Mercado Livre sem pedido", state.filteredDiffs.filter((row) => row.source === "marketplace").length],
  ];

  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.aoa_to_sheet(summary), "Resumo");
  window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.json_to_sheet(state.filteredClosing.map(toExportRow)), "Fechamento");
  window.XLSX.utils.book_append_sheet(wb, window.XLSX.utils.json_to_sheet(state.filteredDiffs.map(toExportRow)), "Divergencias");
  window.XLSX.writeFile(wb, `fechamento_dropshipping_${dateStamp()}.xlsx`);
}

function exportCsv() {
  if (!window.XLSX) return;
  const rows = state.activeTab === "closing" ? state.filteredClosing : state.filteredDiffs;
  const sheet = window.XLSX.utils.json_to_sheet(rows.map(toExportRow));
  const csv = "\uFEFF" + window.XLSX.utils.sheet_to_csv(sheet);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `${state.activeTab === "closing" ? "fechamento" : "divergencias"}_${dateStamp()}.csv`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(link.href);
}

function toExportRow(row) {
  return {
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

function currentMetrics() {
  const rows = state.filteredClosing;
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

  if (state.platform && state.marketplace) {
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

  panel.classList.add(state.platform || state.marketplace ? "warning" : "");
  panel.innerHTML = `<strong>Aguardando importação</strong><span>${escapeHtml(statusLine())}</span>`;
}

function statusLine() {
  const left = state.platform ? `Pedidos: ${state.platform.rows.length}` : "Pedidos pendente";
  const right = state.marketplace ? `Mercado Livre: ${state.marketplace.rows.length}` : "Mercado Livre pendente";
  return `${left} · ${right}`;
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

function statusClass(row) {
  if (!row.matched && row.source === "platform") return "bad";
  if (isCancelled(row)) return "bad";
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
