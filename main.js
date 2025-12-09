// ----- Visualizador de planilha -----
const output = document.getElementById("output");
const fetchBtn = document.getElementById("fetch-btn");
const toolbar = document.getElementById("toolbar");
const planilhaPanel = document.getElementById("planilha-panel");
const sheetSelect = document.getElementById("sheet-select");
const statusEl = document.getElementById("status");
const exportXlsxBtn = document.getElementById("export-xlsx-btn");
const exportCsvBtn = document.getElementById("export-csv-btn");
const exportJsonBtn = document.getElementById("export-json-btn");
const exportPdfBtn = document.getElementById("export-pdf-btn");
const barcodeStatus = document.getElementById("barcode-status");
const barcodeResult = document.getElementById("barcode-result");
const startScanBtn = document.getElementById("start-scan-btn");
const stopScanBtn = document.getElementById("stop-scan-btn");
const barcodeManual = document.getElementById("barcode-manual");
const barcodeGrid = document.getElementById("barcode-grid");
const barcodePanel = document.getElementById("barcode-panel");
const editModal = document.getElementById("edit-modal");
const editForm = document.getElementById("edit-form");
const editId = document.getElementById("edit-id");
const editEan = document.getElementById("edit-ean");
const editNome = document.getElementById("edit-nome");
const editQtd = document.getElementById("edit-qtd");
const editValidade = document.getElementById("edit-validade");
const editTroca = document.getElementById("edit-troca");
const editTrocaLabel = document.getElementById("edit-troca-label");
const editRota = document.getElementById("edit-rota");
const editRotaLabel = document.getElementById("edit-rota-label");
const editStatus = document.getElementById("edit-status");
const editCancelBtn = document.getElementById("edit-cancel-btn");
const editSaveBtn = document.getElementById("edit-save-btn");
const editTitle = document.getElementById("edit-title");
const editHint = document.getElementById("edit-hint");
const bulkActions = document.getElementById("bulk-actions");
const bulkCount = document.getElementById("bulk-count");
const bulkLaunchBtn = document.getElementById("bulk-launch-btn");
const bulkEditBtn = document.getElementById("bulk-edit-btn");
const bulkDeleteBtn = document.getElementById("bulk-delete-btn");
const logoutBtn = document.getElementById("logout-btn");
const envSelect = document.getElementById("env-select");
const testWebhooksBtn = document.getElementById("test-webhooks-btn");
const ENV_KEY = "gp_env";
let currentEnv = localStorage.getItem(ENV_KEY) === "test" ? "test" : "prod";

const styleEl = document.createElement("style");
styleEl.textContent = `
  .status-pill {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    min-width: 120px;
    padding: 6px 10px;
    border-radius: 999px;
    font-weight: 600;
    font-size: 0.85rem;
    border: 1px solid transparent;
  }
  .expiry-red { background: #fee2e2; color: #b91c1c; border-color: #fecdd3; }
  .expiry-yellow { background: #fef3c7; color: #b45309; border-color: #fde68a; }
  .expiry-green { background: #dcfce7; color: #15803d; border-color: #bbf7d0; }
  .lancado-true { background: #e0f2fe; color: #075985; border-color: #bae6fd; }
  .lancado-false { background: #f8fafc; color: #334155; border-color: #e2e8f0; }
  tr.expiry-red td { background: #fff7f7; }
  tr.expiry-yellow td { background: #fffaf0; }
  tr.expiry-green td { background: #f6fff8; }
`;
document.head.appendChild(styleEl);

const DELETE_PASSWORD_KEY = "gp_delete_password";
let renderedRowMap = new Map();
let selectedRowKeys = new Set();
let rowCheckboxes = [];
let selectAllCheckboxEl = null;
let currentEditTargets = [];

let workbook = null;
const versionBadge = document.getElementById("version-badge");
const BASE_VERSION = "1.3.8";

function setVersionBadge(versionText = BASE_VERSION) {
  if (versionBadge) {
    versionBadge.textContent = `Versão ${versionText}`;
  }
}

setVersionBadge(BASE_VERSION);

const AUTH_KEY = "gp_auth";
const USER_KEY = "gp_user";
const isAuthenticated = localStorage.getItem(AUTH_KEY) === "1";
const currentUserEmail = localStorage.getItem(USER_KEY) || "";

if (!isAuthenticated || !currentUserEmail) {
  localStorage.removeItem(AUTH_KEY);
  window.location.href = "login.html";
} else if (planilhaPanel) {
  planilhaPanel.style.display = "block";
}

if (logoutBtn) {
  logoutBtn.addEventListener("click", () => {
    localStorage.removeItem(AUTH_KEY);
    localStorage.removeItem(USER_KEY);
    window.location.href = "login.html";
  });
}

if (envSelect) {
  envSelect.value = currentEnv;
  envSelect.addEventListener("change", () => setEnvironment(envSelect.value));
  // garante status inicial
  setEnvironment(currentEnv);
  envSelect.disabled = true;
  envSelect.title = "Seleção de ambiente temporariamente desabilitada";
}

if (testWebhooksBtn) {
  testWebhooksBtn.addEventListener("click", testWebhooks);
  testWebhooksBtn.disabled = true;
  testWebhooksBtn.title = "Teste de webhooks temporariamente desabilitado";
}

if (editTroca) {
  editTroca.addEventListener("change", () => {
    editTroca.indeterminate = false;
    updateSwitchLabel(editTroca, editTrocaLabel);
  });
  updateSwitchLabel(editTroca, editTrocaLabel);
}

if (editRota) {
  editRota.addEventListener("change", () => {
    editRota.indeterminate = false;
    updateSwitchLabel(editRota, editRotaLabel);
  });
  updateSwitchLabel(editRota, editRotaLabel);
}

const devPorts = ["5500", "5501", "3000", "5173"];
const preferExternal = devPorts.includes(window.location.port || "");

const PROD_ENDPOINTS = {
  planilha: preferExternal
    ? [
        "https://myn8n.seommerce.shop/webhook/planilha-atualizada",
        "/api/planilha-atualizada",
      ]
    : [
        "/api/planilha-atualizada",
        "https://myn8n.seommerce.shop/webhook/planilha-atualizada",
      ],
  produto: preferExternal
    ? [
        "https://myn8n.seommerce.shop/webhook/validade",
        "/api/validade",
      ]
    : [
        "/api/validade",
        "https://myn8n.seommerce.shop/webhook/validade",
      ],
  barcode: preferExternal
    ? [
        "https://myn8n.seommerce.shop/webhook/barcode",
        "/api/barcode",
      ]
    : [
        "/api/barcode",
        "https://myn8n.seommerce.shop/webhook/barcode",
      ],
  actions: ["https://myn8n.seommerce.shop/webhook/a%C3%A7%C3%B5es"],
};

const TEST_ENDPOINTS = {
  planilha: [
    "https://myn8n.seommerce.shop/webhook-test/planilha-atualizada",
  ],
  produto: ["https://myn8n.seommerce.shop/webhook-test/validade"],
  barcode: ["https://myn8n.seommerce.shop/webhook-test/barcode"],
  actions: ["https://myn8n.seommerce.shop/webhook-test/a%C3%A7%C3%B5es"],
};

function setEnvironment(env) {
  currentEnv = env === "test" ? "test" : "prod";
  localStorage.setItem(ENV_KEY, currentEnv);
  if (envSelect) envSelect.value = currentEnv;
  if (statusEl) {
    statusEl.textContent = `Ambiente em uso: ${currentEnv === "test" ? "Teste" : "Produção"}.`;
  }
}

function getEndpoints(key) {
  if (key === "planilha") return PROD_ENDPOINTS.planilha;
  const envSet = currentEnv === "test" ? TEST_ENDPOINTS : PROD_ENDPOINTS;
  return envSet[key] || [];
}

function getUserPayload() {
  if (!currentUserEmail) return {};
  return { user: currentUserEmail, email: currentUserEmail };
}

async function testWebhooks() {
  const envChoice = envSelect?.value === "test" ? "test" : "prod";
  setEnvironment(envChoice);
  const envLabel = envChoice === "test" ? "Teste" : "Produção";
  const set = envChoice === "test" ? TEST_ENDPOINTS : PROD_ENDPOINTS;
  const results = [];
  statusEl.textContent = `Testando webhooks (${envLabel})...`;
  const tests = [
    {
      label: "ações",
      endpoints: set.actions,
      body: { action: "ping", timestamp: Date.now() },
    },
    {
      label: "planilha",
      endpoints: set.planilha,
      body: { origem: "frontend-teste" },
    },
    {
      label: "produtos",
      endpoints: set.produto,
      body: { action: "ping" },
    },
  ];

  for (const test of tests) {
    try {
      await postJsonWithFallback(test.endpoints, test.body);
      results.push(`${test.label}: ok`);
    } catch (err) {
      results.push(`${test.label}: erro (${err.message || err})`);
    }
  }
  statusEl.textContent = `Resultado (${envLabel}): ${results.join(" | ")}`;
}

const PRODUCT_HEADERS = [
  "id",
  "nome",
  "validade",
  "quantidade",
  "ean",
  "troca",
  "lancado",
  "data_lancado",
  "rotatividade_alta",
  "adicionado_em",
];

async function postJsonWithFallback(urls, body) {
  let lastError = null;
  const attempts = [];
  for (const url of urls) {
    try {
      const res = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(body),
      });
      if (!res.ok) {
        throw new Error("Erro HTTP " + res.status);
      }
      return res;
    } catch (err) {
      lastError = err;
      attempts.push(`${url}: ${err.message || err}`);
      console.warn("Erro ao chamar webhook:", url, err);
      continue;
    }
  }
  throw lastError || new Error(`Falha ao chamar webhook. Tentativas: ${attempts.join(" | ")}`);
}

async function postFormWithFallback(urls, formData) {
  let lastError = null;
  const attempts = [];
  for (const url of urls) {
    try {
      const res = await fetch(url, {
        method: "POST",
        body: formData,
      });
      if (!res.ok) {
        throw new Error("Erro HTTP " + res.status);
      }
      return res;
    } catch (err) {
      lastError = err;
      attempts.push(`${url}: ${err.message || err}`);
      console.warn("Erro ao enviar produto:", url, err);
      continue;
    }
  }
  throw lastError || new Error(`Falha ao enviar produto. Tentativas: ${attempts.join(" | ")}`);
}

function normalizeProductEntry(entry = {}) {
  return {
    id: entry.id ?? "",
    nome: entry.nome ?? entry.name ?? "",
    validade: entry.validade ?? entry.vencimento ?? "",
    quantidade: entry.quantidade ?? entry.qtd ?? "",
    ean: entry.ean ?? entry.codigo ?? "",
    troca: toBoolean(entry.troca, false),
    lancado: toBoolean(entry.lancado ?? entry["lançado"], false),
    rotatividade_alta: toBoolean(entry.rotatividade_alta ?? entry.rotatividade, false),
    data_lancado: entry.data_lancado ?? entry.dataLancado ?? "",
    adicionado_em: entry.adicionado_em ?? entry.criado_em ?? "",
  };
}

function buildWorkbookFromProducts(data) {
  const list = Array.isArray(data) ? data.map(normalizeProductEntry) : [];
  const rows = [PRODUCT_HEADERS];
  list.forEach((item) => {
    rows.push(PRODUCT_HEADERS.map((key) => item[key] ?? ""));
  });
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Produtos");
  return { wb, sheetName: "Produtos" };
}

async function fetchProductFromAction(ean) {
  const ACTION_TIMEOUT_MS = 4000;
  try {
    const res = await Promise.race([
      postJsonWithFallback(getEndpoints("actions"), {
        action: "produto",
        ean,
        ...getUserPayload(),
      }),
      new Promise((_, reject) =>
        setTimeout(() => reject(new Error("timeout ação produto")), ACTION_TIMEOUT_MS)
      ),
    ]);
    const contentType = res.headers?.get?.("content-type") || "";
    if (!contentType.includes("application/json")) return null;
    let data = null;
    try {
      data = await res.json();
    } catch {
      return null;
    }
    const product = Array.isArray(data) ? data[0] : data;
    if (!product) return null;

    const message =
      typeof product === "string"
        ? product
        : product.message || product.msg || product.status || "";
    const lowerMessage = typeof message === "string" ? message.toLowerCase() : "";
    if (
      lowerMessage.includes("não cadastrado") ||
      lowerMessage.includes("nao cadastrado")
    ) {
      return null;
    }

    const normalized = normalizeProductEntry(product);
    const hasData = Object.values(normalized).some((v) => Boolean(v));
    return hasData ? normalized : null;
  } catch (err) {
    console.warn("Erro ao buscar produto no webhook de ações:", err);
    return null;
  }
}

function reorderColumns(headerRowValues, dataRows) {
  const lower = (headerRowValues || []).map((cell) =>
    String(cell || "").trim().toLowerCase()
  );
  const idIdx = lower.indexOf("id");
  const lancadoIdx = lower.findIndex(
    (h) => h === "lançado" || h === "lancado"
  );
  const dataLancadoIdx = lower.findIndex(
    (h) => h === "data_lançado" || h === "data_lancado" || h === "datalancado"
  );

  let order = headerRowValues.map((_, i) => i);

  // Move data_lancado logo após lancado
  if (lancadoIdx !== -1 && dataLancadoIdx !== -1 && dataLancadoIdx !== lancadoIdx + 1) {
    order = order.filter((i) => i !== dataLancadoIdx);
    const insertPos = Math.max(order.indexOf(lancadoIdx) + 1, 0);
    order.splice(insertPos, 0, dataLancadoIdx);
  }

  // Mover ID para primeira coluna
  if (idIdx > 0) {
    order = order.filter((i) => i !== idIdx);
    order.unshift(idIdx);
  }

  const newHeader = order.map((i) => headerRowValues[i]);
  const newDataRows = dataRows.map((row) => order.map((i) => row[i]));

  return { headerRowValues: newHeader, dataRows: newDataRows };
}

// ----- Leitor de código de barras (Html5Qrcode/ZXing) -----
function setBarcodeStatus(text, type = "") {
  if (!barcodeStatus) return;
  barcodeStatus.textContent = text;
  barcodeStatus.className = "";
  if (type === "ok") barcodeStatus.classList.add("ok");
  if (type === "erro") barcodeStatus.classList.add("erro");
}

async function lookupProductByEan(ean) {
  try {
    const res = await fetch(
      `https://world.openfoodfacts.org/api/v0/product/${ean}.json`
    );
    if (!res.ok) return null;
    const data = await res.json();
    if (data && data.status === 1 && data.product) {
      return (
        data.product.product_name ||
        data.product.generic_name ||
        data.product.brands
      );
    }
  } catch (err) {
    console.warn("Lookup EAN falhou:", err);
  }
  return null;
}

async function processDetectedEan(ean, source = "scanner") {
  const onlyDigits = (ean || "").replace(/\D/g, "");
  if (!onlyDigits) {
    setBarcodeStatus("EAN inválido ou vazio.", "erro");
    return;
  }
  const validLength = [8, 12, 13].includes(onlyDigits.length);
  if (!validLength) {
    setBarcodeStatus("EAN deve ter 8, 12 ou 13 dígitos.", "erro");
    return;
  }

  ensureFormPanelVisible();
  ensureAtLeastOneItem();
  barcodeResult.textContent = `${source === "manual" ? "EAN digitado" : "EAN detectado"}: ${onlyDigits}`;
  if (barcodeManual) barcodeManual.value = onlyDigits;

  const existing = itens.find((it) => it.codigo === onlyDigits);
  const emptySlot = itens.find((it) => !it.codigo);
  setBarcodeStatus("Consultando produto no servidor...");
  const fetchedProduct = await fetchProductFromAction(onlyDigits);
  const foundFromAction = Boolean(fetchedProduct);
  let nome = fetchedProduct?.nome || null;
  if (!nome) {
    nome = await lookupProductByEan(onlyDigits);
  }

  if (existing) {
    if (!existing.nome && nome) existing.nome = nome;
    if (fetchedProduct) {
      applyProductInfoToItem(existing, fetchedProduct);
    }
    renderItens();
    setBarcodeStatus(
      foundFromAction
        ? "Produto encontrado e dados preenchidos. EAN já estava na lista."
        : "EAN já adicionado à lista.",
      "ok"
    );
    focusItemCode(existing);
  } else if (emptySlot) {
    emptySlot.codigo = onlyDigits;
    emptySlot.allowEditCode = false;
    if (!emptySlot.nome && nome) emptySlot.nome = nome;
    if (fetchedProduct) {
      applyProductInfoToItem(emptySlot, fetchedProduct);
    }
    renderItens();
    setBarcodeStatus(
      foundFromAction
        ? "Produto encontrado e pré-preenchido. Confira e envie."
        : "Produto não encontrado, preencha e envie para cadastrar.",
      "ok"
    );
    focusItemCode(emptySlot);
  } else {
    const newItem = createItemFromCode(onlyDigits, nome || "");
    if (fetchedProduct) {
      applyProductInfoToItem(newItem, fetchedProduct);
    }
    itens.push(newItem);
    renderItens();
    setBarcodeStatus(
      foundFromAction
        ? "Produto encontrado e pré-preenchido. Confira e envie."
        : "Produto não encontrado, preencha e envie para cadastrar.",
      "ok"
    );
    focusItemCode(newItem);
  }
}

const html5Formats = [
  Html5QrcodeSupportedFormats.QR_CODE,
  Html5QrcodeSupportedFormats.CODE_128,
  Html5QrcodeSupportedFormats.CODE_39,
  Html5QrcodeSupportedFormats.EAN_13,
  Html5QrcodeSupportedFormats.EAN_8,
  Html5QrcodeSupportedFormats.UPC_A,
  Html5QrcodeSupportedFormats.UPC_E,
  Html5QrcodeSupportedFormats.ITF
];

let html5QrcodeScanner = null;
let html5Running = false;

async function startHtml5Scanner() {
  if (html5Running) return;
  setBarcodeStatus("Abrindo câmera...");
  barcodeResult.textContent = "";

  try {
    if (!html5QrcodeScanner) {
      html5QrcodeScanner = new Html5Qrcode("reader1", { formatsToSupport: html5Formats });
    }

    const successCallback = async (decodedText) => {
      await processDetectedEan(decodedText, "scanner");
      html5QrcodeScanner.pause();
      setTimeout(() => html5QrcodeScanner.resume(), 800);
    };

    const config = {
      fps: 10,
      qrbox: { width: 220, height: 220 },
      aspectRatio: 1.7777778,
      formatsToSupport: html5Formats,
      experimentalFeatures: { useBarCodeDetectorIfSupported: true },
      rememberLastUsedCamera: true
    };

    await html5QrcodeScanner.start(
      { facingMode: "environment" },
      config,
      successCallback
    );

    html5Running = true;
    startScanBtn.disabled = true;
    stopScanBtn.disabled = false;
    if (barcodeGrid) barcodeGrid.classList.remove("collapsed");
    setBarcodeStatus("Câmera aberta, escaneando...");
  } catch (err) {
    console.error(err);
    setBarcodeStatus(`Erro: ${err.message || err}`, "erro");
    await stopHtml5Scanner(true);
  }
}

async function stopHtml5Scanner(skipStatus = false) {
  if (html5QrcodeScanner) {
    try {
      await html5QrcodeScanner.stop();
    } catch (err) {
      console.warn("Erro ao parar Html5Qrcode:", err);
    }
  }
  html5Running = false;
  startScanBtn.disabled = false;
  stopScanBtn.disabled = true;
  if (barcodeGrid) barcodeGrid.classList.add("collapsed");
  if (!skipStatus) setBarcodeStatus("Scanner parado.");
}

function initBarcodeUI() {
  setBarcodeStatus("Pronto para escanear. Clique em Abrir câmera.");
  if (startScanBtn) startScanBtn.addEventListener("click", startHtml5Scanner);
  if (stopScanBtn) stopScanBtn.addEventListener("click", () => stopHtml5Scanner());
  stopScanBtn.disabled = true;
  if (barcodeGrid) barcodeGrid.classList.add("collapsed");
}

initBarcodeUI();

fetchBtn.addEventListener("click", async () => {
  if (!isAuthenticated) {
    window.location.href = "login.html";
    return;
  }
  clearSelections();
  renderedRowMap = new Map();
  updateBulkActionsUI();
  output.innerHTML = "";
  statusEl.textContent = "Carregando planilha do servidor...";

  try {
    const res = await postJsonWithFallback(getEndpoints("planilha"), {
      origem: "frontend-xlsx-viewer",
      ...getUserPayload(),
    });

    const handleJsonResponse = async (response) => {
      try {
        const data = await response.clone().json();
        const { wb, sheetName } = buildWorkbookFromProducts(data);
        workbook = wb;

        sheetSelect.innerHTML = "";
        const option = document.createElement("option");
        option.value = sheetName;
        option.textContent = sheetName;
        option.selected = true;
        sheetSelect.appendChild(option);
        toolbar.style.display = "flex";
        renderSheet(sheetName);
        statusEl.textContent = "Planilha carregada com sucesso.";
        return true;
      } catch (err) {
        return false;
      }
    };

    const contentType = res.headers.get("content-type") || "";
    if (contentType.includes("application/json")) {
      const handled = await handleJsonResponse(res);
      if (handled) return;
    } else {
      const handledJson = await handleJsonResponse(res);
      if (handledJson) return;
    }

    workbook = await responseToWorkbook(res);

    if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) {
      statusEl.textContent = "Nenhuma planilha encontrada na resposta.";
      output.innerHTML = "";
      toolbar.style.display = "none";
      return;
    }

    sheetSelect.innerHTML = "";
    workbook.SheetNames.forEach((name, index) => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      if (index === 0) option.selected = true;
      sheetSelect.appendChild(option);
    });

    toolbar.style.display = workbook.SheetNames.length ? "flex" : "none";
    renderSheet(workbook.SheetNames[0]);
    statusEl.textContent = "Planilha carregada com sucesso.";
  } catch (err) {
    console.error(err);
    statusEl.textContent = `Erro ao buscar planilha: ${err.message}`;
  }
});

sheetSelect.addEventListener("change", () => {
  if (!workbook) return;
  renderSheet(sheetSelect.value);
});

const BR_TIMEZONE = "America/Sao_Paulo";
const brDateFormatter = new Intl.DateTimeFormat("pt-BR", {
  timeZone: BR_TIMEZONE,
  day: "2-digit",
  month: "2-digit",
  year: "numeric",
});
const brDateTimeFormatter = new Intl.DateTimeFormat("pt-BR", {
  timeZone: BR_TIMEZONE,
  day: "2-digit",
  month: "2-digit",
  year: "numeric",
  hour: "2-digit",
  minute: "2-digit",
});

function makeUtcDate(year, month, day, hour = 0, minute = 0, second = 0) {
  const date = new Date(Date.UTC(year, month - 1, day, hour, minute, second));
  return date;
}

function tryParseDate(value) {
  if (!value) return null;
  const cleaned = String(value).trim();
  if (!cleaned) return null;

  // Normaliza separadores: vírgula ou T viram espaço, múltiplos espaços colapsam
  const normalized = cleaned.replace(/,/g, " ").replace(/T/g, " ").replace(/\s+/g, " ").trim();

  if (!Number.isNaN(Date.parse(normalized))) {
    const dt = new Date(normalized);
    return { date: dt, hasTime: true };
  }

  if (typeof value === "number") {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const millisecondsPerDay = 24 * 60 * 60 * 1000;
    const utcDate = new Date(excelEpoch.getTime() + value * millisecondsPerDay);
    return { date: utcDate, hasTime: true };
  }

  const parts = normalized.split(/\s+/);
  if (parts.length >= 1) {
    const iso = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/;
    const match = iso.exec(normalized);
    if (match) {
      const [, y, m, d, hh, mm, ss] = match;
      const dt = makeUtcDate(y, m, d, hh || 0, mm || 0, ss || 0);
      return {
        date: dt,
        hasTime: Boolean(hh || mm || ss),
        parts: { day: Number(d), month: Number(m), year: Number(y) },
      };
    }

    const brDateTime =
      /^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})(?::(\d{2}))?$/;
    const brDTMatch = brDateTime.exec(normalized);
    if (brDTMatch) {
      const [, d, m, y, hh, mm, ss] = brDTMatch;
      const dt = makeUtcDate(y, m, d, Number(hh) + 3, mm, ss || 0);
      return { date: dt, hasTime: true };
    }

    const br = /^(\d{2})\/(\d{2})\/(\d{4})$/;
    const brMatch = br.exec(normalized);
    if (brMatch) {
      const [, d, m, y] = brMatch;
      const dt = makeUtcDate(y, m, d);
      return {
        date: dt,
        hasTime: false,
        parts: { day: Number(d), month: Number(m), year: Number(y) },
      };
    }
  }

  return null;
}

function formatCellValue(value) {
  const parsed = tryParseDate(value);
  if (parsed?.date) {
    if (parsed.hasTime) {
      return brDateTimeFormatter.format(parsed.date);
    }
    const year = parsed.parts?.year ?? parsed.date.getUTCFullYear();
    const month = parsed.parts?.month ?? parsed.date.getUTCMonth() + 1;
    const day = parsed.parts?.day ?? parsed.date.getUTCDate();
    const dia = String(day).padStart(2, "0");
    const mes = String(month).padStart(2, "0");
    return `${dia}/${mes}/${year}`;
  }
  return value ?? "";
}

function getExpiryMeta(value) {
  const parsed = tryParseDate(value);
  if (!parsed?.date) return null;
  const now = Date.now();
  const diffMs = parsed.date.getTime() - now;
  const days = Math.floor(diffMs / (24 * 3600 * 1000));
  let tag = "";
  let color = "";
  if (diffMs < 0) {
    tag = "⚠️ Vencido";
    color = "expiry-red";
  } else if (days <= 3) {
    tag = `⚠️ Vence em ${days === 0 ? "hoje" : `${days} dias`}`;
    color = "expiry-red";
  } else if (days <= 7) {
    tag = `Vence em ${days} dias`;
    color = "expiry-yellow";
  } else if (days > 30) {
    tag = "Vence em 30+ dias";
    color = "expiry-green";
  } else {
    tag = `Vence em ${days} dias`;
    color = "expiry-yellow";
  }
  return { days, tag, color };
}

function toBoolean(value, defaultValue = false) {
  if (typeof value === "boolean") return value;
  if (value === undefined || value === null || value === "") return defaultValue;
  const normalized = String(value).toLowerCase().trim();
  if (normalized === "true") return true;
  if (normalized === "false") return false;
  return defaultValue;
}

function booleanLabel(value, defaultValue = false) {
  return toBoolean(value, defaultValue) ? "SIM" : "NÃO";
}

function getValidityTimestamp(row, idx) {
  if (!row || idx < 0) return Number.POSITIVE_INFINITY;
  const parsed = tryParseDate(row[idx]);
  return parsed?.date?.getTime?.() ?? Number.POSITIVE_INFINITY;
}

function updateSwitchLabel(input, labelEl) {
  if (!input || !labelEl) return;
  if (input.indeterminate) {
    labelEl.textContent = "MANTER";
  } else {
    labelEl.textContent = input.checked ? "SIM" : "NÃO";
  }
}

function toDateInputValue(value) {
  const parsed = tryParseDate(value);
  if (!parsed?.date) return "";
  const year = parsed.parts?.year ?? parsed.date.getUTCFullYear();
  const month = parsed.parts?.month ?? parsed.date.getUTCMonth() + 1;
  const day = parsed.parts?.day ?? parsed.date.getUTCDate();
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function getLancadoMeta(value) {
  const isTrue = toBoolean(value, false);
  return isTrue
    ? { text: "Lançado", color: "lancado-true" }
    : { text: "Não lançado", color: "lancado-false" };
}

function columnLetter(index) {
  let result = "";
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function getCurrentSheet() {
  if (!workbook || !workbook.SheetNames || !workbook.SheetNames.length) {
    return null;
  }
  const name = sheetSelect.value || workbook.SheetNames[0];
  const sheet = workbook.Sheets[name];
  if (!sheet) return null;
  return { name, sheet };
}

function findFirstArray(value, depth = 0) {
  if (!value || depth > 3) return null;
  if (Array.isArray(value)) return value;
  if (typeof value !== "object") return null;
  for (const key of Object.keys(value)) {
    const val = value[key];
    if (Array.isArray(val)) return val;
  }
  for (const key of Object.keys(value)) {
    const found = findFirstArray(value[key], depth + 1);
    if (found) return found;
  }
  return null;
}

function normalizeJsonRows(json) {
  if (Array.isArray(json)) return json;
  const preferredKeys = [
    "data",
    "items",
    "rows",
    "registros",
    "produtos",
    "result",
    "results",
  ];
  for (const key of preferredKeys) {
    if (json && Array.isArray(json[key])) {
      return json[key];
    }
  }
  const firstArray = findFirstArray(json);
  if (firstArray) return firstArray;
  return [];
}

function responseToWorkbook(res) {
  return res.blob().then((blob) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const arr = [];
        for (let i = 0; i !== data.length; ++i) arr[i] = String.fromCharCode(data[i]);
        const bstr = arr.join("");
        resolve(XLSX.read(bstr, { type: "binary" }));
      };
      reader.readAsArrayBuffer(blob);
    });
  });
}

function downloadBlob(content, mimeType, filename) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function exportCurrentSheetToPdf() {
  const current = getCurrentSheet();
  if (!current) {
    statusEl.textContent = "Nenhuma planilha para exportar.";
    return;
  }
  const rows = XLSX.utils.sheet_to_json(current.sheet, { header: 1, defval: "" });
  if (!rows.length) {
    statusEl.textContent = "Planilha vazia.";
    return;
  }
  const win = window.open("", "_blank", "width=1000,height=800");
  if (!win) {
    statusEl.textContent = "Permita pop-ups para exportar em PDF.";
    return;
  }

  const tableHtml = rows
    .map((row, rowIndex) => {
      const tag = rowIndex === 0 ? "th" : "td";
      const cells = (row || [])
        .map((cell) => `<${tag}>${escapeHtml(cell)}</${tag}>`)
        .join("");
      return `<tr>${cells}</tr>`;
    })
    .join("");

  const styles = `
    <style>
      body { font-family: "Inter", system-ui, -apple-system, "Segoe UI", sans-serif; padding: 24px; color: #0f172a; }
      h3 { margin-bottom: 12px; }
      table { width: 100%; border-collapse: collapse; font-size: 12px; }
      th, td { border: 1px solid #cbd5e1; padding: 8px 10px; text-align: left; }
      th { background: #e2e8f0; font-weight: 700; }
      tr:nth-child(even) td { background: #f8fafc; }
    </style>
  `;

  win.document.write(`
    <html>
      <head><title>${escapeHtml(current.name)} - PDF</title>${styles}</head>
      <body>
        <h3>Planilha: ${escapeHtml(current.name)}</h3>
        <table>${tableHtml}</table>
      </body>
    </html>
  `);
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 300);
}

exportXlsxBtn.addEventListener("click", () => {
  const current = getCurrentSheet();
  if (!current) {
    statusEl.textContent = "Nenhuma planilha para exportar.";
    return;
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, current.sheet, current.name || "Dados");
  XLSX.writeFile(wb, `dados-${current.name || "planilha"}.xlsx`);
});

exportCsvBtn.addEventListener("click", () => {
  const current = getCurrentSheet();
  if (!current) {
    statusEl.textContent = "Nenhuma planilha para exportar.";
    return;
  }
  const csv = XLSX.utils.sheet_to_csv(current.sheet);
  downloadBlob(
    csv,
    "text/csv;charset=utf-8;",
    `dados-${current.name || "planilha"}.csv`
  );
});

exportJsonBtn.addEventListener("click", () => {
  const current = getCurrentSheet();
  if (!current) {
    statusEl.textContent = "Nenhuma planilha para exportar.";
    return;
  }
  const json = XLSX.utils.sheet_to_json(current.sheet, { defval: "" });
  const jsonStr = JSON.stringify(json, null, 2);
  downloadBlob(
    jsonStr,
    "application/json;charset=utf-8;",
    `dados-${current.name || "planilha"}.json`
  );
});

if (exportPdfBtn) {
  exportPdfBtn.addEventListener("click", exportCurrentSheetToPdf);
}

function requestDeletePassword() {
  const stored = localStorage.getItem(DELETE_PASSWORD_KEY);
  const message = stored
    ? "Digite a senha para eliminar produtos:"
    : "Defina uma senha de proteção para eliminar produtos:";
  const input = window.prompt(message);
  if (input === null) return null;
  const pass = (input || "").trim();
  if (!pass) {
    alert("Senha obrigatória para eliminar produtos.");
    return null;
  }
  if (!stored) {
    localStorage.setItem(DELETE_PASSWORD_KEY, pass);
    return pass;
  }
  if (pass !== stored) {
    alert("Senha incorreta. Operação cancelada.");
    return null;
  }
  return pass;
}

function getRowKey(payload, rowNumber) {
  const base = payload?.id || payload?.ean || `row-${rowNumber}`;
  return `${base}::${rowNumber}`;
}

function clearSelections() {
  selectedRowKeys.clear();
  rowCheckboxes.forEach(({ checkbox, tr }) => {
    if (checkbox) checkbox.checked = false;
    if (tr) tr.classList.remove("selected-row");
  });
  if (selectAllCheckboxEl) {
    selectAllCheckboxEl.checked = false;
    selectAllCheckboxEl.indeterminate = false;
  }
  updateBulkActionsUI();
}

function toggleRowSelection(rowKey, checked, tr) {
  if (checked) {
    selectedRowKeys.add(rowKey);
    if (tr) tr.classList.add("selected-row");
  } else {
    selectedRowKeys.delete(rowKey);
    if (tr) tr.classList.remove("selected-row");
  }
  updateBulkActionsUI();
}

function selectAllRows(checked) {
  selectedRowKeys.clear();
  rowCheckboxes.forEach(({ key, checkbox, tr }) => {
    if (checkbox) checkbox.checked = checked;
    if (checked) {
      selectedRowKeys.add(key);
      if (tr) tr.classList.add("selected-row");
    } else if (tr) {
      tr.classList.remove("selected-row");
    }
  });
  updateBulkActionsUI();
}

function updateBulkActionsUI() {
  if (!bulkActions) return;
  const count = selectedRowKeys.size;
  const total = renderedRowMap.size;
  bulkActions.style.display = count ? "flex" : "none";
  if (bulkCount) {
    bulkCount.textContent = count
      ? `${count} produto(s) selecionado(s)`
      : "Nenhum produto selecionado.";
  }
  if (bulkLaunchBtn) bulkLaunchBtn.disabled = count === 0;
  if (bulkEditBtn) bulkEditBtn.disabled = count === 0;
  if (bulkDeleteBtn) bulkDeleteBtn.disabled = count === 0;
  if (selectAllCheckboxEl) {
    selectAllCheckboxEl.checked = total > 0 && count === total;
    selectAllCheckboxEl.indeterminate = count > 0 && count < total;
  }
}

function getSelectedPayloads() {
  const items = [];
  for (const key of selectedRowKeys) {
    const payload = renderedRowMap.get(key);
    if (payload) items.push(payload);
  }
  return items;
}

function setBulkButtonsDisabled(disabled) {
  [bulkLaunchBtn, bulkEditBtn, bulkDeleteBtn].forEach((btn) => {
    if (btn) btn.disabled = disabled || selectedRowKeys.size === 0;
  });
}

async function launchTargets(targets) {
  const list = (targets || []).filter(Boolean);
  if (!list.length) {
    statusEl.textContent = "Selecione ao menos um produto para lançar.";
    return;
  }
  setBulkButtonsDisabled(true);
  statusEl.textContent = `Lançando ${list.length} produto(s)...`;
  try {
    let sent = 0;
    for (const payload of list) {
      const idVal = payload?.id || payload?.ean || "";
      if (!idVal) continue;
      await postJsonWithFallback(getEndpoints("actions"), {
        action: "lancar",
        id: idVal,
        lancado: true,
        ...getUserPayload(),
      });
      sent += 1;
    }
    statusEl.textContent = sent
      ? "Produtos marcados como lançados."
      : "Nenhum produto válido para lançar.";
    clearSelections();
    if (fetchBtn) fetchBtn.click();
  } catch (err) {
    statusEl.textContent = `Erro ao lançar produtos: ${err.message || err}`;
  } finally {
    setBulkButtonsDisabled(false);
  }
}

async function deleteTargets(targets) {
  const list = (targets || []).filter(Boolean);
  if (!list.length) {
    statusEl.textContent = "Selecione ao menos um produto para eliminar.";
    return;
  }
  const password = requestDeletePassword();
  if (!password) {
    statusEl.textContent = "Eliminação cancelada.";
    return;
  }
  setBulkButtonsDisabled(true);
  statusEl.textContent = `Eliminando ${list.length} produto(s)...`;
  try {
    let sent = 0;
    for (const payload of list) {
      const idVal = payload?.id || payload?.ean || "";
      if (!idVal) continue;
      await postJsonWithFallback(getEndpoints("actions"), {
        action: "apagar",
        id: idVal,
        password,
        ...getUserPayload(),
      });
      sent += 1;
    }
    statusEl.textContent = sent
      ? `Solicitação de eliminação enviada para ${sent} produto(s).`
      : "Nenhum produto válido para eliminar.";
    clearSelections();
    if (fetchBtn) fetchBtn.click();
  } catch (err) {
    statusEl.textContent = `Erro ao eliminar: ${err.message || err}`;
  } finally {
    setBulkButtonsDisabled(false);
  }
}

function normalizeHeaderLabel(label, fallback) {
  const raw = (label ?? "").toString().trim();
  if (!raw) return fallback;
  const normalized = raw
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
  const map = {
    id: "ID",
    nome: "NOME",
    validade: "VALIDADE",
    vencimento: "VALIDADE",
    quantidade: "QUANTIDADE",
    qtd: "QUANTIDADE",
    ean: "EAN",
    troca: "TROCA",
    lancado: "LANCADO",
    lancado_: "LANCADO",
    lanado: "LANCADO",
    "data_lancado": "DATA_LANCADO",
    "data_lancamento": "DATA_LANCADO",
    datalancado: "DATA_LANCADO",
    "rotatividade_alta": "ROTATIVIDADE_ALTA",
    rotatividade: "ROTATIVIDADE_ALTA",
    "adicionado_em": "ADICIONADO_EM",
    "criado_em": "ADICIONADO_EM",
    acoes: "AÇÕES",
    acoes_: "AÇÕES",
    actions: "AÇÕES",
  };
  if (map[normalized]) return map[normalized];
  return raw.toUpperCase();
}

function renderSheet(name) {
  const ws = workbook.Sheets[name];
  if (!ws) {
    output.innerHTML = "<p>Planilha não encontrada.</p>";
    clearSelections();
    renderedRowMap = new Map();
    updateBulkActionsUI();
    return;
  }
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!rows.length) {
    output.innerHTML = "<p>Nenhuma linha encontrada.</p>";
    clearSelections();
    renderedRowMap = new Map();
    updateBulkActionsUI();
    return;
  }

  clearSelections();
  renderedRowMap = new Map();
  selectedRowKeys = new Set();
  rowCheckboxes = [];
  selectAllCheckboxEl = null;
  updateBulkActionsUI();

  let headerRowValues = rows[0] || [];
  let dataRows = rows.slice(1);
  const reordered = reorderColumns(headerRowValues, dataRows);
  headerRowValues = reordered.headerRowValues || headerRowValues;
  dataRows = reordered.dataRows || dataRows;
  const maxCols = Math.max(
    headerRowValues.length,
    ...dataRows.map((row) => row.length)
  );

  const headerLower = headerRowValues.map((cell) =>
    String(cell || "").trim().toLowerCase()
  );
  const validadeColumnIndex = headerLower.findIndex(
    (h) => h === "validade" || h === "vencimento" || h === "data"
  );
  if (validadeColumnIndex !== -1) {
    dataRows = dataRows
      .slice()
      .sort(
        (a, b) =>
          getValidityTimestamp(a, validadeColumnIndex) -
          getValidityTimestamp(b, validadeColumnIndex)
      );
  }
  const skuColumnIndex = headerRowValues.findIndex((cell) => {
    if (cell === null || cell === undefined) return false;
    return String(cell).trim().toLowerCase() === "sku";
  });
  const eanColumnIndex = headerRowValues.findIndex((cell) => {
    if (cell === null || cell === undefined) return false;
    return String(cell).trim().toLowerCase() === "ean";
  });
  const lancadoColumnIndex = headerLower.findIndex(
    (h) => h === "lançado" || h === "lancado"
  );

  const table = document.createElement("table");
  table.className = "sheet-table";

  const thead = document.createElement("thead");
  const headerRow = document.createElement("tr");

  const corner = document.createElement("th");
  corner.textContent = "";
  headerRow.appendChild(corner);

  const selectAllTh = document.createElement("th");
  selectAllTh.className = "select-col";
  const selectAllInput = document.createElement("input");
  selectAllInput.type = "checkbox";
  selectAllInput.title = "Selecionar todos os produtos visíveis";
  selectAllInput.addEventListener("change", () =>
    selectAllRows(selectAllInput.checked)
  );
  selectAllTh.appendChild(selectAllInput);
  selectAllCheckboxEl = selectAllInput;
  headerRow.appendChild(selectAllTh);

  for (let c = 0; c < maxCols; c++) {
    const th = document.createElement("th");
    const headerLabel = normalizeHeaderLabel(
      headerRowValues[c],
      columnLetter(c)
    );
    th.textContent = headerLabel;
    headerRow.appendChild(th);
  }

  const actionsTh = document.createElement("th");
  actionsTh.textContent = "AÇÕES";
  actionsTh.classList.add("actions-col");
  headerRow.appendChild(actionsTh);

  thead.appendChild(headerRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  dataRows.forEach((row, dataIndex) => {
    const tr = document.createElement("tr");
    const rowHeader = document.createElement("th");
    rowHeader.className = "row-index";
    rowHeader.textContent = dataIndex + 2; // numeração real da planilha
    tr.appendChild(rowHeader);

    const payload = buildPayloadFromRow(headerRowValues, row) || {};
    const rowKey = getRowKey(payload, dataIndex + 2);
    renderedRowMap.set(rowKey, payload);

    const selectTd = document.createElement("td");
    selectTd.className = "select-cell";
    const selectCheckbox = document.createElement("input");
    selectCheckbox.type = "checkbox";
    selectCheckbox.title = "Selecionar produto";
    selectCheckbox.addEventListener("change", () =>
      toggleRowSelection(rowKey, selectCheckbox.checked, tr)
    );
    selectTd.appendChild(selectCheckbox);
    rowCheckboxes.push({ key: rowKey, checkbox: selectCheckbox, tr });
    tr.appendChild(selectTd);

    let metaForRow = null;

    for (let c = 0; c < maxCols; c++) {
      const td = document.createElement("td");
      const value = row[c];
      const headerName = (headerRowValues[c] || "").toString().trim().toLowerCase();
      const keepRaw = headerName === "id" || headerName === "quantidade";
      const isBoolField =
        headerName === "troca" ||
        headerName === "rotatividade_alta" ||
        headerName === "rotatividade alta" ||
        headerName === "rotatividade";
      if (isBoolField) {
        td.textContent = booleanLabel(value, false);
      } else {
        td.textContent = keepRaw ? value ?? "" : formatCellValue(value);
      }

      if (!metaForRow) {
        const header = (headerRowValues[c] || "").toString().toLowerCase().trim();
        if (header === "validade" || header === "vencimento" || header === "data") {
          metaForRow = getExpiryMeta(row[c]);
        }
      }

      if (lancadoColumnIndex === c) {
        const lancadoMeta = getLancadoMeta(value);
        td.className = `status-pill ${lancadoMeta.color}`;
        td.textContent = lancadoMeta.text;
      }

      if (skuColumnIndex !== -1 && c === skuColumnIndex) {
        td.classList.add("sku-cell");
        td.title = "Clique para copiar o SKU";
        td.addEventListener("click", () => {
          const text = (td.textContent || "").trim();
          if (!text) return;
          copyTextToClipboard(text);
          statusEl.textContent = `SKU copiado: ${text}`;
        });
      }
      if (eanColumnIndex !== -1 && c === eanColumnIndex) {
        td.style.cursor = "pointer";
        td.title = "Clique para copiar EAN";
        td.addEventListener("click", () => {
          const text = (td.textContent || "").trim();
          if (!text) return;
          copyTextToClipboard(text);
          statusEl.textContent = `EAN copiado: ${text}`;
        });
      }
      tr.appendChild(td);
    }

    if (metaForRow) {
      tr.classList.add(metaForRow.color);
    }

    const actionsTd = document.createElement("td");
    actionsTd.style.whiteSpace = "nowrap";
    actionsTd.classList.add("actions-cell");

    const btnEdit = document.createElement("button");
    btnEdit.type = "button";
    btnEdit.className = "btn btn-secondary btn-sm";
    btnEdit.textContent = "Editar";
    btnEdit.addEventListener("click", () => {
      if (!payload) return;
      openEditModal(payload);
    });

    const btnDelete = document.createElement("button");
    btnDelete.type = "button";
    btnDelete.className = "btn btn-error btn-sm";
    btnDelete.style.marginLeft = "6px";
    btnDelete.textContent = "Eliminar";
    btnDelete.addEventListener("click", async () => {
      if (!payload) return;
      await deleteTargets([payload]);
    });

    const btnLancado = document.createElement("button");
    btnLancado.type = "button";
    btnLancado.className = "btn btn-success btn-sm";
    btnLancado.style.marginLeft = "6px";
    btnLancado.textContent = "Lançar";
    btnLancado.addEventListener("click", async () => {
      if (!payload) return;
      await launchTargets([payload]);
    });

    actionsTd.appendChild(btnEdit);
    actionsTd.appendChild(btnDelete);
    actionsTd.appendChild(btnLancado);

    tr.appendChild(actionsTd);

    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  output.innerHTML = "";
  output.appendChild(table);
  updateBulkActionsUI();
}

if (bulkLaunchBtn) {
  bulkLaunchBtn.addEventListener("click", () =>
    launchTargets(getSelectedPayloads())
  );
}

if (bulkDeleteBtn) {
  bulkDeleteBtn.addEventListener("click", () =>
    deleteTargets(getSelectedPayloads())
  );
}

if (bulkEditBtn) {
  bulkEditBtn.addEventListener("click", () =>
    openEditModal(getSelectedPayloads())
  );
}

function buildPayloadFromRow(headers, row) {
  if (!headers || !row) return null;
  const lower = headers.map((h) => String(h || "").trim().toLowerCase());
  const get = (key) => {
    const idx = lower.indexOf(key);
    if (idx === -1) return "";
    return row[idx] ?? "";
  };
  const findIndex = (keys = []) => {
    for (const key of keys) {
      const idx = lower.indexOf(key);
      if (idx !== -1) return idx;
    }
    return -1;
  };
  const quantidadeIdx = findIndex(["quantidade"]);
  const lancadoIdx = findIndex(["lançado", "lancado"]);
  const dataLancadoIdx = findIndex(["data_lançado", "data_lancado", "datalancado"]);
  const rotatividadeIdx = findIndex(["rotatividade_alta", "rotatividade alta", "rotatividade"]);

  const getByIdx = (idx) => {
    if (idx === -1) return "";
    return row[idx] ?? "";
  };

  const quantidadeRaw = getByIdx(quantidadeIdx);
  const lancadoRaw = getByIdx(lancadoIdx);
  const dataLancado = getByIdx(dataLancadoIdx);
  const rotatividadeRaw = getByIdx(rotatividadeIdx);

  const qty = Number.isFinite(Number(quantidadeRaw))
    ? Number(quantidadeRaw)
    : 0;
  const lancado = toBoolean(lancadoRaw, false);
  return {
    id: get("id"),
    ean: get("ean"),
    nome: String(get("nome") || "").trim(),
    quantidade: qty,
    quantidade_text:
      quantidadeRaw !== undefined && quantidadeRaw !== "" ? String(quantidadeRaw) : String(qty),
    validade: get("validade"),
    troca: toBoolean(get("troca"), false),
    rotatividade_alta: toBoolean(rotatividadeRaw, false),
    lancado,
    data_lancado: dataLancado,
  };
}

function openEditModal(payload) {
  const targets = Array.isArray(payload)
    ? payload.filter(Boolean)
    : [payload].filter(Boolean);
  if (!targets.length) {
    statusEl.textContent = "Selecione ao menos um produto para editar.";
    return;
  }
  currentEditTargets = targets;
  const isBulk = targets.length > 1;
  const first = targets[0];
  if (editTitle) {
    editTitle.textContent = isBulk
      ? `Editar ${targets.length} registros`
      : "Editar registro";
  }
  if (editHint) {
    editHint.textContent = isBulk
      ? "Campos vazios não serão alterados nos itens selecionados."
      : "Edite os campos e clique em salvar.";
  }
  editId.value = isBulk ? `${targets.length} selecionados` : first.id || "";
  editEan.value = isBulk ? "" : first.ean || "";
  editEan.placeholder = isBulk ? "Deixe em branco para manter" : "";
  editNome.value = isBulk ? "" : first.nome || "";
  editNome.placeholder = isBulk ? "Deixe em branco para manter" : "";
  editQtd.value = isBulk
    ? ""
    : first.quantidade ?? first.quantidade_text ?? "";
  editValidade.value = isBulk ? "" : first.validade || "";
  const trocaValue = toBoolean(first.troca, false);
  const rotaValue = toBoolean(first.rotatividade_alta, false);
  if (editTroca) {
    editTroca.indeterminate = isBulk;
    editTroca.checked = isBulk ? false : trocaValue;
    updateSwitchLabel(editTroca, editTrocaLabel);
  }
  if (editRota) {
    editRota.indeterminate = isBulk;
    editRota.checked = isBulk ? false : rotaValue;
    updateSwitchLabel(editRota, editRotaLabel);
  }
  editStatus.textContent = "";
  editStatus.className = "status";
  editModal.style.display = "flex";
}

function closeEditModal() {
  editModal.style.display = "none";
  editStatus.textContent = "";
  editStatus.className = "status";
  currentEditTargets = [];
}

function copyTextToClipboard(text) {
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).catch(() => fallbackCopy(text));
  } else {
    fallbackCopy(text);
  }
}

function fallbackCopy(text) {
  const temp = document.createElement("textarea");
  temp.value = text;
  document.body.appendChild(temp);
  temp.select();
  try {
    document.execCommand("copy");
  } catch (e) {}
  document.body.removeChild(temp);
}

// ----- Cadastro de produto -----
const formPanel = document.getElementById("form-panel");
const openFormBtn = document.getElementById("open-form-btn");
const produtoForm = document.getElementById("produto-form");
const mensagem = document.getElementById("mensagem");
const itensContainer = document.getElementById("itens-container");
const enviarTodosBtn = document.getElementById("enviar-todos-btn");
let itens = [];

function ensureFormPanelVisible() {
  if (formPanel) formPanel.style.display = "block";
  if (barcodePanel) barcodePanel.style.display = "block";
  if (openFormBtn) openFormBtn.textContent = "Cadastrar produto";
}

function ensureAtLeastOneItem() {
  if (!itens.length) {
    itens.push(createItemFromCode("", ""));
  }
  renderItens();
}

openFormBtn.addEventListener("click", () => {
  ensureFormPanelVisible();
  formPanel?.scrollIntoView({ behavior: "smooth", block: "start" });
});

produtoForm.addEventListener("submit", (event) => event.preventDefault());

function createItemFromCode(codigo, nome = "") {
  return {
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    codigo,
    nome,
    quantidade: "",
    validade: "",
    troca: false,
    rotatividade_alta: false,
    status: "",
    sent: false,
    allowEditCode: !codigo,
    lastLookupEan: "",
  };
}

function focusItemCode(item) {
  if (!item) return;
  const card = document.querySelector(`[data-id="${item.id}"]`);
  const codeInput = card?.querySelector(".field input[type='text']");
  if (codeInput) {
    codeInput.value = item.codigo || "";
    codeInput.readOnly = !item.allowEditCode && Boolean(item.codigo);
    codeInput.focus();
  }
}

function applyProductInfoToItem(item, product) {
  if (!item || !product) return;
  if (product.nome && !item.nome) item.nome = product.nome;
  if (typeof product.troca === "boolean") item.troca = product.troca;
  if (typeof product.rotatividade_alta === "boolean") {
    item.rotatividade_alta = product.rotatividade_alta;
  }
}

async function lookupAndFillItemByEan(item, ean, controls = {}) {
  const { nomeInput, trocaSwitch, rotaSwitch, statusEl } = controls;
  if (!ean || ean.length !== 13) return;
  if (item.lastLookupEan === ean) return;
  item.lastLookupEan = ean;
  if (statusEl) setStatus(statusEl, "Consultando produto...", "");
  const product = await fetchProductFromAction(ean);
  if (!product) {
    if (statusEl) setStatus(statusEl, "Produto não encontrado.", "");
    return;
  }
  applyProductInfoToItem(item, product);
  if (nomeInput) nomeInput.value = item.nome || "";
  if (trocaSwitch?.input) {
    trocaSwitch.input.checked = !!item.troca;
    updateSwitchLabel(trocaSwitch.input, trocaSwitch.text);
  }
  if (rotaSwitch?.input) {
    rotaSwitch.input.checked = !!item.rotatividade_alta;
    updateSwitchLabel(rotaSwitch.input, rotaSwitch.text);
  }
  if (statusEl) setStatus(statusEl, "Produto encontrado.", "ok");
}

function createSwitch(initialValue, onChange) {
  const wrapper = document.createElement("label");
  wrapper.className = "switch";
  const input = document.createElement("input");
  input.type = "checkbox";
  input.checked = !!initialValue;
  const slider = document.createElement("span");
  slider.className = "slider";
  const text = document.createElement("span");
  text.className = "switch-label";
  const syncLabel = () => {
    text.textContent = input.checked ? "SIM" : "NÃO";
  };
  syncLabel();
  input.addEventListener("change", () => {
    syncLabel();
    onChange?.(input.checked);
  });
  wrapper.appendChild(input);
  wrapper.appendChild(slider);
  wrapper.appendChild(text);
  return { wrapper, input, text };
}

function renderItens() {
  itensContainer.innerHTML = "";
  if (!itens.length) {
    itensContainer.className = "itens-container extra-empty";
    itensContainer.textContent = "Nenhum produto adicionado. Escaneie ou adicione manualmente.";
    return;
  }
  itensContainer.className = "itens-container";

  itens.forEach((item, index) => {
    const card = document.createElement("div");
    card.className = "item-card";
    card.dataset.id = item.id;

    const header = document.createElement("div");
    header.className = "item-header";
    const title = document.createElement("div");
    title.innerHTML = `Produto ${index + 1}`;
    const actions = document.createElement("div");
    actions.className = "item-actions";

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "btn btn-secondary btn-sm";
    removeBtn.textContent = "Remover local";
    removeBtn.addEventListener("click", () => {
      itens = itens.filter((it) => it.id !== item.id);
      renderItens();
    });

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "btn btn-error btn-sm";
    deleteBtn.textContent = "Eliminar remoto";
    deleteBtn.addEventListener("click", async () => {
      const password = requestDeletePassword();
      if (!password) {
        mensagem.textContent = "Eliminação cancelada.";
        mensagem.className = "erro";
        return;
      }
      try {
        await postJsonWithFallback(getEndpoints("actions"), {
          action: "apagar",
          id: item.codigo || item.id || "",
          password,
          ...getUserPayload(),
        });
        mensagem.textContent = "Solicitado eliminar produto no servidor.";
        mensagem.className = "ok";
      } catch (err) {
        mensagem.textContent = `Erro ao eliminar remoto: ${err.message || err}`;
        mensagem.className = "erro";
      }
    });

    actions.appendChild(removeBtn);
    actions.appendChild(deleteBtn);
    header.appendChild(title);
    header.appendChild(actions);
    card.appendChild(header);

    const codeField = document.createElement("div");
    codeField.className = "field";
    const codeLabel = document.createElement("label");
    codeLabel.textContent = "Código de barras (EAN)";
    const codeInput = document.createElement("input");
    codeInput.type = "text";
    codeInput.value = item.codigo || "";
    codeInput.readOnly = !item.allowEditCode && Boolean(item.codigo);
    codeInput.required = true;
    codeInput.placeholder = "Digite ou escaneie o EAN";
    codeInput.addEventListener("input", (e) => {
      const digits = (e.target.value || "").replace(/\D/g, "");
      e.target.value = digits;
      item.codigo = digits;
      if (digits.length === 13) {
        lookupAndFillItemByEan(item, digits, {
          nomeInput,
          trocaSwitch,
          rotaSwitch,
          statusEl: status,
        });
      }
    });
    codeField.appendChild(codeLabel);
    codeField.appendChild(codeInput);
    card.appendChild(codeField);

    const fields = document.createElement("div");
    fields.className = "item-fields";

    const nomeField = document.createElement("div");
    nomeField.className = "field";
    const nomeLabel = document.createElement("label");
    nomeLabel.textContent = "Nome do produto";
    const nomeInput = document.createElement("input");
    nomeInput.type = "text";
    nomeInput.value = item.nome || "";
    nomeInput.required = true;
    nomeInput.placeholder = "Obrigatório";
    nomeInput.addEventListener("input", (e) => {
      item.nome = e.target.value;
    });
    nomeField.appendChild(nomeLabel);
    nomeField.appendChild(nomeInput);

    const qtdField = document.createElement("div");
    qtdField.className = "field";
    const qtdLabel = document.createElement("label");
    qtdLabel.textContent = "Quantidade";
    const qtdInput = document.createElement("input");
    qtdInput.type = "number";
    qtdInput.min = "1";
    qtdInput.step = "1";
    qtdInput.required = true;
    qtdInput.placeholder = "Ex.: 5";
    qtdInput.value = item.quantidade || "";
    qtdInput.addEventListener("input", (e) => {
      const digits = (e.target.value || "").replace(/\D/g, "");
      e.target.value = digits;
      item.quantidade = digits ? Number(digits) : "";
    });
    qtdField.appendChild(qtdLabel);
    qtdField.appendChild(qtdInput);

    const validadeField = document.createElement("div");
    validadeField.className = "field";
    const validadeLabel = document.createElement("label");
    validadeLabel.textContent = "Validade";
    const validadeInput = document.createElement("input");
    validadeInput.type = "date";
    validadeInput.required = true;
    validadeInput.value = item.validade || "";
    validadeInput.addEventListener("input", (e) => {
      item.validade = e.target.value;
    });
    validadeField.appendChild(validadeLabel);
    validadeField.appendChild(validadeInput);

    const trocaField = document.createElement("div");
    trocaField.className = "field";
    const trocaLabel = document.createElement("label");
    trocaLabel.textContent = "Troca";
    const trocaSwitch = createSwitch(item.troca, (checked) => {
      item.troca = checked;
    });
    trocaField.appendChild(trocaLabel);
    trocaField.appendChild(trocaSwitch.wrapper);

    const rotaField = document.createElement("div");
    rotaField.className = "field";
    const rotaLabel = document.createElement("label");
    rotaLabel.textContent = "Rotatividade alta";
    const rotaSwitch = createSwitch(item.rotatividade_alta, (checked) => {
      item.rotatividade_alta = checked;
    });
    rotaField.appendChild(rotaLabel);
    rotaField.appendChild(rotaSwitch.wrapper);

    fields.appendChild(nomeField);
    fields.appendChild(qtdField);
    fields.appendChild(validadeField);
    fields.appendChild(trocaField);
    fields.appendChild(rotaField);
    card.appendChild(fields);

    const status = document.createElement("div");
    status.className = "item-status";
    status.textContent = item.status || "";
    card.appendChild(status);

    itensContainer.appendChild(card);
  });
}

function setStatus(el, text, type = "") {
  if (!el) return;
  el.textContent = text;
  el.className = "item-status";
  if (type === "erro") el.classList.add("erro");
  if (type === "ok") el.classList.add("ok");
}

function validateItem(item) {
  if (!item.codigo) return "Escaneie o código de barras.";
  if (!item.nome || !item.nome.trim()) return "Informe o nome do produto.";
  const qtd = Number(item.quantidade);
  if (!item.quantidade || Number.isNaN(qtd) || qtd <= 0)
    return "Quantidade deve ser maior que zero.";
  if (!item.validade) return "Informe a data de validade.";
  return "";
}

async function enviarItem(item, index, statusEl) {
  const qty = Number.isFinite(Number(item.quantidade)) ? Number(item.quantidade) : 0;
  const payload = {
    action: "insert",
    id: item.id || item.codigo || "",
    ean: item.codigo,
    nome: item.nome.trim(),
    quantidade: qty,
    quantidade_text: String(qty),
    validade: item.validade,
    troca: !!item.troca,
    rotatividade_alta: !!item.rotatividade_alta,
    ...getUserPayload(),
  };

  await postJsonWithFallback(getEndpoints("produto"), payload);
}

async function enviarTodos() {
  if (!itens.length) {
    mensagem.textContent = "Adicione ao menos um produto.";
    mensagem.className = "erro";
    return;
  }
  mensagem.textContent = "";
  mensagem.className = "";

  for (let i = 0; i < itens.length; i++) {
    const item = itens[i];
    const statusEl = document.querySelector(
      `[data-id="${item.id}"] .item-status`
    );
    const error = validateItem(item);
    if (error) {
      setStatus(statusEl, error, "erro");
      continue;
    }
    setStatus(statusEl, "Enviando...");
    try {
      await enviarItem(item, i, statusEl);
      setStatus(statusEl, "Enviado com sucesso.", "ok");
      item.sent = true;
    } catch (err) {
      console.error("Erro ao enviar produto:", err);
      setStatus(
        statusEl,
        "Erro ao enviar produto. Tente novamente.",
        "erro"
      );
    }
  }

  const enviados = itens.filter((item) => item.sent).length;
  itens = itens.filter((item) => !item.sent);
  renderItens();
  if (enviados) {
    mensagem.textContent = `${enviados} produto(s) enviado(s).`;
    mensagem.className = "ok";
    setTimeout(() => {
      fetchBtn.click();
    }, 1000);
  }
}

enviarTodosBtn.addEventListener("click", enviarTodos);

// Modal edição
editCancelBtn.addEventListener("click", closeEditModal);
editSaveBtn.addEventListener("click", async () => {
  const eanValue = (editEan.value || "").trim();
  const nomeValue = (editNome.value || "").trim();
  const validadeValue = editValidade.value;
  const qtdRaw = (editQtd.value || "").trim();
const targets = currentEditTargets.length
  ? currentEditTargets
  : [
      {
        id: editId.value || "",
          ean: eanValue,
          nome: nomeValue,
          quantidade: Number(qtdRaw || 0),
          quantidade_text: String(qtdRaw || ""),
          validade: validadeValue,
          troca: !!editTroca?.checked,
          lancado: false,
          rotatividade_alta: !!editRota?.checked,
          data_lancado: "",
        },
      ];
    const isBulk = targets.length > 1;

  if (!isBulk && (!eanValue || !nomeValue)) {
    editStatus.textContent = "Preencha EAN e nome.";
    editStatus.className = "status err";
    return;
  }
  if (isBulk) {
    const hasChanges =
      Boolean(eanValue || nomeValue || validadeValue || qtdRaw) ||
      (editTroca && !editTroca.indeterminate) ||
      (editRota && !editRota.indeterminate);
    if (!hasChanges) {
      editStatus.textContent =
        "Preencha pelo menos um campo para aplicar nos selecionados.";
      editStatus.className = "status err";
      return;
    }
  }

  editStatus.textContent = "Enviando alterações...";
  editStatus.className = "status";

  try {
    for (const target of targets) {
      const trocaParsed =
        editTroca && editTroca.indeterminate
          ? toBoolean(target.troca, false)
          : !!editTroca?.checked;
      const rotaParsed =
        editRota && editRota.indeterminate
          ? toBoolean(target.rotatividade_alta, false)
          : !!editRota?.checked;
      const quantidadeAtual =
        qtdRaw === "" ? target.quantidade : Number(qtdRaw || 0);
      const lancadoAtual = toBoolean(target.lancado, false);
      const payload = {
        action: "editar",
        id: target.id || "",
        ean: eanValue || target.ean,
        nome: nomeValue || target.nome,
        quantidade: quantidadeAtual,
        quantidade_text:
          qtdRaw === ""
            ? target.quantidade_text ?? String(target.quantidade ?? "")
            : String(qtdRaw || ""),
        validade: validadeValue || target.validade,
        troca: trocaParsed,
        lancado: lancadoAtual,
        rotatividade_alta: rotaParsed,
        data_lancado: target.data_lancado || "",
        ...getUserPayload(),
      };
      if (!payload.ean || !payload.nome) {
        editStatus.textContent = "EAN e nome são obrigatórios.";
        editStatus.className = "status err";
        return;
      }
      await postJsonWithFallback(getEndpoints("actions"), payload);
    }
    editStatus.textContent = isBulk
      ? "Registros atualizados."
      : "Registro enviado.";
    editStatus.className = "status ok";
    clearSelections();
    closeEditModal();
    fetchBtn.click();
  } catch (err) {
    editStatus.textContent = `Erro ao salvar: ${err.message || err}`;
    editStatus.className = "status err";
  }
});

ensureAtLeastOneItem();

if (isAuthenticated && fetchBtn) {
  fetchBtn.click();
}
