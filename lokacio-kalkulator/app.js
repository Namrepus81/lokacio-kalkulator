const SOURCE_REQUIRED_COLUMNS = [
  "cikkszam",
  "lokacio",
];

const MASTER_FIELDS = [
  "min",
  "max",
];

const STORAGE_KEYS = {
  masterRows: "raktar-asszisztens-master-rows-v1",
  masterStatus: "raktar-asszisztens-master-status-v1",
};

const SHARED_EXPORT_URLS = ["data/EXPORT.XLSX", "data/EXPORT.xlsx"];

const SAMPLE_ROWS = [
  {
    cikkszam: "1013818",
    megnevezes: "Minta csavar",
    lokacio: "A-A-1-1",
    min: 500,
    max: 3000,
    forgas: "sűrű",
    tarolo: "kis KLT",
    suly_kategoria: "könnyű",
    kapacitas: 1200,
    megjegyzes: "jó helyen van",
  },
  {
    cikkszam: "1013818",
    megnevezes: "Minta csavar",
    lokacio: "A-A-1-2",
    min: 500,
    max: 3000,
    forgas: "sűrű",
    tarolo: "kis KLT",
    suly_kategoria: "könnyű",
    kapacitas: 800,
    megjegyzes: "második alsó szint",
  },
  {
    cikkszam: "1013818",
    megnevezes: "Minta csavar",
    lokacio: "A-A-1-3",
    min: 500,
    max: 3000,
    forgas: "sűrű",
    tarolo: "kis KLT",
    suly_kategoria: "könnyű",
    kapacitas: 1500,
    megjegyzes: "felső tartalékhely",
  },
  {
    cikkszam: "1022455",
    megnevezes: "Minta fedel",
    lokacio: "C-A-1-2",
    min: 1000,
    max: 6000,
    forgas: "sűrű",
    tarolo: "XL láda",
    suly_kategoria: "nehéz",
    kapacitas: 2000,
    megjegyzes: "rossz hely minta",
  },
  {
    cikkszam: "1088891",
    megnevezes: "Minta konzol",
    lokacio: "B-A-1-1",
    min: 200,
    max: 1500,
    forgas: "lassú",
    tarolo: "nagy KLT",
    suly_kategoria: "közepes",
    kapacitas: 1800,
    megjegyzes: "lassú anyag túl jó helyen",
  },
  {
    cikkszam: "1099911",
    megnevezes: "Hiányos adat",
    lokacio: "B-A-1-4",
    min: "",
    max: 2200,
    forgas: "közepes",
    tarolo: "raklap",
    suly_kategoria: "nehéz",
    kapacitas: 1200,
    megjegyzes: "hiányzó minimum",
  },
];

let rows = [];
let analyzedRows = [];
let importRows = [];
let masterRowsByItem = new Map();

const fileInput = document.querySelector("#fileInput");
const sharedExportButton = document.querySelector("#sharedExportButton");
const masterFileInput = document.querySelector("#masterFileInput");
const pickingPhotoInput = document.querySelector("#pickingPhotoInput");
const quickPickingInput = document.querySelector("#quickPickingInput");
const buildQuickPickingButton = document.querySelector("#buildQuickPickingButton");
const searchForm = document.querySelector("#searchForm");
const searchInput = document.querySelector("#searchInput");
const topSearchInput = document.querySelector("#topSearchInput");
const menuButton = document.querySelector("#menuButton");
const sidebarBackdrop = document.querySelector("#sidebarBackdrop");
const sideNav = document.querySelector(".side-nav");
const itemSearchMode = document.querySelector("#itemSearchMode");
const locationSearchMode = document.querySelector("#locationSearchMode");
const searchTitle = document.querySelector("#searchTitle");
const resultBox = document.querySelector("#resultBox");
const dataStatus = document.querySelector("#dataStatus");
const itemCount = document.querySelector("#itemCount");
const warningCount = document.querySelector("#warningCount");
const errorCount = document.querySelector("#errorCount");
const issuesList = document.querySelector("#issuesList");
const issueFilter = document.querySelector("#issueFilter");
const dataTableBody = document.querySelector("#dataTableBody");
const loadSampleButton = document.querySelector("#loadSampleButton");
const logicButton = document.querySelector("#logicButton");
const logicList = document.querySelector("#logicList");
const masterStatus = document.querySelector("#masterStatus");
const missingMasterList = document.querySelector("#missingMasterList");
const scanButton = document.querySelector("#scanButton");
const scannerModal = document.querySelector("#scannerModal");
const scannerVideo = document.querySelector("#scannerVideo");
const scannerStatus = document.querySelector("#scannerStatus");
const closeScannerButton = document.querySelector("#closeScannerButton");
const scanConfirm = document.querySelector("#scanConfirm");
const scannedCodeText = document.querySelector("#scannedCodeText");
const useScannedCodeButton = document.querySelector("#useScannedCodeButton");
const retryScanButton = document.querySelector("#retryScanButton");
const pickingStatus = document.querySelector("#pickingStatus");
const pickingList = document.querySelector("#pickingList");
const ocrTextBox = document.querySelector("#ocrTextBox");

let scannerStream = null;
let scannerDetector = null;
let scannerActive = false;
let pendingScannedCode = "";
let searchMode = "item";
let pickingRows = [];

loadSavedMasterRows();

fileInput.addEventListener("change", handleFileChange);
sharedExportButton.addEventListener("click", loadSharedExport);
masterFileInput.addEventListener("change", handleMasterFileChange);
pickingPhotoInput.addEventListener("change", handlePickingPhotoChange);
buildQuickPickingButton.addEventListener("click", buildQuickPickingList);
searchForm.addEventListener("submit", handleSearch);
pickingList.addEventListener("change", (event) => {
  const checkbox = event.target.closest("[data-picking-id]");
  if (checkbox) {
    togglePickingDone(checkbox.dataset.pickingId, checkbox.checked);
    return;
  }

  const input = event.target.closest("[data-picking-field]");
  if (input) {
    updatePickingField(input.dataset.pickingRowId, input.dataset.pickingField, input.value);
  }
});
pickingList.addEventListener("input", (event) => {
  const input = event.target.closest("[data-picking-field]");
  if (!input) return;
  updatePickingField(input.dataset.pickingRowId, input.dataset.pickingField, input.value, false);
});
pickingList.addEventListener("click", (event) => {
  const removeButton = event.target.closest("[data-remove-picking-id]");
  if (!removeButton) return;
  removePickingRow(removeButton.dataset.removePickingId);
});
topSearchInput.addEventListener("input", () => {
  searchInput.value = topSearchInput.value;
});
searchInput.addEventListener("input", () => {
  topSearchInput.value = searchInput.value;
});
itemSearchMode.addEventListener("click", () => setSearchMode("item"));
locationSearchMode.addEventListener("click", () => setSearchMode("location"));
issueFilter.addEventListener("input", renderIssues);
loadSampleButton.addEventListener("click", () => setRows(SAMPLE_ROWS, "Minta adatok betöltve"));
logicButton.addEventListener("click", renderLogicalPlacement);
scanButton.addEventListener("click", startScanner);
closeScannerButton.addEventListener("click", stopScanner);
useScannedCodeButton.addEventListener("click", confirmScannedCode);
retryScanButton.addEventListener("click", retryScanner);
menuButton?.addEventListener("click", toggleMobileMenu);
sidebarBackdrop?.addEventListener("click", closeMobileMenu);
sideNav?.addEventListener("click", (event) => {
  if (event.target.closest("a")) closeMobileMenu();
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeMobileMenu();
});

function toggleMobileMenu() {
  const isOpen = document.body.classList.toggle("menu-open");
  menuButton?.setAttribute("aria-expanded", String(isOpen));
  menuButton?.setAttribute("aria-label", isOpen ? "Menü bezárása" : "Menü megnyitása");
}

function closeMobileMenu() {
  document.body.classList.remove("menu-open");
  menuButton?.setAttribute("aria-expanded", "false");
  menuButton?.setAttribute("aria-label", "Menü megnyitása");
}

function handleFileChange(event) {
  const file = event.target.files[0];
  if (!file) {
    dataStatus.textContent = "Nem lett fájl kiválasztva.";
    return;
  }
  dataStatus.textContent = `Napi export kiválasztva: ${file.name}. Olvasás indul...`;
  readUploadedRows(file, (parsedRows, sheetName) => setRows(parsedRows, `${file.name} / ${sheetName}`), "source");
}

async function loadSharedExport() {
  if (!window.XLSX) {
    showImportError("Az Excel olvasó nem töltődött be. Frissítsd az oldalt, vagy próbáld meg később.");
    return;
  }

  sharedExportButton.disabled = true;
  dataStatus.textContent = "GitHub export betöltése indul...";

  try {
    const { response, url } = await fetchFirstAvailableExport();

    const buffer = await response.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const firstSheetName = chooseWorkbookSheet(workbook, "source");
    const sheet = workbook.Sheets[firstSheetName];
    const parsedRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    setRows(parsedRows, `GitHub export (${url}) / ${firstSheetName}`);
    importCompanionMasterSheet(workbook, "GitHub export", firstSheetName);
  } catch (error) {
    showImportError(`GitHub export nem tölthető be: ${error.message}. Ellenőrizd, hogy létezik-e valamelyik: ${SHARED_EXPORT_URLS.join(", ")}`);
  } finally {
    sharedExportButton.disabled = false;
  }
}

async function fetchFirstAvailableExport() {
  const cacheBuster = `v=${Date.now()}`;
  const errors = [];

  for (const url of SHARED_EXPORT_URLS) {
    const separator = url.includes("?") ? "&" : "?";
    const response = await fetch(`${url}${separator}${cacheBuster}`, { cache: "no-store" });
    if (response.ok) return { response, url };
    errors.push(`${url}: ${response.status}`);
  }

  throw new Error(`nem elérhető (${errors.join("; ")})`);
}

function handleMasterFileChange(event) {
  const file = event.target.files[0];
  if (!file) {
    masterStatus.textContent = "Nem lett törzsadat fájl kiválasztva.";
    return;
  }
  masterStatus.textContent = `Törzsadat kiválasztva: ${file.name}. Olvasás indul...`;
  readUploadedRows(file, (parsedRows, sheetName) => setMasterRows(parsedRows, `${file.name} / ${sheetName}`), "master");
}

function buildQuickPickingList() {
  const lines = quickPickingInput.value
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (!lines.length) {
    pickingStatus.textContent = "Írj be legalább egy sort: cikkszám mennyiség.";
    return;
  }

  const builtRows = [];
  const missingRows = [];

  lines.forEach((line, index) => {
    const parsed = parseQuickPickingLine(line);
    if (!parsed) {
      missingRows.push(`${line} -> nem értelmezhető`);
      return;
    }

    const itemRows = findRowsForPickingSku(parsed.sku);
    if (!itemRows.length) {
      missingRows.push(`${parsed.sku} -> nincs lokáció a feltöltött Excelben`);
      return;
    }

    const pickingLocations = selectPickingRowsForItem(itemRows, parsed.quantity);

    if (!pickingLocations.length) {
      missingRows.push(`${parsed.sku} -> nincs kiadható lokáció vagy készletfedezet`);
      return;
    }

    const requested = Number(toNumber(parsed.quantity));
    const covered = pickingLocations.reduce((sum, pick) => sum + (Number(toNumber(pick.quantity)) || 0), 0);
    if (Number.isFinite(requested) && requested > 0 && covered > 0 && covered < requested) {
      missingRows.push(`${parsed.sku} -> nincs elég fedezet, hiány: ${formatPickingQuantity(requested - covered)}`);
    }

    pickingLocations.forEach((pick, rowIndex) => {
      const row = pick.row;
      builtRows.push({
        id: `quick-${parsed.sku}-${index}-${rowIndex}`,
        raw: line,
        location: row.lokacio || "",
        sku: row.cikkszam,
        description: row.megnevezes || "Nincs megnevezés",
        quantity: pick.quantity,
        statusText: pick.note,
        done: false,
      });
    });
  });

  pickingRows = builtRows.sort(comparePickingRows);
  ocrTextBox.textContent = missingRows.join("\n");
  renderPickingList();

  if (missingRows.length) {
    pickingStatus.textContent += ` Nem talált/hibás sor: ${missingRows.length}. Részletek az OCR nyers szöveg alatt.`;
  }
}

function parseQuickPickingLine(line) {
  const skuMatch = line.match(/\b[A-Z0-9][A-Z0-9-]{3,}\b/i);
  const numberMatches = line.match(/\b\d+(?:[,.]\d+)?\b/g) || [];
  if (!skuMatch || !numberMatches.length) return null;

  const sku = skuMatch[0].toUpperCase();
  const quantity = numberMatches[numberMatches.length - 1].replace(",", ".");
  return { sku, quantity };
}

function findRowsForPickingSku(sku) {
  const normalizedSku = stringify(sku).toLowerCase();
  const candidates = analyzedRows.filter((row) => row.cikkszam.toLowerCase() === normalizedSku);
  return candidates.length ? candidates : rows.filter((row) => row.cikkszam.toLowerCase() === normalizedSku);
}

function selectPickingRowsForItem(itemRows, requestedQuantity) {
  const requested = Number(toNumber(requestedQuantity));
  const lowerRows = itemRows
    .filter((row) => Number(row.szint) <= 2)
    .sort(compareRowsByPickingRoute);
  const upperRows = itemRows
    .filter((row) => Number(row.szint) >= 3)
    .sort(compareRowsByPickingRoute);

  if (!Number.isFinite(requested) || requested <= 0) {
    return lowerRows.map((row) => ({
      row,
      quantity: requestedQuantity,
      note: "Alsó szint, mennyiség ellenőrzés nélkül",
    }));
  }

  const hasStockData = itemRows.some((row) => row.mennyiseg !== "");

  if (!hasStockData) {
    return lowerRows.map((row, index) => ({
      row,
      quantity: index === 0 ? requestedQuantity : "",
      note: "Alsó szint, nincs készletadat a felső szint számításához",
    }));
  }

  let remaining = requested;
  const selected = [];

  lowerRows.forEach((row) => {
    if (remaining <= 0) return;
    const available = Math.max(0, Number(row.mennyiseg) || 0);
    if (available <= 0) return;
    const quantity = Math.min(available, remaining);
    selected.push({
      row,
      quantity: formatPickingQuantity(quantity),
      note: "Alsó szint",
    });
    remaining -= quantity;
  });

  if (remaining <= 0) return selected;

  upperRows.forEach((row) => {
    if (remaining <= 0) return;
    const available = Math.max(0, Number(row.mennyiseg) || 0);
    if (available <= 0) return;
    const quantity = Math.min(available, remaining);
    selected.push({
      row,
      quantity: formatPickingQuantity(quantity),
      note: "Felső szint csak fedezetpótlás",
    });
    remaining -= quantity;
  });

  return selected;
}

function compareRowsByPickingRoute(left, right) {
  return comparePickingRows(
    { location: left.lokacio || "", sku: left.cikkszam || "" },
    { location: right.lokacio || "", sku: right.cikkszam || "" },
  );
}

function formatPickingQuantity(value) {
  return Number.isInteger(value) ? String(value) : String(Math.round(value * 1000) / 1000);
}

async function handlePickingPhotoChange(event) {
  const file = event.target.files[0];
  if (!file) {
    pickingStatus.textContent = "Nem lett fotó kiválasztva.";
    return;
  }

  if (!window.Tesseract) {
    pickingStatus.textContent = "Az OCR olvasó nem töltődött be. Frissítsd az oldalt, vagy nézd meg van-e internet.";
    return;
  }

  pickingStatus.textContent = `Fotó kiválasztva: ${file.name}. OCR olvasás indul...`;
  pickingList.innerHTML = `<div class="result-box empty">Olvasás folyamatban. Ez mobilon lehet 10-30 másodperc is.</div>`;
  ocrTextBox.textContent = "";

  try {
    const result = await Tesseract.recognize(file, "eng+hun", {
      logger: (progress) => {
        if (progress.status === "recognizing text") {
          pickingStatus.textContent = `OCR olvasás: ${Math.round(progress.progress * 100)}%`;
        }
      },
    });

    const text = result.data.text || "";
    ocrTextBox.textContent = text;
    pickingRows = parsePickingText(text).sort(comparePickingRows);
    renderPickingList();
  } catch (error) {
    pickingStatus.textContent = `Nem sikerült kiolvasni a fotót: ${error.message}`;
    pickingList.innerHTML = `<div class="result-box"><strong>OCR hiba:</strong> ${escapeHtml(error.message)}</div>`;
  }
}

function readUploadedRows(file, onRows, mode = "source") {
  const extension = file.name.split(".").pop().toLowerCase();
  const reader = new FileReader();

  reader.onerror = () => {
    showImportError(`Nem sikerült olvasni a fájlt: ${file.name}`);
  };

  reader.onload = (loadEvent) => {
    try {
      if (extension === "csv") {
        const text = loadEvent.target.result;
        onRows(parseCsv(text), "CSV");
        return;
      }

      if (!window.XLSX) {
        showImportError("Az Excel olvasó nem töltődött be mobilon. Frissítsd az oldalt, vagy próbáld CSV-ként mentve feltölteni.");
        return;
      }

      const workbook = XLSX.read(loadEvent.target.result, { type: "array" });
      if (mode === "master") {
        const fixedMasterRows = buildFixedLocationMasterRows(workbook);
        if (fixedMasterRows.length) {
          onRows(fixedMasterRows, "Fix tárhely + Munka1 min/max");
          return;
        }
      }

      const firstSheetName = chooseWorkbookSheet(workbook, mode);
      const sheet = workbook.Sheets[firstSheetName];
      const parsedRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      onRows(parsedRows, firstSheetName);

      if (mode === "source") {
        importCompanionMasterSheet(workbook, file.name, firstSheetName);
      }
    } catch (error) {
      showImportError(`Nem sikerült beolvasni a fájlt: ${error.message}`);
    }
  };

  if (extension === "csv") {
    reader.readAsText(file, "utf-8");
  } else {
    reader.readAsArrayBuffer(file);
  }
}

function buildFixedLocationMasterRows(workbook) {
  const locationSheetName = workbook.SheetNames.find((name) => normalizeText(name).includes("terhelyek") || normalizeText(name).includes("tarhelyek"));
  const minMaxSheetName = workbook.SheetNames.find((name) => normalizeText(name) === "munka1");
  if (!locationSheetName || !minMaxSheetName) return [];

  const locationRows = XLSX.utils.sheet_to_json(workbook.Sheets[locationSheetName], { defval: "" });
  const minMaxRows = XLSX.utils.sheet_to_json(workbook.Sheets[minMaxSheetName], { defval: "" });
  if (!locationRows.length || !minMaxRows.length) return [];

  const minMaxBySku = new Map();
  const issueQuantities = [];
  minMaxRows.forEach((row) => {
    const normalized = normalizeObjectKeys(row);
    const sku = stringify(normalized.cikkszam || normalized.termek || normalized.anyag);
    if (!sku) return;
    const issueQuantity = toNumber(normalized.kiadas_mennyiseg || normalized.kiadasi_gyakorisag || normalized.kiadas || normalized.forgas);
    if (issueQuantity !== "") issueQuantities.push(Number(issueQuantity));
    minMaxBySku.set(sku, {
      min: normalized.minimum_keszlet,
      max: normalized.maximum_keszlet,
      issueQuantity,
    });
  });
  const issueThresholds = getIssueQuantityThresholds(issueQuantities);

  return locationRows.map((row) => {
    const normalized = normalizeObjectKeys(row);
    const sku = stringify(normalized.termek || normalized.cikkszam || normalized.anyag);
    const minMax = minMaxBySku.get(sku) || {};
    return {
      "cikkszám": sku,
      "megnevezés": normalized.termek_rovid_leirasa || normalized.megnevezes || "",
      "lokáció": normalized.tarhely || normalized.terhely || normalized.raktarhely || "",
      "min": minMax.min ?? "",
      "max": minMax.max ?? "",
      "forgás": classifyIssueQuantity(minMax.issueQuantity, issueThresholds),
      "kiadás_mennyiség": minMax.issueQuantity ?? "",
    };
  });
}

function getIssueQuantityThresholds(values) {
  const sorted = values.filter((value) => Number.isFinite(value)).sort((left, right) => left - right);
  if (!sorted.length) return { high: Infinity, low: -Infinity };
  return {
    low: sorted[Math.floor(sorted.length * 0.25)] ?? sorted[0],
    high: sorted[Math.floor(sorted.length * 0.8)] ?? sorted[sorted.length - 1],
  };
}

function classifyIssueQuantity(value, thresholds) {
  const number = Number(value);
  if (!Number.isFinite(number)) return "";
  if (number <= 0) return "lassú";
  if (number >= thresholds.high) return "sűrű";
  if (number <= thresholds.low) return "lassú";
  return "közepes";
}

function normalizeObjectKeys(row) {
  const normalized = {};
  Object.entries(row).forEach(([key, value]) => {
    normalized[normalizeKey(key)] = typeof value === "string" ? value.trim() : value;
  });
  return normalized;
}

function importCompanionMasterSheet(workbook, fileName, sourceSheetName) {
  const masterSheetName = chooseWorkbookSheet(workbook, "master");
  if (!masterSheetName || masterSheetName === sourceSheetName) return;

  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[masterSheetName], { defval: "" });
  if (!rows.length) return;

  const headers = Object.keys(rows[0] || {}).map(normalizeKey);
  if (scoreSheetHeaders(headers, "master", masterSheetName) < 5) return;

  setMasterRows(rows, `${fileName} / ${masterSheetName}`);
}

function chooseWorkbookSheet(workbook, mode) {
  const candidates = workbook.SheetNames
    .map((sheetName) => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "", header: 1 });
      const header = rows.find((row) => row.some((cell) => stringify(cell))) || [];
      const headers = header.map(normalizeKey);
      return {
        sheetName,
        score: scoreSheetHeaders(headers, mode, sheetName),
        rowCount: rows.length,
      };
    })
    .sort((left, right) => right.score - left.score || right.rowCount - left.rowCount);

  return candidates[0]?.sheetName || workbook.SheetNames[0];
}

function scoreSheetHeaders(headers, mode, sheetName = "") {
  const normalizedSheetName = normalizeText(sheetName);
  const hasItem = headers.some((key) => ["cikkszam", "termek", "anyag", "cikkszam_"].includes(key));
  const hasLocation = headers.some((key) => ["lokacio", "location", "raktarhely", "tarhely", "terhely"].includes(key));
  const hasDescription = headers.some((key) => ["megnevezes", "nev", "anyagnev", "termek_rovid_leirasa"].includes(key));
  const hasMin = headers.some((key) => ["min", "minimum_keszlet", "minimum_mennyiseg"].includes(key));
  const hasMax = headers.some((key) => ["max", "maximum_keszlet", "maximum_mennyiseg", "max_mennyiseg"].includes(key));
  const hasContainer = headers.some((key) => ["tarolo", "lada", "tarolo_tipus"].includes(key));

  if (mode === "master") {
    const nameBonus = normalizedSheetName.includes("min") || normalizedSheetName.includes("max") ? 5 : 0;
    return [hasItem, hasMin, hasMax, hasLocation, hasContainer, hasDescription].filter(Boolean).length + nameBonus;
  }

  const nameBonus = normalizedSheetName.includes("cikkszam") || normalizedSheetName.includes("terhely") || normalizedSheetName.includes("tarhely") ? 5 : 0;
  return [hasItem, hasLocation, hasDescription, hasDescription, hasMin, hasMax].filter(Boolean).length + nameBonus;
}

function setRows(rawRows, sourceName) {
  const normalized = rawRows.map(normalizeRow).filter((row) => row.cikkszam);
  const missingColumns = getMissingColumns(rawRows[0] || {}, SOURCE_REQUIRED_COLUMNS);

  importRows = normalized;
  rows = importRows.map(mergeMasterData);
  analyzedRows = rows.map((row) => ({
    ...row,
    analysis: analyzeRow(row),
  }));

  dataStatus.textContent = missingColumns.length
    ? `${sourceName}: ${rows.length} sor, hiányzó oszlop: ${missingColumns.join(", ")}`
    : `${sourceName}: ${rows.length} sor betöltve`;

  renderAll();
}

function setMasterRows(rawRows, sourceName) {
  const normalized = rawRows.map(normalizeRow).filter((row) => row.cikkszam);
  masterRowsByItem = new Map();

  normalized.forEach((row) => {
    const existing = masterRowsByItem.get(row.cikkszam) || {};
    masterRowsByItem.set(row.cikkszam, {
      ...existing,
      ...removeEmptyFields(row),
    });
  });

  masterStatus.textContent = `${sourceName}: ${masterRowsByItem.size} cikkszám a törzsben`;
  saveMasterRows(normalized, masterStatus.textContent);

  if (importRows.length) {
    rows = importRows.map(mergeMasterData);
    analyzedRows = rows.map((row) => ({
      ...row,
      analysis: analyzeRow(row),
    }));
    renderAll();
  } else {
    renderMissingMasterData();
  }
}

function saveMasterRows(normalizedRows, statusText) {
  try {
    localStorage.setItem(STORAGE_KEYS.masterRows, JSON.stringify(normalizedRows));
    localStorage.setItem(STORAGE_KEYS.masterStatus, statusText);
  } catch (error) {
    masterStatus.textContent = `${statusText} | Mentés nem sikerült ezen az eszközön.`;
  }
}

function loadSavedMasterRows() {
  try {
    const saved = localStorage.getItem(STORAGE_KEYS.masterRows);
    if (!saved) return;

    const parsed = JSON.parse(saved);
    if (!Array.isArray(parsed) || !parsed.length) return;

    masterRowsByItem = new Map();
    parsed.forEach((row) => {
      const existing = masterRowsByItem.get(row.cikkszam) || {};
      masterRowsByItem.set(row.cikkszam, {
        ...existing,
        ...removeEmptyFields(row),
      });
    });

    masterStatus.textContent = localStorage.getItem(STORAGE_KEYS.masterStatus) || `Mentett törzsadat: ${masterRowsByItem.size} cikkszám`;
  } catch (error) {
    masterStatus.textContent = "Mentett törzsadat nem olvasható, töltsd be újra az Excelt.";
  }
}

function mergeMasterData(row) {
  const master = masterRowsByItem.get(row.cikkszam);
  if (!master) return row;

  return {
    ...row,
    megnevezes: row.megnevezes || master.megnevezes || "",
    min: row.min === "" ? master.min ?? "" : row.min,
    max: row.max === "" ? master.max ?? "" : row.max,
    forgas: row.forgas || master.forgas || "",
    kiadas_mennyiseg: row.kiadas_mennyiseg === "" ? master.kiadas_mennyiseg ?? "" : row.kiadas_mennyiseg,
    tarolo: row.tarolo || master.tarolo || "",
    suly_kategoria: row.suly_kategoria || master.suly_kategoria || "",
    kapacitas: row.kapacitas === "" ? master.kapacitas ?? "" : row.kapacitas,
    megjegyzes: row.megjegyzes || master.megjegyzes || "",
  };
}

function removeEmptyFields(row) {
  return Object.fromEntries(Object.entries(row).filter(([, value]) => value !== ""));
}

function normalizeRow(row) {
  const normalized = {};

  Object.entries(row).forEach(([key, value]) => {
    normalized[normalizeKey(key)] = typeof value === "string" ? value.trim() : value;
  });

  const lokacio = stringify(normalized.lokacio || normalized.location || normalized.raktarhely || normalized.tarhely || normalized.terhely);
  const explicitLevel = toNumber(normalized.szint);

  return {
    cikkszam: stringify(normalized.cikkszam || normalized.termek || normalized.anyag || normalized.cikkszam_),
    megnevezes: stringify(normalized.megnevezes || normalized.nev || normalized.anyagnev || normalized.termek_rovid_leirasa),
    lokacio,
    min: toNumber(normalized.min || normalized.minimum_keszlet || normalized.minimum_mennyiseg),
    max: toNumber(normalized.max || normalized.maximum_keszlet || normalized.maximum_mennyiseg || normalized.max_mennyiseg),
    forgas: stringify(normalized.forgas),
    kiadas_mennyiseg: toNumber(normalized.kiadas_mennyiseg || normalized.kiadasi_gyakorisag || normalized.kiadas),
    tarolo: stringify(normalized.tarolo || normalized.lada || normalized.tarolo_tipus),
    suly_kategoria: stringify(normalized.suly_kategoria || normalized.suly || normalized.nehezseg),
    szint: explicitLevel === "" ? getLocationLevel(lokacio) : explicitLevel,
    kapacitas: toNumber(normalized.kapacitas || normalized.also_ket_szint_kapacitas || normalized.befogadokepesseg),
    mennyiseg: toNumber(normalized.mennyiseg || normalized.keszlet || normalized.aktualis_keszlet),
    megjegyzes: stringify(normalized.megjegyzes),
  };
}

function normalizeKey(key) {
  return String(key)
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "_");
}

function normalizeText(value) {
  return stringify(value)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function stringify(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function toNumber(value) {
  if (value === null || value === undefined || value === "") return "";
  const number = Number(String(value).replace(",", ".").replace(/\s/g, ""));
  return Number.isFinite(number) ? number : "";
}

function getMissingColumns(firstRow, requiredColumns = SOURCE_REQUIRED_COLUMNS) {
  const available = Object.keys(firstRow).map(normalizeKey);
  return requiredColumns.filter((column) => !available.includes(column));
}

function analyzeRow(row) {
  const warnings = [];
  const errors = [];
  const locationZone = getLocationZone(row.lokacio);
  const level = Number(row.szint);
  const isInfoOnlyLevel = level >= 3;
  const forgas = normalizeText(row.forgas);
  const tarolo = normalizeText(row.tarolo);
  const suly = normalizeText(row.suly_kategoria);

  if (!row.lokacio) errors.push("Hiányzik a lokáció.");
  if (row.min === "") warnings.push("Hiányzik a minimum készlet.");
  if (row.max === "") warnings.push("Hiányzik a maximum készlet.");
  if (row.szint === "") warnings.push("Hiányzik a szint.");

  if (row.min !== "" && row.max !== "" && row.min > row.max) {
    errors.push("A minimum nagyobb, mint a maximum.");
  }

  if (!isInfoOnlyLevel && forgas.includes("suru") && locationZone === "hatul") {
    warnings.push("Sűrűn forgó anyag hátul van, érdemes előrébb tenni.");
  }

  if (!isInfoOnlyLevel && forgas.includes("lassu") && locationZone === "elol") {
    warnings.push("Lassan forgó anyag elöl van, jó helyet foglalhat a sűrű anyagok elől.");
  }

  if (!isInfoOnlyLevel && suly.includes("nehez") && level >= 2) {
    errors.push("Nehéz anyag magas szinten van.");
  }

  if (!isInfoOnlyLevel && (tarolo.includes("xl") || tarolo.includes("nagy")) && level >= 2) {
    warnings.push("Nagy vagy XL láda magas szinten van.");
  }

  if (!isInfoOnlyLevel && tarolo.includes("raklap") && level >= 2) {
    errors.push("Raklapos anyag nem ideális felső szinten.");
  }

  const severity = errors.length ? "error" : warnings.length ? "warning" : "ok";
  return { warnings, errors, severity, locationZone };
}

function getLocationZone(location) {
  const column = parseLocationParts(location).column;
  if (column >= 1 && column <= 10) return "elol";
  if (column >= 11 && column <= 20) return "kozep";
  if (column >= 21 && column <= 30) return "hatul";
  return "ismeretlen";
}

function getLocationLevel(location) {
  const parts = stringify(location).split("-").map((part) => part.trim()).filter(Boolean);
  const lastPart = parts[parts.length - 1];
  const level = toNumber(lastPart);
  return level === "" ? "" : level;
}

function getLocationPrefix(location) {
  const parts = stringify(location).split("-").map((part) => part.trim().toUpperCase()).filter(Boolean);
  if (parts.length < 2) return "";
  return `${parts[0]}-${parts[1]}`;
}

function isCalculationLocation(location) {
  return ["A-A", "A-B", "A-C", "A-D"].includes(getLocationPrefix(location));
}

function scorePlacement(row, itemRows) {
  if (!isCalculationLocation(row.lokacio) || Number(row.szint) > 2) {
    return { score: -999, reasons: ["nem kalkulációs lokáció, csak tájékoztató"] };
  }

  const forgas = normalizeText(row.forgas);
  const tarolo = normalizeText(row.tarolo);
  const suly = normalizeText(row.suly_kategoria);
  const zone = row.analysis.locationZone;
  const level = Number(row.szint);
  const item = analyzeItem(itemRows);
  let score = 0;
  const reasons = [];

  if (forgas.includes("suru")) {
    if (zone === "elol") {
      score += 45;
      reasons.push("sűrű anyag elöl van");
    } else if (zone === "kozep") {
      score += 15;
      reasons.push("sűrű anyag közép zónában van");
    } else if (zone === "hatul") {
      score -= 45;
      reasons.push("sűrű anyag hátul van");
    }
  }

  if (forgas.includes("lassu")) {
    if (zone === "hatul") {
      score += 30;
      reasons.push("lassú anyag hátrébb lehet");
    } else if (zone === "elol") {
      score -= 25;
      reasons.push("lassú anyag jó helyet foglal elöl");
    }
  }

  if (suly.includes("nehez")) {
    if (level === 1) {
      score += 55;
      reasons.push("nehéz anyag alsó szinten van");
    } else if (level >= 2) {
      score -= level >= 3 ? 90 : 60;
      reasons.push("nehéz anyag magas szinten van");
    }
  }

  if (tarolo.includes("xl") || tarolo.includes("nagy") || tarolo.includes("raklap")) {
    if (level === 1) {
      score += 35;
      reasons.push("nagy tároló alsó szinten van");
    } else if (level >= 3) {
      score -= 60;
      reasons.push("nagy tároló túl magasan van");
    } else if (level === 2) {
      score -= 15;
      reasons.push("nagy tároló második szinten figyelendő");
    }
  }

  if (item.max !== "" && row.kapacitas !== "") {
    const ratio = Number(row.kapacitas) / Number(item.max);
    if (ratio >= 0.35 && ratio <= 0.8) {
      score += 15;
      reasons.push("kapacitása hasznos arányban van a maxhoz");
    } else if (ratio > 1.5) {
      score -= 15;
      reasons.push("túl nagy helyet foglalhat a max készlethez képest");
    }
  }

  return { score, reasons };
}

function handleSearch(event) {
  event.preventDefault();
  runSearch(searchInput.value);
}

function setSearchMode(mode) {
  searchMode = mode;
  const isLocation = mode === "location";
  itemSearchMode.classList.toggle("active", !isLocation);
  locationSearchMode.classList.toggle("active", isLocation);
  searchTitle.textContent = isLocation ? "Tárhely keresés" : "Cikkszám keresés";
  searchInput.placeholder = isLocation ? "pl. A-A-1-1" : "pl. 1013818 vagy BH0-0820-BK-500";
  topSearchInput.placeholder = isLocation ? "Keresés tárhelyre..." : "Keresés cikkszámra...";
}

function runSearch(value) {
  const query = value.trim();
  if (!query) return;

  if (searchMode === "location") {
    const normalizedQuery = normalizeLocationQuery(query);
    const foundRows = analyzedRows.filter((row) => normalizeLocationQuery(row.lokacio).includes(normalizedQuery));
    renderLocationResult(foundRows, query);
    return;
  }

  const normalizedQuery = query.toLowerCase();
  const foundRows = analyzedRows.filter((row) => row.cikkszam.toLowerCase() === normalizedQuery);
  renderResult(foundRows, query);
}

function normalizeLocationQuery(value) {
  return stringify(value).toUpperCase().replace(/\s+/g, "");
}

async function startScanner() {
  if (!("mediaDevices" in navigator) || !navigator.mediaDevices.getUserMedia) {
    showScannerMessage("A böngésző nem engedi a kamera használatát.");
    return;
  }

  if (!("BarcodeDetector" in window)) {
    showScannerMessage("Ez a böngésző nem támogatja a kamerás vonalkódolvasást. Próbáld Chrome Androidon.");
    return;
  }

  scannerModal.hidden = false;
  scanConfirm.hidden = true;
  pendingScannedCode = "";
  scannerStatus.textContent = "Kamera indítása...";

  try {
    scannerDetector = new BarcodeDetector({
      formats: ["code_128", "code_39", "ean_13", "ean_8", "qr_code", "itf"],
    });
    scannerStream = await navigator.mediaDevices.getUserMedia({
      video: {
        facingMode: { ideal: "environment" },
        width: { ideal: 1280 },
        height: { ideal: 720 },
      },
      audio: false,
    });
    scannerVideo.srcObject = scannerStream;
    await scannerVideo.play();
    scannerActive = true;
    scannerStatus.textContent = "Tartsd a vonalkódot a keretbe.";
    scanLoop();
  } catch (error) {
    stopScanner();
    showScannerMessage(`Nem sikerült elindítani a kamerát: ${error.message}`);
  }
}

async function scanLoop() {
  if (!scannerActive || !scannerDetector) return;

  try {
    const codes = await scannerDetector.detect(scannerVideo);
    if (codes.length) {
      const rawValue = stringify(codes[0].rawValue);
      if (rawValue) {
        useScannedCode(rawValue);
        return;
      }
    }
  } catch (error) {
    scannerStatus.textContent = "Olvasás folyamatban...";
  }

  requestAnimationFrame(scanLoop);
}

function useScannedCode(rawValue) {
  const code = rawValue.trim();
  pendingScannedCode = code;
  scannerActive = false;
  scannedCodeText.textContent = code;
  scanConfirm.hidden = false;
  scannerStatus.textContent = "Ellenőrizd, hogy ez a cikkszám-e.";
}

function confirmScannedCode() {
  if (!pendingScannedCode) return;
  const code = pendingScannedCode;
  searchInput.value = code;
  topSearchInput.value = code;
  stopScanner();
  runSearch(code);
  dataStatus.textContent = `Beolvasva: ${code}`;
}

function retryScanner() {
  pendingScannedCode = "";
  scannedCodeText.textContent = "-";
  scanConfirm.hidden = true;
  scannerStatus.textContent = "Tartsd a jó vonalkódot a keretbe.";
  scannerActive = true;
  scanLoop();
}

function stopScanner() {
  scannerActive = false;
  pendingScannedCode = "";
  scanConfirm.hidden = true;
  if (scannerStream) {
    scannerStream.getTracks().forEach((track) => track.stop());
  }
  scannerStream = null;
  scannerVideo.srcObject = null;
  scannerModal.hidden = true;
}

function showScannerMessage(message) {
  scannerModal.hidden = false;
  scannerStatus.textContent = message;
}

function renderResult(matchingRows, query) {
  if (!matchingRows.length) {
    resultBox.className = "result-box";
    resultBox.innerHTML = `<strong>Nincs találat:</strong> ${escapeHtml(query)}`;
    return;
  }

  const item = analyzeItem(matchingRows);
  const firstRow = matchingRows[0];
  const messages = [...item.errors, ...item.warnings];
  const searchedRecommendation = buildSingleItemRecommendation(matchingRows);
  resultBox.className = "result-box";
  resultBox.innerHTML = `
    <div class="result-title">
      <div>
        <strong>${escapeHtml(firstRow.cikkszam)}</strong>
        <div class="issue-meta">${escapeHtml(firstRow.megnevezes || "Nincs megnevezés")} | ${matchingRows.length} lokáció</div>
      </div>
      ${renderBadge(item.severity)}
    </div>
    <div class="detail-grid">
      <div><small>Min / Max</small><strong>${formatValue(item.min)} / ${formatValue(item.max)}</strong></div>
      <div><small>Alsó 2 szint kapacitás</small><strong>${formatValue(item.lowerCapacity)}</strong></div>
      <div><small>Info kapacitás</small><strong>${formatValue(item.infoCapacity)}</strong></div>
      <div><small>Forgás</small><strong>${escapeHtml(firstRow.forgas || "-")}</strong></div>
      <div><small>6 havi kiadás</small><strong>${formatValue(firstRow.kiadas_mennyiseg)}</strong></div>
    </div>
    ${
      messages.length
        ? `<ul class="warning-list">${messages.map((message) => `<li>${escapeHtml(message)}</li>`).join("")}</ul>`
        : "<p>Állapot: rendben, a megadott kapacitás alapján az alsó két szint elég.</p>"
    }
    ${renderSearchedRecommendation(searchedRecommendation)}
    <div class="location-list">
      ${matchingRows
        .map(
          (row) => `
            <article class="location-card ${row.analysis.severity}">
              <div>
                <strong>${escapeHtml(row.lokacio || "-")}</strong>
                <small>${escapeHtml(row.analysis.locationZone)} | szint ${formatValue(row.szint)}</small>
              </div>
              <div><small>Kapacitás</small><strong>${formatValue(row.kapacitas)}</strong></div>
              <div><small>Tároló</small><strong>${escapeHtml(row.tarolo || "-")}</strong></div>
              <div><small>Súly</small><strong>${escapeHtml(row.suly_kategoria || "-")}</strong></div>
              ${renderBadge(row.analysis.severity)}
            </article>
          `
        )
        .join("")}
    </div>
  `;
}

function buildSingleItemRecommendation(itemRows) {
  const calculationRows = itemRows
    .filter((row) => isCalculationLocation(row.lokacio) && Number(row.szint) <= 2)
    .map((row) => ({
      ...row,
      placement: scorePlacement(row, itemRows),
    }))
    .sort((a, b) => b.placement.score - a.placement.score);

  return {
    bestRows: calculationRows.slice(0, 2),
    reasons: calculationRows.flatMap((row) => row.placement.reasons).slice(0, 3),
  };
}

function renderSearchedRecommendation(recommendation) {
  if (!recommendation.bestRows.length) {
    return `<div class="inline-advice">Ehhez a cikkszámhoz nincs 1-2. szintű A-A, A-B, A-C vagy A-D javasolható lokáció.</div>`;
  }

  const meaningfulRows = recommendation.bestRows.filter((row) => row.placement.score !== 0 || row.placement.reasons.length);
  if (!meaningfulRows.length) {
    return `<div class="inline-advice"><strong>Rövid ajánlás</strong><div>Nincs érdemi ajánlás ehhez a cikkszámhoz.</div><small>Nincs 6 havi kiadás vagy más pontozható adat, ezért a lokáció csak információként jelenik meg.</small></div>`;
  }

  const bestScore = meaningfulRows[0].placement.score;
  if (bestScore < 0) {
    return `
      <div class="inline-advice warning-advice">
        <strong>Rövid ajánlás</strong>
        <div>Nem ideális elhelyezés: ${meaningfulRows.map((row) => `${escapeHtml(row.lokacio)} (${row.placement.score} pont)`).join(" → ")}</div>
        <small>${escapeHtml(recommendation.reasons[0] || "Lassan mozgó vagy rossz zónában lévő anyag, érdemes hátrébb tenni.")}</small>
      </div>
    `;
  }

  return `
    <div class="inline-advice">
      <strong>Rövid ajánlás</strong>
      <div>${meaningfulRows.map((row) => `${escapeHtml(row.lokacio)} (${row.placement.score} pont)`).join(" → ")}</div>
      <small>${escapeHtml(recommendation.reasons[0] || "Pontozható adatok alapján ez a legjobb sorrend.")}</small>
    </div>
  `;
}

function renderLocationResult(matchingRows, query) {
  if (!matchingRows.length) {
    resultBox.className = "result-box";
    resultBox.innerHTML = `<strong>Nincs találat erre a tárhelyre:</strong> ${escapeHtml(query)}`;
    return;
  }

  const exactRows = matchingRows.filter((row) => normalizeLocationQuery(row.lokacio) === normalizeLocationQuery(query));
  const rowsToShow = exactRows.length ? exactRows : matchingRows;
  const locationName = exactRows.length ? exactRows[0].lokacio : query;
  const errorCountForLocation = rowsToShow.filter((row) => row.analysis.severity === "error").length;
  const warningCountForLocation = rowsToShow.filter((row) => row.analysis.severity === "warning").length;
  const totalQuantity = rowsToShow.reduce((sum, row) => sum + (Number(row.mennyiseg) || 0), 0);

  resultBox.className = "result-box";
  resultBox.innerHTML = `
    <div class="result-title">
      <div>
        <strong>${escapeHtml(locationName)}</strong>
        <div class="issue-meta">${rowsToShow.length} sor ezen a tárhelyen</div>
      </div>
      ${renderBadge(errorCountForLocation ? "error" : warningCountForLocation ? "warning" : "ok")}
    </div>
    <div class="detail-grid">
      <div><small>Anyagsor</small><strong>${rowsToShow.length}</strong></div>
      <div><small>Össz. mennyiség</small><strong>${formatValue(totalQuantity || "")}</strong></div>
      <div><small>Figyelendő</small><strong>${warningCountForLocation}</strong></div>
      <div><small>Kritikus</small><strong>${errorCountForLocation}</strong></div>
    </div>
    <div class="location-list">
      ${rowsToShow
        .map((row) => {
          const messages = [...row.analysis.errors, ...row.analysis.warnings];
          return `
            <article class="location-card ${row.analysis.severity}">
              <div>
                <strong>${escapeHtml(row.cikkszam || "-")}</strong>
                <small>${escapeHtml(row.megnevezes || "Nincs megnevezés")}</small>
              </div>
              <div><small>Mennyiség</small><strong>${formatValue(row.mennyiseg)}</strong></div>
              <div><small>Min / Max</small><strong>${formatValue(row.min)} / ${formatValue(row.max)}</strong></div>
              <div><small>Tároló</small><strong>${escapeHtml(row.tarolo || "-")}</strong></div>
              ${renderBadge(row.analysis.severity)}
              ${messages.length ? `<div class="location-message">${messages.map(escapeHtml).join("<br>")}</div>` : ""}
            </article>
          `;
        })
        .join("")}
    </div>
  `;
}

function analyzeItem(itemRows) {
  const minValues = itemRows.map((row) => row.min).filter((value) => value !== "");
  const maxValues = itemRows.map((row) => row.max).filter((value) => value !== "");
  const min = minValues.length ? Math.max(...minValues) : "";
  const max = maxValues.length ? Math.max(...maxValues) : "";
  const lowerCapacity = itemRows
    .filter((row) => isCalculationLocation(row.lokacio) && Number(row.szint) <= 2)
    .reduce((sum, row) => sum + (Number(row.kapacitas) || 0), 0);
  const infoCapacity = itemRows
    .filter((row) => !isCalculationLocation(row.lokacio) || Number(row.szint) >= 3)
    .reduce((sum, row) => sum + (Number(row.kapacitas) || 0), 0);
  const warnings = [];
  const errors = [];

  itemRows.filter((row) => Number(row.szint) <= 2).forEach((row) => {
    errors.push(...row.analysis.errors.map((message) => `${row.lokacio || "-"}: ${message}`));
    warnings.push(...row.analysis.warnings.map((message) => `${row.lokacio || "-"}: ${message}`));
  });

  if (!itemRows.some((row) => row.kapacitas !== "")) {
    warnings.push("Nincs megadva kapacitás, ezért nem lehet alsó két szintre számolni.");
  } else {
    if (min !== "" && lowerCapacity < min) {
      errors.push(`Az alsó két szint kapacitása (${lowerCapacity}) még a minimum készlethez (${min}) sem elég.`);
    } else if (max !== "" && lowerCapacity < max) {
      warnings.push(`Az alsó két szint kapacitása (${lowerCapacity}) nem elég a maximum készlethez (${max}).`);
    }

    if (infoCapacity > 0) {
      warnings.push(`A nem kalkulációs vagy 3-4. szinten van még ${infoCapacity} kapacitás, de ez csak tájékoztató.`);
    }
  }

  const severity = errors.length ? "error" : warnings.length ? "warning" : "ok";
  return { min, max, lowerCapacity, infoCapacity, warnings, errors, severity };
}

function groupRowsByItem() {
  const groups = new Map();
  analyzedRows.forEach((row) => {
    if (!groups.has(row.cikkszam)) groups.set(row.cikkszam, []);
    groups.get(row.cikkszam).push(row);
  });
  return Array.from(groups.values());
}

function buildPlacementRecommendations() {
  return groupRowsByItem()
    .map((itemRows) => {
      const item = analyzeItem(itemRows);
      const calculationRows = itemRows.filter((row) => isCalculationLocation(row.lokacio) && Number(row.szint) <= 2);
      const scoredRows = calculationRows
        .map((row) => ({
          ...row,
          placement: scorePlacement(row, itemRows),
        }))
        .sort((a, b) => b.placement.score - a.placement.score);

      return {
        cikkszam: itemRows[0].cikkszam,
        megnevezes: itemRows[0].megnevezes,
        item,
        bestRows: scoredRows.slice(0, 3),
        hasCalculationRows: calculationRows.length > 0,
        score: scoredRows.reduce((sum, row) => sum + row.placement.score, 0),
      };
    })
    .sort((a, b) => {
      if (a.item.severity === "error" && b.item.severity !== "error") return -1;
      if (a.item.severity !== "error" && b.item.severity === "error") return 1;
      if (a.item.severity === "warning" && b.item.severity === "ok") return -1;
      if (a.item.severity === "ok" && b.item.severity === "warning") return 1;
      return b.score - a.score;
    });
}

function renderLogicalPlacement() {
  if (!analyzedRows.length) {
    logicList.textContent = "Nincs még adat.";
    return;
  }

  const recommendations = buildPlacementRecommendations();

  logicList.innerHTML = recommendations
    .map((recommendation) => {
      const messages = [...recommendation.item.errors, ...recommendation.item.warnings];
      return `
        <article class="logic-card ${recommendation.item.severity}">
          <div class="logic-head">
            <div>
              <strong>${escapeHtml(recommendation.cikkszam)}</strong>
              <small>${escapeHtml(recommendation.megnevezes || "-")}</small>
            </div>
            ${renderBadge(recommendation.item.severity)}
          </div>
          <div class="logic-summary">
            <div><small>Alsó 2 szint</small><strong>${formatValue(recommendation.item.lowerCapacity)}</strong></div>
            <div><small>Info kapacitás</small><strong>${formatValue(recommendation.item.infoCapacity)}</strong></div>
            <div><small>Min / Max</small><strong>${formatValue(recommendation.item.min)} / ${formatValue(recommendation.item.max)}</strong></div>
          </div>
          ${
            messages.length
              ? `<ul class="warning-list">${messages.slice(0, 4).map((message) => `<li>${escapeHtml(message)}</li>`).join("")}</ul>`
              : "<p>Jelenlegi adatok alapján rendben.</p>"
          }
          <div class="logic-recommendation">
            <h3>Javasolt sorrend</h3>
            ${
              recommendation.hasCalculationRows
                ? recommendation.bestRows
                    .map(
                      (row) => `
                        <div class="logic-row">
                          <strong>${escapeHtml(row.lokacio)}</strong>
                          <span>${escapeHtml(row.analysis.locationZone)} | szint ${formatValue(row.szint)} | kap. ${formatValue(row.kapacitas)}</span>
                          <em>${row.placement.score} pont</em>
                        </div>
                      `
                    )
                    .join("")
                : `<p class="logic-note">Ehhez a cikkszámhoz nincs A-A, A-B, A-C vagy A-D lokáció.</p>`
            }
          </div>
          <p class="logic-note">${escapeHtml(getBestReason(recommendation.bestRows))}</p>
        </article>
      `;
    })
    .join("");
}

function getBestReason(rowsForItem) {
  const reasons = rowsForItem.flatMap((row) => row.placement.reasons);
  if (!reasons.length) return "Nincs külön indok, az adatok alapján semleges elhelyezés.";
  return `Fő indok: ${reasons[0]}.`;
}

function parsePickingText(text) {
  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map(parsePickingLine)
    .filter(Boolean);
}

function parsePickingLine(line, index) {
  const normalized = normalizeText(line);
  const locationMatch = line.match(/\b[A-Z]-[A-Z]-\d{1,3}-\d\b/i);
  if (!locationMatch) return null;

  const isBooked = normalized.includes("konyvelve");
  const isOpen = normalized.includes("uj tetel") || !isBooked;
  if (!isOpen) return null;

  const location = locationMatch[0].toUpperCase();
  const beforeLocation = line.slice(0, locationMatch.index);
  const afterLocation = line.slice(locationMatch.index + location.length);
  const skuMatch = beforeLocation.match(/\b\d{6,8}\b/);
  const numbersBeforeLocation = beforeLocation.match(/\b\d{1,6}\b/g) || [];
  const quantity = numbersBeforeLocation.length ? numbersBeforeLocation[numbersBeforeLocation.length - 1] : "";
  const sku = skuMatch ? skuMatch[0] : "";
  const description = cleanPickingDescription(beforeLocation, sku, quantity);

  return {
    id: `${location}-${sku || index}-${index}`,
    raw: line,
    location,
    sku,
    description,
    quantity,
    statusText: afterLocation.trim(),
    done: false,
  };
}

function cleanPickingDescription(text, sku, quantity) {
  let value = text;
  if (sku) value = value.replace(sku, " ");
  if (quantity) value = value.replace(new RegExp(`\\b${quantity}\\b\\s*$`), " ");
  value = value
    .replace(/\b\d{4}[.:-]\d{1,2}[.:-]\d{1,2}\b/g, " ")
    .replace(/\b\d{1,2}:\d{2}\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return value || "Nincs megnevezés";
}

function comparePickingRows(left, right) {
  const leftParts = getPickingRouteKey(left.location);
  const rightParts = getPickingRouteKey(right.location);
  return (
    leftParts.route - rightParts.route ||
    leftParts.column - rightParts.column ||
    leftParts.level - rightParts.level ||
    left.sku.localeCompare(right.sku, "hu", { numeric: true, sensitivity: "base" })
  );
}

function getPickingRouteKey(location) {
  const parts = parseLocationParts(location);
  const prefix = `${parts.zone}-${parts.aisle}`;
  const routeMap = {
    "A-A": { route: 1, direction: 1 },
    "A-B": { route: 2, direction: 1 },
    "A-C": { route: 3, direction: -1 },
    "A-D": { route: 4, direction: -1 },
  };
  const route = routeMap[prefix];

  if (!route) {
    return {
      route: 99,
      column: parts.column,
      level: parts.level,
    };
  }

  return {
    route: route.route,
    column: parts.column * route.direction,
    level: parts.level,
  };
}

function parseLocationParts(location) {
  const parts = stringify(location).toUpperCase().split("-");
  return {
    zone: parts[0] || "",
    aisle: parts[1] || "",
    column: Number(parts[2] || 0),
    level: Number(parts[3] || 0),
  };
}

function renderPickingList() {
  if (!pickingRows.length) {
    pickingStatus.textContent = "Nem találtam nyitott kiadási sort. Ha van nyitott sor a képen, próbáld közelebbről vagy szemből fotózni.";
    pickingList.innerHTML = `<div class="result-box empty">Nincs felismerhető nyitott tétel.</div>`;
    return;
  }

  const doneCount = pickingRows.filter((row) => row.done).length;
  pickingStatus.textContent = `Felismert nyitott sor: ${pickingRows.length}. Kész: ${doneCount}/${pickingRows.length}. Ellenőrizd, mert fotóból olvasva lehet tévesztés.`;
  pickingList.innerHTML = `
    <div class="picking-progress">
      <strong>${doneCount}/${pickingRows.length} kész</strong>
      <span>Következő: ${escapeHtml(nextPickingLocation() || "-")}</span>
    </div>
    ${pickingRows.map(renderPickingCard).join("")}
  `;
}

function renderPickingCard(row, index) {
  return `
    <article class="picking-card ${row.done ? "done" : ""}">
      <label class="picking-check">
        <input type="checkbox" data-picking-id="${escapeHtml(row.id)}" ${row.done ? "checked" : ""}>
        <span>${index + 1}</span>
      </label>
      <div class="picking-edit-grid">
        <label>
          <small>Tárhely</small>
          <input data-picking-row-id="${escapeHtml(row.id)}" data-picking-field="location" value="${escapeHtml(row.location)}">
        </label>
        <label>
          <small>Cikkszám</small>
          <input data-picking-row-id="${escapeHtml(row.id)}" data-picking-field="sku" value="${escapeHtml(row.sku)}">
        </label>
      </div>
      <div class="picking-edit-grid">
        <label>
          <small>Mennyiség</small>
          <input data-picking-row-id="${escapeHtml(row.id)}" data-picking-field="quantity" inputmode="numeric" value="${escapeHtml(row.quantity)}">
        </label>
        <label>
          <small>Megnevezés</small>
          <input data-picking-row-id="${escapeHtml(row.id)}" data-picking-field="description" value="${escapeHtml(row.description)}">
        </label>
      </div>
      <details>
        <summary>Nyers sor</summary>
        <small>${escapeHtml(row.raw)}</small>
      </details>
      ${row.statusText ? `<small class="picking-note">${escapeHtml(row.statusText)}</small>` : ""}
      <button class="danger-button" data-remove-picking-id="${escapeHtml(row.id)}" type="button">Törlés</button>
    </article>
  `;
}

function nextPickingLocation() {
  const next = pickingRows.find((row) => !row.done);
  return next ? next.location : "";
}

function togglePickingDone(id, checked) {
  const row = pickingRows.find((item) => item.id === id);
  if (!row) return;
  row.done = checked;
  renderPickingList();
}

function updatePickingField(id, field, value, rerender = true) {
  const row = pickingRows.find((item) => item.id === id);
  if (!row || !["location", "sku", "quantity", "description"].includes(field)) return;
  row[field] = field === "location" ? value.toUpperCase().trim() : value.trim();
  if (rerender) {
    pickingRows.sort(comparePickingRows);
    renderPickingList();
  } else {
    const doneCount = pickingRows.filter((item) => item.done).length;
    pickingStatus.textContent = `Felismert nyitott sor: ${pickingRows.length}. Kész: ${doneCount}/${pickingRows.length}. Szerkesztés után a mező elhagyásakor rendez újra.`;
  }
}

function removePickingRow(id) {
  pickingRows = pickingRows.filter((row) => row.id !== id);
  renderPickingList();
}

function renderAll() {
  renderSummary();
  renderIssues();
  renderTable();
  renderLogicalPlacement();
  renderMissingMasterData();
}

function renderMissingMasterData() {
  if (!importRows.length) {
    missingMasterList.textContent = "Nincs még napi export betöltve.";
    return;
  }

  const missingItems = buildMissingMasterItems();

  if (!missingItems.length) {
    missingMasterList.textContent = "Minden betöltött cikkszámhoz megvannak a szükséges törzsadatok.";
    return;
  }

  missingMasterList.innerHTML = missingItems
    .map(
      (item) => `
        <article class="issue-card warning">
          <div>
            <div class="issue-code">${escapeHtml(item.cikkszam)}</div>
            <div class="issue-meta">${escapeHtml(item.megnevezes || "-")} | ${item.locations} lokáció</div>
          </div>
          <div>Hiányzik: ${item.missing.map(escapeHtml).join(", ")}</div>
          ${renderBadge("warning")}
        </article>
      `
    )
    .join("");
}

function buildMissingMasterItems() {
  const groups = new Map();

  rows.forEach((row) => {
    if (!groups.has(row.cikkszam)) groups.set(row.cikkszam, []);
    groups.get(row.cikkszam).push(row);
  });

  return Array.from(groups.entries())
    .map(([cikkszam, itemRows]) => {
      const firstRow = itemRows[0];
      const missing = MASTER_FIELDS.filter((field) => !itemRows.some((row) => row[field] !== ""));
      return {
        cikkszam,
        megnevezes: firstRow.megnevezes,
        locations: itemRows.length,
        missing,
      };
    })
    .filter((item) => item.missing.length);
}

function renderSummary() {
  const warningRows = analyzedRows.filter((row) => row.analysis.severity === "warning").length;
  const errorRows = analyzedRows.filter((row) => row.analysis.severity === "error").length;
  itemCount.textContent = analyzedRows.length;
  warningCount.textContent = warningRows;
  errorCount.textContent = errorRows;
}

function renderIssues() {
  const filter = issueFilter.value.trim().toLowerCase();
  const issueRows = analyzedRows
    .filter((row) => row.analysis.severity !== "ok")
    .filter((row) => {
      const haystack = `${row.cikkszam} ${row.megnevezes} ${row.lokacio} ${row.forgas} ${row.tarolo} ${row.suly_kategoria}`.toLowerCase();
      return !filter || haystack.includes(filter);
    });

  if (!issueRows.length) {
    issuesList.textContent = analyzedRows.length ? "Nincs probléma a szűrés alapján." : "Nincs még adat.";
    return;
  }

  issuesList.innerHTML = issueRows
    .map((row) => {
      const messages = [...row.analysis.errors, ...row.analysis.warnings];
      return `
        <article class="issue-card ${row.analysis.severity}">
          <div>
            <div class="issue-code">${escapeHtml(row.cikkszam)}</div>
            <div class="issue-meta">${escapeHtml(row.lokacio || "-")} | ${escapeHtml(row.megnevezes || "-")}</div>
          </div>
          <div>${messages.map(escapeHtml).join("<br>")}</div>
          ${renderBadge(row.analysis.severity)}
        </article>
      `;
    })
    .join("");
}

function renderTable() {
  if (!analyzedRows.length) {
    dataTableBody.innerHTML = `<tr><td colspan="11">Nincs betöltött adat.</td></tr>`;
    return;
  }

  dataTableBody.innerHTML = analyzedRows
    .map(
      (row) => `
        <tr>
          <td>${escapeHtml(row.cikkszam)}</td>
          <td>${escapeHtml(row.megnevezes || "-")}</td>
          <td>${escapeHtml(row.lokacio || "-")}</td>
          <td>${formatValue(row.min)}</td>
          <td>${formatValue(row.max)}</td>
          <td>${escapeHtml(row.forgas || "-")}</td>
          <td>${escapeHtml(row.tarolo || "-")}</td>
          <td>${escapeHtml(row.suly_kategoria || "-")}</td>
          <td>${formatValue(row.szint)}</td>
          <td>${formatValue(row.kapacitas)}</td>
          <td>${renderBadge(row.analysis.severity)}</td>
        </tr>
      `
    )
    .join("");
}

function renderBadge(severity) {
  const labels = {
    ok: "OK",
    warning: "Figyelendő",
    error: "Kritikus",
  };
  return `<span class="badge ${severity}">${labels[severity]}</span>`;
}

function formatValue(value) {
  return value === "" ? "-" : escapeHtml(value);
}

function parseCsv(text) {
  const lines = text.split(/\r?\n/).filter((line) => line.trim());
  if (!lines.length) return [];
  const delimiter = lines[0].includes(";") ? ";" : ",";
  const headers = splitCsvLine(lines[0], delimiter);

  return lines.slice(1).map((line) => {
    const values = splitCsvLine(line, delimiter);
    return headers.reduce((row, header, index) => {
      row[header] = values[index] || "";
      return row;
    }, {});
  });
}

function splitCsvLine(line, delimiter) {
  const values = [];
  let current = "";
  let inQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const nextChar = line[index + 1];

    if (char === '"' && nextChar === '"') {
      current += '"';
      index += 1;
    } else if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === delimiter && !inQuotes) {
      values.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }

  values.push(current.trim());
  return values;
}

function showImportError(message) {
  dataStatus.textContent = message;
  resultBox.className = "result-box";
  resultBox.innerHTML = `<strong>Import hiba:</strong> ${escapeHtml(message)}`;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
