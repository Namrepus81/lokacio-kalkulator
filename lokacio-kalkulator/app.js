const SOURCE_REQUIRED_COLUMNS = [
  "cikkszam",
  "lokacio",
];

const MASTER_FIELDS = [
  "min",
  "max",
  "forgas",
  "tarolo",
  "suly_kategoria",
  "kapacitas",
];

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
const masterFileInput = document.querySelector("#masterFileInput");
const searchForm = document.querySelector("#searchForm");
const searchInput = document.querySelector("#searchInput");
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

let scannerStream = null;
let scannerDetector = null;
let scannerActive = false;

fileInput.addEventListener("change", handleFileChange);
masterFileInput.addEventListener("change", handleMasterFileChange);
searchForm.addEventListener("submit", handleSearch);
issueFilter.addEventListener("input", renderIssues);
loadSampleButton.addEventListener("click", () => setRows(SAMPLE_ROWS, "Minta adatok betöltve"));
logicButton.addEventListener("click", renderLogicalPlacement);
scanButton.addEventListener("click", startScanner);
closeScannerButton.addEventListener("click", stopScanner);

function handleFileChange(event) {
  const file = event.target.files[0];
  if (!file) {
    dataStatus.textContent = "Nem lett fájl kiválasztva.";
    return;
  }
  dataStatus.textContent = `Napi export kiválasztva: ${file.name}. Olvasás indul...`;
  readUploadedRows(file, (parsedRows) => setRows(parsedRows, file.name));
}

function handleMasterFileChange(event) {
  const file = event.target.files[0];
  if (!file) {
    masterStatus.textContent = "Nem lett törzsadat fájl kiválasztva.";
    return;
  }
  masterStatus.textContent = `Törzsadat kiválasztva: ${file.name}. Olvasás indul...`;
  readUploadedRows(file, (parsedRows) => setMasterRows(parsedRows, file.name));
}

function readUploadedRows(file, onRows) {
  const extension = file.name.split(".").pop().toLowerCase();
  const reader = new FileReader();

  reader.onerror = () => {
    showImportError(`Nem sikerült olvasni a fájlt: ${file.name}`);
  };

  reader.onload = (loadEvent) => {
    try {
      if (extension === "csv") {
        const text = loadEvent.target.result;
        onRows(parseCsv(text));
        return;
      }

      if (!window.XLSX) {
        showImportError("Az Excel olvasó nem töltődött be mobilon. Frissítsd az oldalt, vagy próbáld CSV-ként mentve feltölteni.");
        return;
      }

      const workbook = XLSX.read(loadEvent.target.result, { type: "array" });
      const firstSheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[firstSheetName];
      const parsedRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      onRows(parsedRows);
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

function mergeMasterData(row) {
  const master = masterRowsByItem.get(row.cikkszam);
  if (!master) return row;

  return {
    ...row,
    megnevezes: row.megnevezes || master.megnevezes || "",
    min: row.min === "" ? master.min ?? "" : row.min,
    max: row.max === "" ? master.max ?? "" : row.max,
    forgas: row.forgas || master.forgas || "",
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

  const lokacio = stringify(normalized.lokacio || normalized.location || normalized.raktarhely);
  const explicitLevel = toNumber(normalized.szint);

  return {
    cikkszam: stringify(normalized.cikkszam || normalized.termek || normalized.anyag || normalized.cikkszam_),
    megnevezes: stringify(normalized.megnevezes || normalized.nev || normalized.anyagnev || normalized.termek_rovid_leirasa),
    lokacio,
    min: toNumber(normalized.min),
    max: toNumber(normalized.max),
    forgas: stringify(normalized.forgas),
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
  const forgas = normalizeText(row.forgas);
  const tarolo = normalizeText(row.tarolo);
  const suly = normalizeText(row.suly_kategoria);

  if (!row.lokacio) errors.push("Hiányzik a lokáció.");
  if (row.min === "") warnings.push("Hiányzik a minimum készlet.");
  if (row.max === "") warnings.push("Hiányzik a maximum készlet.");
  if (!row.forgas) warnings.push("Hiányzik a forgás kategória.");
  if (!row.tarolo) warnings.push("Hiányzik a tároló típusa.");
  if (!row.suly_kategoria) warnings.push("Hiányzik a súly/nehezség kategória.");
  if (row.szint === "") warnings.push("Hiányzik a szint.");
  if (row.kapacitas === "") warnings.push("Hiányzik a lokáció kapacitása.");

  if (row.min !== "" && row.max !== "" && row.min > row.max) {
    errors.push("A minimum nagyobb, mint a maximum.");
  }

  if (forgas.includes("suru") && locationZone === "hatul") {
    warnings.push("Sűrűn forgó anyag hátul van, érdemes előrébb tenni.");
  }

  if (forgas.includes("lassu") && locationZone === "elol") {
    warnings.push("Lassan forgó anyag elöl van, jó helyet foglalhat a sűrű anyagok elől.");
  }

  if (suly.includes("nehez") && Number(row.szint) >= 2) {
    errors.push("Nehéz anyag magas szinten van.");
  }

  if ((tarolo.includes("xl") || tarolo.includes("nagy")) && Number(row.szint) >= 2) {
    warnings.push("Nagy vagy XL láda magas szinten van.");
  }

  if (tarolo.includes("raklap") && Number(row.szint) >= 2) {
    errors.push("Raklapos anyag nem ideális felső szinten.");
  }

  const severity = errors.length ? "error" : warnings.length ? "warning" : "ok";
  return { warnings, errors, severity, locationZone };
}

function getLocationZone(location) {
  const prefix = stringify(location).split("-")[0].charAt(0).toUpperCase();
  if (["A", "B"].includes(prefix)) return "elol";
  if (["C", "D"].includes(prefix)) return "kozep";
  if (["E", "F", "G", "H"].includes(prefix)) return "hatul";
  return "ismeretlen";
}

function getLocationLevel(location) {
  const parts = stringify(location).split("-").map((part) => part.trim()).filter(Boolean);
  const lastPart = parts[parts.length - 1];
  const level = toNumber(lastPart);
  return level === "" ? "" : level;
}

function scorePlacement(row, itemRows) {
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

  if (level === 1) score += 18;
  if (level === 2) score += 10;
  if (level === 3) {
    score -= 35;
    reasons.push("3. szint csak tájékoztató tartalék");
  }
  if (level >= 4) {
    score -= 55;
    reasons.push("4. szint csak tájékoztató tartalék");
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
  const query = searchInput.value.trim().toLowerCase();
  if (!query) return;

  const foundRows = analyzedRows.filter((row) => row.cikkszam.toLowerCase() === query);
  renderResult(foundRows, query);
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
  searchInput.value = code;
  stopScanner();
  renderResult(
    analyzedRows.filter((row) => row.cikkszam.toLowerCase() === code.toLowerCase()),
    code
  );
  dataStatus.textContent = `Beolvasva: ${code}`;
}

function stopScanner() {
  scannerActive = false;
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
      <div><small>3-4. szint tájékoztató</small><strong>${formatValue(item.infoCapacity)}</strong></div>
      <div><small>Forgás</small><strong>${escapeHtml(firstRow.forgas || "-")}</strong></div>
    </div>
    ${
      messages.length
        ? `<ul class="warning-list">${messages.map((message) => `<li>${escapeHtml(message)}</li>`).join("")}</ul>`
        : "<p>Állapot: rendben, a megadott kapacitás alapján az alsó két szint elég.</p>"
    }
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

function analyzeItem(itemRows) {
  const minValues = itemRows.map((row) => row.min).filter((value) => value !== "");
  const maxValues = itemRows.map((row) => row.max).filter((value) => value !== "");
  const min = minValues.length ? Math.max(...minValues) : "";
  const max = maxValues.length ? Math.max(...maxValues) : "";
  const lowerCapacity = itemRows
    .filter((row) => Number(row.szint) <= 2)
    .reduce((sum, row) => sum + (Number(row.kapacitas) || 0), 0);
  const infoCapacity = itemRows
    .filter((row) => Number(row.szint) >= 3)
    .reduce((sum, row) => sum + (Number(row.kapacitas) || 0), 0);
  const warnings = [];
  const errors = [];

  itemRows.forEach((row) => {
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
      warnings.push(`A 3-4. szinten van még ${infoCapacity} kapacitás, de ez csak tájékoztató, nem számít bele a max kalkulációba.`);
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
      const scoredRows = itemRows
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
            <div><small>3-4. szint info</small><strong>${formatValue(recommendation.item.infoCapacity)}</strong></div>
            <div><small>Min / Max</small><strong>${formatValue(recommendation.item.min)} / ${formatValue(recommendation.item.max)}</strong></div>
          </div>
          ${
            messages.length
              ? `<ul class="warning-list">${messages.slice(0, 4).map((message) => `<li>${escapeHtml(message)}</li>`).join("")}</ul>`
              : "<p>Jelenlegi adatok alapján rendben.</p>"
          }
          <div class="logic-recommendation">
            <h3>Javasolt sorrend</h3>
            ${recommendation.bestRows
              .map(
                (row) => `
                  <div class="logic-row">
                    <strong>${escapeHtml(row.lokacio)}</strong>
                    <span>${escapeHtml(row.analysis.locationZone)} | szint ${formatValue(row.szint)} | kap. ${formatValue(row.kapacitas)}</span>
                    <em>${row.placement.score} pont</em>
                  </div>
                `
              )
              .join("")}
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
