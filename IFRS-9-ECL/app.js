const sampleRows = [
  {
    "Customer ID": "CUST-001",
    "Customer Name": "Aster Manufacturing Ltd",
    "Disbursement Date": "2024-02-15",
    "Due Date": "2027-02-15",
    "Loan Amount": 125000000,
    "Outstanding Balance": 118000000,
    "Loan Classification": "Performing",
    "Past Due Days": 0,
    "EIR %": 13,
    "Collateral Type": "Real Estate",
    "Collateral Value": 160000000,
    "Years In Business": 8,
    "Credit Search": "Positive",
    "Restructuring": "No"
  },
  {
    "Customer ID": "CUST-002",
    "Customer Name": "Nile Trading Plc",
    "Disbursement Date": "2025-01-20",
    "Due Date": "2026-01-20",
    "Loan Amount": 38000000,
    "Outstanding Balance": 36000000,
    "Loan Classification": "Watchlist",
    "Past Due Days": 48,
    "EIR %": 11,
    "Collateral Type": "Receivable",
    "Collateral Value": 20000000,
    "Years In Business": 3,
    "Credit Search": "Positive",
    "Restructuring": "No"
  },
  {
    "Customer ID": "CUST-003",
    "Customer Name": "Kora Energy Services",
    "Disbursement Date": "2024-06-01",
    "Due Date": "2025-06-01",
    "Loan Amount": 73000000,
    "Outstanding Balance": 69000000,
    "Credit Limit": 90000000,
    "Loan Classification": "Overdraft",
    "Past Due Days": 97,
    "EIR %": 15.5,
    "Collateral Type": "Equipment",
    "Collateral Value": 35000000,
    "Years In Business": 1,
    "Credit Search": "Negative",
    "Restructuring": "Yes"
  }
];

const fieldAliases = {
  customerId: ["customer id", "customer_id", "client id", "obligor id", "asset id", "loan id", "account number", "id"],
  customerName: ["customer name", "borrower", "customer", "client", "counterparty", "obligor", "name"],
  loanClassification: ["loan classification", "classification", "risk classification", "asset class", "product", "loan type", "facility type"],
  amount: ["amount", "loan amount", "principal", "facility amount", "approved amount", "limit", "total line"],
  outstanding: ["outstanding", "outstanding balance", "gross carrying amount", "amortized cost", "amortised cost", "carrying amount", "ead", "balance"],
  dueDate: ["due date", "maturity date", "expiry date", "end date", "repayment date"],
  daysPastDue: ["past due day", "past due days", "days past due", "dpd", "arrears days", "no. of days past due"],
  eir: ["eir", "effective interest rate", "interest rate", "discount rate"],
  collateralValue: ["collateral value", "security value", "forced sale value", "fsv", "collateral amount"],
  collateralType: ["collateral type", "security type", "security", "collateral class"],
  creditLimit: ["credit limit", "total line", "approved limit", "facility limit"],
  undrawnLimit: ["undrawn limit", "undrawn", "unused limit", "limit undrawn"],
  ccf: ["ccf", "credit conversion factor"],
  stage: ["stage", "ifrs 9 stage", "ecl stage"],
  restructuring: ["restructuring", "restructured", "forbearance"],
  defaulted: ["defaulted", "default flag", "credit impaired", "impaired", "npl", "non performing"]
};

const profiles = [
  { key: "overdraft", basePd: 0.07, lgd: 0.55, ccf: 0.75, stage: 1 },
  { key: "watch", basePd: 0.08, lgd: 0.5, ccf: 0.1, stage: 2 },
  { key: "special mention", basePd: 0.1, lgd: 0.52, ccf: 0.1, stage: 2 },
  { key: "substandard", basePd: 0.25, lgd: 0.6, ccf: 0.2, stage: 2 },
  { key: "doubtful", basePd: 0.5, lgd: 0.75, ccf: 0.5, stage: 3 },
  { key: "non performing", basePd: 1, lgd: 0.8, ccf: 0.5, stage: 3 },
  { key: "npl", basePd: 1, lgd: 0.8, ccf: 0.5, stage: 3 },
  { key: "lost", basePd: 1, lgd: 0.95, ccf: 0.5, stage: 3 },
  { key: "default", basePd: 1, lgd: 0.85, ccf: 0.5, stage: 3 },
  { key: "staff", basePd: 0.012, lgd: 0.3, ccf: 0, stage: 1 },
  { key: "lpo", basePd: 0.05, lgd: 0.45, ccf: 0.2, stage: 1 },
  { key: "trade", basePd: 0.06, lgd: 0.5, ccf: 0, stage: 2 },
  { key: "performing", basePd: 0.02, lgd: 0.45, ccf: 0, stage: 1 },
  { key: "standard", basePd: 0.02, lgd: 0.45, ccf: 0, stage: 1 },
  { key: "normal", basePd: 0.02, lgd: 0.45, ccf: 0, stage: 1 }
];

const collateralHaircuts = [
  ["cash", 0.95],
  ["deposit", 0.95],
  ["treasury", 0.9],
  ["government", 0.9],
  ["real estate", 0.7],
  ["property", 0.7],
  ["vehicle", 0.55],
  ["equipment", 0.5],
  ["inventory", 0.45],
  ["receivable", 0.45],
  ["guarantee", 0.35],
  ["unsecured", 0]
];

const state = {
  sourceName: "",
  pendingFile: null,
  pendingScorecardFile: null,
  rows: [],
  mappings: {},
  results: []
};

const els = {
  fileInput: document.querySelector("#fileInput"),
  scorecardInput: document.querySelector("#scorecardInput"),
  fileStatus: document.querySelector("#fileStatus"),
  scorecardStatus: document.querySelector("#scorecardStatus"),
  dropzone: document.querySelector("#dropzone"),
  scorecardDropzone: document.querySelector("#scorecardDropzone"),
  computeBtn: document.querySelector("#computeBtn"),
  exportBtn: document.querySelector("#exportBtn"),
  clearBtn: document.querySelector("#clearBtn"),
  loadSampleBtn: document.querySelector("#loadSampleBtn"),
  downloadTemplateBtn: document.querySelector("#downloadTemplateBtn"),
  fetchMacroBtn: document.querySelector("#fetchMacroBtn"),
  countrySelect: document.querySelector("#countrySelect"),
  inflationRate: document.querySelector("#inflationRate"),
  gdpGrowthRate: document.querySelector("#gdpGrowthRate"),
  macroStatus: document.querySelector("#macroStatus"),
  mappingList: document.querySelector("#mappingList"),
  resultsBody: document.querySelector("#resultsBody"),
  searchInput: document.querySelector("#searchInput"),
  toast: document.querySelector("#toast"),
  totalExposure: document.querySelector("#totalExposure"),
  totalEcl: document.querySelector("#totalEcl"),
  coverageRatio: document.querySelector("#coverageRatio"),
  rowsProcessed: document.querySelector("#rowsProcessed"),
  defaultDiscount: document.querySelector("#defaultDiscount"),
  stage2Dpd: document.querySelector("#stage2Dpd"),
  stage3Dpd: document.querySelector("#stage3Dpd"),
  autoStage: document.querySelector("#autoStage"),
  simplifiedApproach: document.querySelector("#simplifiedApproach"),
  undrawnExposure: document.querySelector("#undrawnExposure")
};

wireFileDrop(els.fileInput, els.dropzone, loadLoanFile);
wireFileDrop(els.scorecardInput, els.scorecardDropzone, loadScorecardFile);
els.computeBtn.addEventListener("click", computePortfolio);
els.exportBtn.addEventListener("click", exportResults);
els.clearBtn.addEventListener("click", clearAll);
els.loadSampleBtn.addEventListener("click", loadSample);
els.downloadTemplateBtn.addEventListener("click", downloadTemplate);
els.fetchMacroBtn.addEventListener("click", fetchMacroData);
els.searchInput.addEventListener("input", () => renderResults());

function wireFileDrop(input, dropzone, handler) {
  input.addEventListener("change", (event) => {
    const file = event.target.files?.[0];
    if (file) handler(file);
  });
  dropzone.addEventListener("dragover", (event) => {
    event.preventDefault();
    dropzone.classList.add("dragover");
  });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
  dropzone.addEventListener("drop", (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragover");
    const file = event.dataTransfer.files?.[0];
    if (file) handler(file);
  });
}

async function loadLoanFile(file) {
  state.pendingFile = file;
  state.sourceName = file.name;
  state.results = [];
  state.mappings = {};
  els.exportBtn.disabled = true;

  if (file.name.toLowerCase().endsWith(".csv")) {
    state.rows = parseCsv(await file.text());
    state.mappings = detectMappings(state.rows);
    els.fileStatus.textContent = `${state.rows.length} row${state.rows.length === 1 ? "" : "s"} loaded from ${file.name}.`;
    renderMappings();
    renderEmpty("Click Compute ECL to run the guided model.");
  } else {
    state.rows = [];
    els.fileStatus.textContent = `${file.name} is ready. Click Compute ECL to process the workbook.`;
    renderMappings();
    renderEmpty("Click Compute ECL to upload this workbook to the local IFRS 9 engine.");
  }
  updateMetrics([]);
}

function loadScorecardFile(file) {
  state.pendingScorecardFile = file;
  els.scorecardStatus.textContent = `${file.name} will override the default scorecard.`;
  toast("Optional scorecard loaded.");
}

function loadSample() {
  state.sourceName = "Guided sample loan schedule";
  state.pendingFile = null;
  state.rows = structuredClone(sampleRows);
  state.mappings = detectMappings(state.rows);
  state.results = [];
  els.fileStatus.textContent = `${state.rows.length} sample loans ready.`;
  els.exportBtn.disabled = true;
  renderMappings();
  renderEmpty("Click Compute ECL to run the guided model.");
  updateMetrics([]);
  toast("Sample data loaded.");
}

async function fetchMacroData() {
  if (!canUseBackend()) {
    toast("Macro fetch needs the local server.");
    return;
  }
  try {
    els.fetchMacroBtn.disabled = true;
    els.fetchMacroBtn.textContent = "Fetching...";
    const response = await fetch(`/api/macro?country=${encodeURIComponent(els.countrySelect.value)}`);
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Could not fetch macro data.");
    const inflation = payload.indicators?.inflation;
    const gdp = payload.indicators?.gdpGrowth;
    if (inflation?.value !== undefined) els.inflationRate.value = Number(inflation.value).toFixed(2);
    if (gdp?.value !== undefined) els.gdpGrowthRate.value = Number(gdp.value).toFixed(2);
    els.macroStatus.textContent = `Loaded ${payload.countryName}: inflation ${formatNumber(inflation?.value)}% (${inflation?.year}), GDP growth ${formatNumber(gdp?.value)}% (${gdp?.year}). Source: ${payload.source}.`;
    toast("Macro data loaded.");
  } catch (error) {
    els.macroStatus.textContent = "Macro data could not be fetched. You can type inflation and GDP growth manually.";
    toast(error.message || "Macro data could not be fetched.");
  } finally {
    els.fetchMacroBtn.disabled = false;
    els.fetchMacroBtn.textContent = "Fetch macro data";
  }
}

async function computePortfolio() {
  if (state.pendingFile && canUseBackend()) {
    await computeWithBackend(state.pendingFile);
    return;
  }
  if (!state.rows.length) {
    toast("Upload a loan schedule or load sample data first.");
    return;
  }
  const settings = readSettings();
  state.results = state.rows.map((row, index) => computeClientRow(row, index, settings));
  renderResults();
  renderMappings();
  updateMetrics(state.results);
  els.exportBtn.disabled = false;
  toast(`Computed ECL for ${state.results.length} loan${state.results.length === 1 ? "" : "s"}.`);
}

async function computeWithBackend(file) {
  const formData = new FormData();
  formData.append("file", file);
  if (state.pendingScorecardFile) formData.append("scorecard", state.pendingScorecardFile);
  formData.append("settings", JSON.stringify(readSettings()));

  try {
    els.computeBtn.disabled = true;
    els.computeBtn.textContent = "Computing...";
    const response = await fetch("/api/compute", { method: "POST", body: formData });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "Server computation failed.");
    state.rows = payload.rows || [];
    state.mappings = payload.mappings || {};
    state.results = payload.results || [];
    renderResults();
    renderMappings(payload.model);
    updateMetrics(state.results);
    els.fileStatus.textContent = `${state.results.length} loan${state.results.length === 1 ? "" : "s"} processed from ${file.name}.`;
    els.exportBtn.disabled = false;
    toast(`Computed ECL for ${state.results.length} loan${state.results.length === 1 ? "" : "s"}.`);
  } catch (error) {
    toast(error.message || "Unable to compute ECL.");
  } finally {
    els.computeBtn.disabled = false;
    els.computeBtn.textContent = "Compute ECL";
  }
}

function readSettings() {
  return {
    defaultDiscount: percentInput(els.defaultDiscount.value),
    inflationRate: percentInput(els.inflationRate.value),
    gdpGrowthRate: percentInput(els.gdpGrowthRate.value),
    stage2Dpd: numericInput(els.stage2Dpd.value, 30),
    stage3Dpd: numericInput(els.stage3Dpd.value, 90),
    autoStage: els.autoStage.checked,
    simplifiedApproach: els.simplifiedApproach.checked,
    includeUndrawn: els.undrawnExposure.checked
  };
}

function computeClientRow(row, index, settings) {
  const mappings = state.mappings;
  const value = (field) => getMappedValue(row, mappings, field);
  const customerId = cleanText(value("customerId")) || `Row ${index + 1}`;
  const customerName = cleanText(value("customerName")) || "Unassigned";
  const loanClass = cleanText(value("loanClassification")) || "Performing";
  const profile = profileFor(loanClass);
  const amount = numericInput(value("amount"), 0);
  const outstanding = numericInput(value("outstanding"), amount);
  const daysPastDue = numericInput(value("daysPastDue"), 0);
  const eir = percentValue(value("eir"), settings.defaultDiscount);
  const creditLimit = numericInput(value("creditLimit"), amount);
  const undrawn = numericInput(value("undrawnLimit"), Math.max(creditLimit - outstanding, 0));
  const ccf = percentValue(value("ccf"), profile.ccf);
  const ead = Math.max(outstanding + (settings.includeUndrawn ? undrawn * ccf : 0), 0);
  const stage = determineClientStage(value, loanClass, daysPastDue, settings, profile);
  const macroMultiplier = computeMacroMultiplier(settings.inflationRate, settings.gdpGrowthRate);
  const dpdMultiplier = daysPastDue <= 0 ? 0.8 : daysPastDue <= 30 ? 1 : daysPastDue <= 60 ? 2 : daysPastDue <= 90 ? 4 : 12;
  const scoreRatio = Math.max(0, Math.min(1, 1 - daysPastDue / 120));
  const pd12 = stage === 3 ? 1 : clamp(profile.basePd * (1 + (1 - scoreRatio) * 2) * dpdMultiplier * macroMultiplier, 0.001, 0.95);
  const dueDate = new Date(value("dueDate"));
  const lifeMonths = Number.isNaN(dueDate.getTime()) ? 12 : Math.max((dueDate.getFullYear() - new Date().getFullYear()) * 12 + dueDate.getMonth() - new Date().getMonth(), 1);
  const lifetimePd = stage === 3 ? 1 : clamp(1 - Math.pow(1 - pd12, Math.max(lifeMonths / 12, 1 / 12)), 0, 1);
  const collateralValue = numericInput(value("collateralValue"), 0);
  const haircut = collateralHaircut(value("collateralType"));
  const lgd = collateralValue > 0 && ead > 0 ? clamp(1 - Math.min(collateralValue * haircut, ead) / ead, 0.02, 1) : profile.lgd;
  const horizonMonths = stage === 1 ? Math.min(12, lifeMonths) : lifeMonths;
  const pdUsed = stage === 1 ? pd12 : lifetimePd;
  const discountFactor = eir > 0 ? 1 / Math.pow(1 + eir, horizonMonths / 12) : 1;
  const ecl = ead * pdUsed * lgd * discountFactor;
  return {
    assetId: customerId,
    customerId,
    borrower: customerName,
    customerName,
    assetClass: loanClass,
    loanClassification: loanClass,
    stage,
    stageLabel: `Stage ${stage}`,
    rationale: stage === 3 ? "Default or over 90 DPD" : stage === 2 ? "SICR or over 30 DPD" : "Performing",
    ead,
    eadBasis: undrawn > 0 && ccf > 0 ? "Outstanding plus CCF-adjusted undrawn exposure" : "Outstanding/amortised cost balance",
    exposureBase: outstanding,
    pd12,
    lifetimePd,
    pdUsed,
    pdBasis: "Classification + DPD + macro",
    scoreRatio,
    scoreBasis: "Fallback behavioural score from DPD",
    lgd,
    lgdBasis: collateralValue > 0 ? "Collateral haircut estimate" : "Default LGD by classification",
    discountRate: eir,
    horizonMonths,
    discountFactor,
    scenarioWeight: 1,
    macroMultiplier,
    ecl,
    daysPastDue,
    warnings: ""
  };
}

function determineClientStage(value, loanClass, daysPastDue, settings, profile) {
  const explicitStage = parseStage(value("stage"));
  if (explicitStage) return explicitStage;
  if (truthy(value("defaulted")) || daysPastDue > settings.stage3Dpd || profile.stage === 3) return 3;
  if (truthy(value("restructuring")) || daysPastDue > settings.stage2Dpd || profile.stage === 2) return 2;
  if (settings.simplifiedApproach && /trade|lease|contract|receivable/i.test(loanClass)) return 2;
  return 1;
}

function profileFor(loanClass) {
  const normalized = normalizeHeader(loanClass);
  return profiles.find((profile) => normalized.includes(profile.key)) || { basePd: 0.035, lgd: 0.5, ccf: 0, stage: 1 };
}

function collateralHaircut(value) {
  const normalized = normalizeHeader(value);
  return collateralHaircuts.find(([key]) => normalized.includes(key))?.[1] ?? (normalized ? 0.4 : 0);
}

function computeMacroMultiplier(inflationRate, gdpGrowthRate) {
  const multiplier = 1 + ((inflationRate * 100 - 10) * 0.025) - ((gdpGrowthRate * 100 - 3) * 0.04);
  return clamp(multiplier, 0.7, 1.8);
}

function detectMappings(rows) {
  const headers = Object.keys(rows[0] || {});
  const mappings = {};
  Object.entries(fieldAliases).forEach(([field, aliases]) => {
    const match = headers.find((header) => {
      const normalized = normalizeHeader(header);
      return aliases.map(normalizeHeader).some((alias) => normalized === alias || normalized.includes(alias));
    });
    if (match) mappings[field] = match;
  });
  return mappings;
}

function getMappedValue(row, mappings, field) {
  const mappedHeader = mappings[field];
  return mappedHeader ? row[mappedHeader] : "";
}

function renderMappings(model = null) {
  els.mappingList.innerHTML = "";
  if (!state.rows.length && !state.pendingFile) {
    els.mappingList.innerHTML = '<p class="muted">Mappings will appear after a file is loaded.</p>';
    return;
  }
  const entries = {
    "Customer ID": state.mappings.customerId,
    "Customer name": state.mappings.customerName,
    "Classification": state.mappings.loanClassification,
    "Amount": state.mappings.amount,
    "Outstanding / EAD": state.mappings.outstanding,
    "Due date": state.mappings.dueDate,
    "Days past due": state.mappings.daysPastDue,
    "EIR": state.mappings.eir,
    "Collateral value": state.mappings.collateralValue,
    "Collateral type": state.mappings.collateralType
  };
  Object.entries(entries).forEach(([label, mapped]) => {
    const item = document.createElement("div");
    item.className = "mapping-item";
    item.innerHTML = `<span>${escapeHtml(label)}</span><strong>${escapeHtml(mapped || "Not found")}</strong>`;
    els.mappingList.appendChild(item);
  });
  if (model) {
    const item = document.createElement("div");
    item.className = "mapping-item";
    item.innerHTML = `<span>Model</span><strong>${escapeHtml(model.scorecardRules)} scorecard rules</strong>`;
    els.mappingList.appendChild(item);
  }
}

function renderResults() {
  const search = els.searchInput.value.trim().toLowerCase();
  const rows = state.results.filter((row) => {
    if (!search) return true;
    return [row.customerId, row.customerName, row.loanClassification, row.stageLabel, row.rationale].join(" ").toLowerCase().includes(search);
  });
  if (!state.results.length) {
    renderEmpty("Upload a loan schedule or load sample data, then click Compute ECL.");
    return;
  }
  if (!rows.length) {
    renderEmpty("No computed loans match your search.");
    return;
  }
  els.resultsBody.innerHTML = rows
    .map(
      (row) => `
        <tr>
          <td>${escapeHtml(row.customerId || row.assetId)}<br><small>${escapeHtml(row.customerName || row.borrower)}</small></td>
          <td>${escapeHtml(row.loanClassification || row.assetClass)}</td>
          <td><span class="stage-pill stage-${row.stage}">${row.stageLabel}</span></td>
          <td>${formatMoney(row.ead)}<br><small>${escapeHtml(row.eadBasis || "")}</small></td>
          <td>${formatPercent(row.pd12)}</td>
          <td>${formatPercent(row.lifetimePd)}</td>
          <td>${formatPercent(row.lgd)}<br><small>${escapeHtml(row.lgdBasis || "")}</small></td>
          <td><strong>${formatMoney(row.ecl)}</strong></td>
          <td>${escapeHtml(row.pdBasis || row.rationale)}<br><small>${escapeHtml(row.rationale || "")}</small></td>
        </tr>
      `
    )
    .join("");
}

function renderEmpty(message) {
  els.resultsBody.innerHTML = `<tr><td colspan="9" class="empty-state">${escapeHtml(message)}</td></tr>`;
}

function updateMetrics(rows) {
  const totalExposure = rows.reduce((sum, row) => sum + row.ead, 0);
  const totalEcl = rows.reduce((sum, row) => sum + row.ecl, 0);
  els.totalExposure.textContent = formatMoney(totalExposure);
  els.totalEcl.textContent = formatMoney(totalEcl);
  els.coverageRatio.textContent = totalExposure > 0 ? formatPercent(totalEcl / totalExposure) : "0.00%";
  els.rowsProcessed.textContent = rows.length.toLocaleString();
}

async function exportResults() {
  if (!state.results.length) {
    toast("Compute ECL before exporting results.");
    return;
  }
  const exportRows = state.results.map((row) => ({
    "Customer ID": row.customerId || row.assetId,
    "Customer Name": row.customerName || row.borrower,
    "Loan Classification": row.loanClassification || row.assetClass,
    Stage: row.stageLabel,
    "Stage Rationale": row.rationale,
    "EAD/EOD": round(row.ead),
    "EAD Basis": row.eadBasis,
    "12M PD": row.pd12,
    "Lifetime PD": row.lifetimePd,
    "PD Used": row.pdUsed,
    "PD Basis": row.pdBasis,
    "Score Ratio": row.scoreRatio,
    LGD: row.lgd,
    "LGD Basis": row.lgdBasis,
    "Discount Rate": row.discountRate,
    "Horizon Months": row.horizonMonths,
    "Macro Multiplier": row.macroMultiplier,
    ECL: round(row.ecl),
    "Days Past Due": row.daysPastDue
  }));
  if (canUseBackend()) {
    const response = await fetch("/api/export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ rows: exportRows })
    });
    if (response.ok) {
      await saveResponseBlob(response, "ifrs9-ecl-results.xlsx");
      return;
    }
  }
  downloadText("ifrs9-ecl-results.csv", toCsv(exportRows), "text/csv");
}

async function downloadTemplate() {
  if (canUseBackend()) {
    const response = await fetch("/api/template");
    if (response.ok) {
      await saveResponseBlob(response, "ifrs9-guided-ecl-template.xlsx");
      return;
    }
  }
  downloadText("ifrs9-guided-ecl-template.csv", toCsv(sampleRows), "text/csv");
}

function clearAll() {
  state.sourceName = "";
  state.pendingFile = null;
  state.pendingScorecardFile = null;
  state.rows = [];
  state.mappings = {};
  state.results = [];
  els.fileInput.value = "";
  els.scorecardInput.value = "";
  els.fileStatus.textContent = "No file loaded yet.";
  els.scorecardStatus.textContent = "Default scorecard will be used.";
  els.searchInput.value = "";
  els.exportBtn.disabled = true;
  renderMappings();
  renderEmpty("Upload a loan schedule or load sample data, then click Compute ECL.");
  updateMetrics([]);
}

function canUseBackend() {
  return window.location.protocol === "http:" || window.location.protocol === "https:";
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let value = "";
  let insideQuotes = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (char === '"' && next === '"') {
      value += '"';
      index += 1;
    } else if (char === '"') {
      insideQuotes = !insideQuotes;
    } else if (char === "," && !insideQuotes) {
      row.push(value);
      value = "";
    } else if ((char === "\n" || char === "\r") && !insideQuotes) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
    } else {
      value += char;
    }
  }
  if (value || row.length) {
    row.push(value);
    rows.push(row);
  }
  const headers = rows.shift()?.map((header) => header.trim()) || [];
  return rows
    .filter((line) => line.some((cell) => String(cell).trim() !== ""))
    .map((line) =>
      headers.reduce((record, header, index) => {
        record[header || `Column ${index + 1}`] = line[index] ?? "";
        return record;
      }, {})
    );
}

function toCsv(rows) {
  if (!rows.length) return "";
  const headers = Object.keys(rows[0]);
  return [headers.join(","), ...rows.map((row) => headers.map((header) => csvEscape(row[header])).join(","))].join("\n");
}

function csvEscape(value) {
  const text = String(value ?? "");
  return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

async function saveResponseBlob(response, fallbackFilename) {
  const blob = await response.blob();
  const header = response.headers.get("Content-Disposition") || "";
  const match = header.match(/filename="?([^"]+)"?/i);
  const filename = match?.[1] || fallbackFilename;
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function downloadText(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[%()]/g, "")
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ");
}

function numericInput(value, fallback = 0) {
  if (typeof value === "number") return Number.isFinite(value) ? value : fallback;
  const cleaned = String(value ?? "").replace(/,/g, "").replace(/%/g, "").trim();
  if (!cleaned) return fallback;
  const number = Number(cleaned);
  return Number.isFinite(number) ? number : fallback;
}

function percentInput(value) {
  return clamp(numericInput(value, 0) / 100, -1, 1);
}

function percentValue(value, fallback) {
  if (value === undefined || value === null || String(value).trim() === "") return fallback;
  const raw = numericInput(value, NaN);
  if (!Number.isFinite(raw)) return fallback;
  return clamp(raw > 1 ? raw / 100 : raw, 0, 1);
}

function parseStage(value) {
  const match = cleanText(value).match(/[123]/);
  return match ? Number(match[0]) : null;
}

function truthy(value) {
  return ["yes", "y", "true", "1", "default", "defaulted", "impaired", "npl", "non performing", "stage 3"].includes(cleanText(value).toLowerCase());
}

function cleanText(value) {
  return String(value ?? "").trim();
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function formatMoney(value) {
  return new Intl.NumberFormat(undefined, { maximumFractionDigits: 0 }).format(value || 0);
}

function formatNumber(value) {
  return Number(value || 0).toFixed(2);
}

function formatPercent(value) {
  return new Intl.NumberFormat(undefined, {
    style: "percent",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(value || 0);
}

function round(value) {
  return Math.round((value + Number.EPSILON) * 100) / 100;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function toast(message) {
  els.toast.textContent = message;
  els.toast.classList.add("show");
  window.clearTimeout(toast.timer);
  toast.timer = window.setTimeout(() => els.toast.classList.remove("show"), 3600);
}
