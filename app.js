const OUTPUT_HEADERS = [
  "Referencia del pedido - TECH",
  "Proveedor",
  "Cliente",
  "Fecha inicio",
  "Total",
  "Moneda",
  "EQ. USD",
  "Probabilidad",
  "SECTOR",
];

const state = {
  files: [],
  outputRows: [],
  traceRows: [],
  warnings: [],
  filters: {
    search: "",
    tech: "",
    currency: "",
  },
  traceFilters: {
    search: "",
    included: "",
    providerSource: "",
  },
};

const els = {
  fileInput: document.querySelector("#fileInput"),
  processButton: document.querySelector("#processButton"),
  downloadButton: document.querySelector("#downloadButton"),
  downloadTraceButton: document.querySelector("#downloadTraceButton"),
  clearButton: document.querySelector("#clearButton"),
  dropZone: document.querySelector("#dropZone"),
  fileList: document.querySelector("#fileList"),
  resultBody: document.querySelector("#resultBody"),
  warnings: document.querySelector("#warnings"),
  fileCount: document.querySelector("#fileCount"),
  rowCount: document.querySelector("#rowCount"),
  warningCount: document.querySelector("#warningCount"),
  grandTotal: document.querySelector("#grandTotal"),
  searchInput: document.querySelector("#searchInput"),
  techFilter: document.querySelector("#techFilter"),
  currencyFilter: document.querySelector("#currencyFilter"),
  resetFiltersButton: document.querySelector("#resetFiltersButton"),
  filteredCount: document.querySelector("#filteredCount"),
  traceBody: document.querySelector("#traceBody"),
  traceCount: document.querySelector("#traceCount"),
  traceSearchInput: document.querySelector("#traceSearchInput"),
  traceIncludedFilter: document.querySelector("#traceIncludedFilter"),
  traceProviderSourceFilter: document.querySelector("#traceProviderSourceFilter"),
  resetTraceFiltersButton: document.querySelector("#resetTraceFiltersButton"),
};

els.fileInput.addEventListener("change", (event) => {
  addFiles(Array.from(event.target.files || []));
  els.fileInput.value = "";
});

els.processButton.addEventListener("click", async () => {
  await processSelectedFiles();
});

els.downloadButton.addEventListener("click", () => {
  downloadWorkbook();
});

els.downloadTraceButton.addEventListener("click", () => {
  downloadTraceWorkbook();
});

els.clearButton.addEventListener("click", () => {
  state.files = [];
  state.outputRows = [];
  state.traceRows = [];
  state.warnings = [];
  resetFilters();
  resetTraceFilters();
  render();
});

els.searchInput.addEventListener("input", (event) => {
  state.filters.search = event.target.value;
  renderResults();
  renderTrace();
  renderFilterState();
});

els.techFilter.addEventListener("change", (event) => {
  state.filters.tech = event.target.value;
  renderResults();
  renderTrace();
  renderFilterState();
});

els.currencyFilter.addEventListener("change", (event) => {
  state.filters.currency = event.target.value;
  renderResults();
  renderTrace();
  renderFilterState();
});

els.resetFiltersButton.addEventListener("click", () => {
  resetFilters();
  renderFilterControls();
  renderResults();
  renderTrace();
  renderFilterState();
});

els.traceSearchInput.addEventListener("input", (event) => {
  state.traceFilters.search = event.target.value;
  renderTrace();
  renderTraceFilterState();
});

els.traceIncludedFilter.addEventListener("change", (event) => {
  state.traceFilters.included = event.target.value;
  renderTrace();
  renderTraceFilterState();
});

els.traceProviderSourceFilter.addEventListener("change", (event) => {
  state.traceFilters.providerSource = event.target.value;
  renderTrace();
  renderTraceFilterState();
});

els.resetTraceFiltersButton.addEventListener("click", () => {
  resetTraceFilters();
  renderTraceFilterControls();
  renderTrace();
  renderTraceFilterState();
});

["dragenter", "dragover"].forEach((eventName) => {
  els.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.dropZone.classList.add("dragging");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  els.dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.dropZone.classList.remove("dragging");
  });
});

els.dropZone.addEventListener("drop", (event) => {
  addFiles(Array.from(event.dataTransfer.files || []));
});

function addFiles(files) {
  const excelFiles = files.filter((file) => /\.(xlsx|xlsm|xls)$/i.test(file.name) && !file.name.startsWith("~$"));
  const known = new Set(state.files.map((file) => `${file.name}|${file.size}|${file.lastModified}`));

  for (const file of excelFiles) {
    const key = `${file.name}|${file.size}|${file.lastModified}`;
    if (!known.has(key)) {
      state.files.push(file);
      known.add(key);
    }
  }

  render();
}

async function processSelectedFiles() {
  if (!window.XLSX) {
    state.warnings = ["No se pudo cargar la libreria XLSX. Revise la conexion a internet o use una copia local de la libreria."];
    render();
    return;
  }

  state.outputRows = [];
  state.traceRows = [];
  state.warnings = [];
  const groups = new Map();

  els.processButton.disabled = true;
  els.processButton.textContent = "Procesando...";

  for (const file of state.files) {
    try {
      const result = await processWorkbookFile(file);
      for (const warning of result.warnings) {
        state.warnings.push(warning);
      }
      for (const traceRow of result.traceRows) {
        state.traceRows.push(traceRow);
      }
      for (const line of result.lines) {
        const key = [line.tech, normalizeProvider(line.provider), line.currency].join("||");
        const existing = groups.get(key) || {
          tech: line.tech,
          provider: line.provider,
          client: line.client || "",
          total: 0,
          currency: line.currency || "",
          itemCount: 0,
        };
        existing.total += line.total;
        existing.itemCount += 1;
        if (!existing.client && line.client) {
          existing.client = line.client;
        }
        groups.set(key, existing);
      }
    } catch (error) {
      state.warnings.push(`${file.name}: no se pudo procesar (${error.message}).`);
    }
  }

  state.outputRows = Array.from(groups.values())
    .sort((a, b) => a.tech.localeCompare(b.tech) || a.provider.localeCompare(b.provider))
    .map((row) => ({
      "Referencia del pedido - TECH": row.tech,
      Proveedor: row.provider,
      Cliente: row.client,
      "Fecha inicio": "",
      Total: roundMoney(row.total),
      Moneda: row.currency,
      "EQ. USD": "",
      Probabilidad: "",
      SECTOR: "CORPORATIVO",
      _itemCount: row.itemCount,
    }));

  resetFilters();
  resetTraceFilters();
  els.processButton.textContent = "Procesar";
  render();
}

async function processWorkbookFile(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellNF: true,
    raw: true,
  });

  const warnings = [];
  const fileTechs = extractTechs(file.name);
  if (fileTechs.length === 0) {
    warnings.push(`${file.name}: no se encontro TECH en el nombre del archivo.`);
  }
  if (fileTechs.length > 1) {
    warnings.push(`${file.name}: se encontraron varios TECH (${fileTechs.join(", ")}). Se intenta detectar TECH por fila; si no aparece, se usa ${fileTechs[0]}.`);
  }

  const pclientes = findBestClientSheet(workbook);
  if (!pclientes) {
    warnings.push(`${file.name}: no se encontro una hoja PClientes utilizable.`);
    return { lines: [], traceRows: [], warnings };
  }

  const bomTables = findBomTables(workbook, pclientes.sheetName);
  if (bomTables.length === 0) {
    warnings.push(`${file.name}: no se encontro una hoja BOM utilizable. Se usara proveedor desde PClientes si existe.`);
  }

  const clientName = detectClientName(file.name);
  const lines = [];
  const traceRows = [];
  let missingProvider = 0;
  let missingTotal = 0;

  for (const item of pclientes.items) {
    const rowTechs = extractTechs(item.rawText);
    const tech = rowTechs[0] || fileTechs[0] || "TECH PENDIENTE";
    const providerMatch = item.provider
      ? { provider: item.provider, source: "PClientes", detail: pclientes.sheetName }
      : findProviderForItem(item, bomTables);
    const provider = providerMatch.provider || "PROVEEDOR PENDIENTE";
    const amount = calculateLineAmount(item);
    const total = amount.total;
    const traceBase = {
      tech,
      provider,
      client: clientName,
      total: Number.isFinite(total) ? total : "",
      currency: item.currency || detectCurrency(item.rawText),
      sourceFile: file.name,
      clientSheet: pclientes.sheetName,
      item: item.item,
      part: item.part,
      description: item.description,
      quantity: Number.isFinite(item.quantity) ? item.quantity : "",
      unit: Number.isFinite(item.unit) ? item.unit : "",
      costAsuUnit: Number.isFinite(item.costAsuUnit) ? item.costAsuUnit : "",
      costAsuSubtotal: Number.isFinite(item.costAsuSubtotal) ? item.costAsuSubtotal : "",
      calculation: amount.detail,
      providerSource: providerMatch.source || "Pendiente",
      providerMatchDetail: providerMatch.detail || "Sin match",
    };

    if (provider === "PROVEEDOR PENDIENTE") {
      missingProvider += 1;
    }
    if (!Number.isFinite(total) || total === 0) {
      missingTotal += 1;
      traceRows.push({ ...traceBase, included: "NO" });
      continue;
    }

    const line = { ...traceBase, included: "SI", total };
    lines.push(line);
    traceRows.push(line);
  }

  if (missingProvider > 0) {
    warnings.push(`${file.name}: ${missingProvider} item(s) quedaron sin proveedor.`);
  }
  if (missingTotal > 0) {
    warnings.push(`${file.name}: ${missingTotal} item(s) fueron omitidos por no tener Costo Asu sub-Total numerico.`);
  }
  if (pclientes.items.length === 0) {
    warnings.push(`${file.name}: la hoja ${pclientes.sheetName} no produjo items validos.`);
  }

  warnings.push(`${file.name}: hoja principal usada ${pclientes.sheetName}; ${lines.length} item(s) validos.`);
  return { lines, traceRows, warnings };
}

function findBestClientSheet(workbook) {
  const candidates = workbook.SheetNames
    .filter((name) => normalizeText(name).includes("pclientes"))
    .map((sheetName) => {
      const rows = sheetRows(workbook.Sheets[sheetName]);
      const tables = extractTables(rows, "client");
      const items = tables.flatMap((table) => table.items);
      return {
        sheetName,
        rows,
        tables,
        items,
        score: items.length * 10 + tables.length * 5 + (normalizeText(sheetName) === "pclientes" ? 4 : 0),
      };
    })
    .filter((candidate) => candidate.tables.length > 0);

  candidates.sort((a, b) => b.score - a.score);
  return candidates[0] || null;
}

function findBomTables(workbook, clientSheetName) {
  const tables = [];
  for (const sheetName of workbook.SheetNames) {
    const normalized = normalizeText(sheetName);
    if (sheetName === clientSheetName) {
      continue;
    }
    if (!normalized.includes("bom") && !normalized.includes("infra")) {
      continue;
    }
    const rows = sheetRows(workbook.Sheets[sheetName]);
    for (const table of extractTables(rows, "bom")) {
      tables.push({ ...table, sheetName });
    }
  }
  return tables;
}

function sheetRows(sheet) {
  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    blankrows: false,
    defval: null,
    raw: true,
  });
}

function extractTables(rows, type) {
  const tables = [];
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const header = mapHeader(rows[rowIndex], type);
    if (!header) {
      continue;
    }
    Object.assign(header, mapAuxiliaryHeader(rows, rowIndex));

    const items = [];
    for (let index = rowIndex + 1; index < rows.length; index += 1) {
      const row = rows[index] || [];
      if (mapHeader(row, type)) {
        break;
      }
      if (isSummaryOrEmptyRow(row)) {
        if (items.length > 0 && countUsefulCells(row) === 0) {
          break;
        }
        continue;
      }

      const item = rowToItem(row, header);
      if (isValidItem(item)) {
        items.push(item);
      }
    }

    if (items.length > 0) {
      tables.push({ header, startRow: rowIndex + 1, items });
    }
  }
  return tables;
}

function mapHeader(row, type) {
  const map = {};
  const labels = row.map((cell) => normalizeText(cell));

  labels.forEach((label, index) => {
    if (!label) {
      return;
    }
    const compactLabel = label.replace(/[^a-z0-9]/g, "");
    if (label === "item") {
      map.item = index;
    } else if (["nro de parte", "nro parte", "numero de parte", "codigo", "codigo o sku", "sku", "codigo sku"].includes(label)) {
      map.part = index;
    } else if (label.includes("descrip") || label === "producto") {
      map.description = index;
    } else if (["cant", "cantidad"].includes(label) || compactLabel === "cant") {
      map.quantity = index;
    } else if (label === "unit" || compactLabel === "unit" || label.includes("precio unitario") || label.includes("costo unitario")) {
      map.unit = index;
    } else if (label.includes("sub total") || label.includes("subtotal") || compactLabel.includes("subtotal") || label.includes("precio total") || label.includes("costo total")) {
      map.total = index;
    } else if (label.includes("proveedor")) {
      map.provider = index;
    } else if (label.includes("divisa") || label.includes("moneda")) {
      map.currency = index;
    }
  });

  const hasCore = map.item !== undefined && (map.total !== undefined || map.unit !== undefined) && (map.part !== undefined || map.description !== undefined);
  const hasProviderBom = type === "bom" && map.provider !== undefined && (map.item !== undefined || map.part !== undefined || map.description !== undefined);
  const hasClient = type === "client" && hasCore;

  if (!hasClient && !hasProviderBom) {
    return null;
  }
  if (type === "client" && map.provider === undefined && map.total !== undefined) {
    map.provider = map.total + 1;
  }
  return map;
}

function mapAuxiliaryHeader(rows, headerRowIndex) {
  const map = {};
  const start = 0;
  for (let rowIndex = start; rowIndex < headerRowIndex; rowIndex += 1) {
    const labels = (rows[rowIndex] || []).map((cell) => normalizeText(cell));
    labels.forEach((label, index) => {
      const compactLabel = label.replace(/[^a-z0-9]/g, "");
      if (label.includes("costo asu") && label.includes("unitario")) {
        map.costAsuUnit = index;
      } else if (label.includes("costo asu") && (label.includes("subtotal") || label.includes("sub total") || compactLabel.includes("subtotal"))) {
        map.costAsuSubtotal = index;
      }
    });
  }
  return map;
}

function rowToItem(row, header) {
  const rawText = row.map((cell) => stringify(cell)).join(" ");
  const quantity = toNumber(valueAt(row, header.quantity));
  const unit = toNumber(valueAt(row, header.unit));
  const total = toNumber(valueAt(row, header.total));
  const costAsuUnit = toNumber(valueAt(row, header.costAsuUnit));
  const costAsuSubtotal = toNumber(valueAt(row, header.costAsuSubtotal));

  return {
    item: stringify(valueAt(row, header.item)),
    part: stringify(valueAt(row, header.part)),
    description: stringify(valueAt(row, header.description)),
    quantity,
    unit,
    total,
    costAsuUnit,
    costAsuSubtotal,
    provider: cleanProvider(valueAt(row, header.provider)),
    currency: cleanCurrency(valueAt(row, header.currency)) || detectCurrency(rawText),
    rawText,
  };
}

function isValidItem(item) {
  if (!item.item && !item.part && !item.description) {
    return false;
  }
  if (/^(total|subtotal|resumen|costos)/i.test(item.item) || /resumen ejecutivo/i.test(item.rawText)) {
    return false;
  }
  return Number.isFinite(item.total) || Number.isFinite(item.unit) || item.provider;
}

function isSummaryOrEmptyRow(row) {
  const text = normalizeText(row.map((cell) => stringify(cell)).join(" "));
  return !text || text.includes("resumen ejecutivo") || text === "total" || text.startsWith("costos ");
}

function countUsefulCells(row) {
  return row.filter((cell) => stringify(cell).trim() !== "").length;
}

function findProviderForItem(item, bomTables) {
  const itemKey = normalizeKey(item.item);
  const partKey = normalizeKey(item.part);
  const descKey = normalizeKey(item.description);
  const exactPartMatches = [];
  const exactDescriptionMatches = [];
  const itemDescriptionMatches = [];
  const itemOnlyMatches = [];

  for (const table of bomTables) {
    for (const bomItem of table.items) {
      if (!bomItem.provider) {
        continue;
      }

      const bomItemKey = normalizeKey(bomItem.item);
      const bomPartKey = normalizeKey(bomItem.part);
      const bomDescKey = normalizeKey(bomItem.description);
      const sameItem = itemKey && bomItemKey === itemKey;
      const samePart = partKey && bomPartKey === partKey;
      const sameDesc = descKey && bomDescKey === descKey;
      const similarDesc = descriptionsLookRelated(descKey, bomDescKey);
      const matchBase = {
        provider: bomItem.provider,
        source: "BOM",
        sheetName: table.sheetName,
      };

      if (samePart) {
        exactPartMatches.push({ ...matchBase, detail: `${table.sheetName}: Nro de Parte` });
      } else if (sameDesc) {
        exactDescriptionMatches.push({ ...matchBase, detail: `${table.sheetName}: Descripcion exacta` });
      } else if (sameItem && similarDesc) {
        itemDescriptionMatches.push({ ...matchBase, detail: `${table.sheetName}: Item + descripcion similar` });
      } else if (sameItem) {
        itemOnlyMatches.push({ ...matchBase, detail: `${table.sheetName}: Item solamente` });
      }
    }
  }

  return firstUniqueProvider(exactPartMatches)
    || firstUniqueProvider(exactDescriptionMatches)
    || firstUniqueProvider(itemDescriptionMatches)
    || firstUniqueProvider(itemOnlyMatches)
    || {};
}

function firstUniqueProvider(matches) {
  const unique = Array.from(new Set(matches.map((match) => normalizeComparableProvider(match.provider))));
  if (unique.length !== 1) {
    return null;
  }
  return matches[0];
}

function normalizeComparableProvider(value) {
  return normalizeText(value).replace(/\s+/g, " ");
}

function descriptionsLookRelated(a, b) {
  if (!a || !b) {
    return false;
  }
  if (a.includes(b) || b.includes(a)) {
    return true;
  }
  const aTokens = new Set(a.match(/[a-z0-9]{4,}/g) || []);
  const bTokens = new Set(b.match(/[a-z0-9]{4,}/g) || []);
  let common = 0;
  for (const token of aTokens) {
    if (bTokens.has(token)) {
      common += 1;
    }
  }
  return common >= 2;
}

function calculateLineAmount(item) {
  if (Number.isFinite(item.costAsuSubtotal)) {
    return {
      total: item.costAsuSubtotal * 1.1,
      detail: `Costo Asu sub-Total (${item.costAsuSubtotal}) x 1,1`,
    };
  }
  return {
    total: NaN,
    detail: "Omitido: sin Costo Asu sub-Total numerico",
  };
}

function downloadWorkbook() {
  if (!state.outputRows.length) {
    return;
  }

  const dataRows = state.outputRows.map((row) => OUTPUT_HEADERS.map((header) => row[header] ?? ""));
  const worksheet = XLSX.utils.aoa_to_sheet([
    [null, null, null, null, "TC", 6000, 0, null, null],
    OUTPUT_HEADERS,
    ...dataRows,
  ]);

  worksheet["!cols"] = [
    { wch: 28 },
    { wch: 28 },
    { wch: 28 },
    { wch: 14 },
    { wch: 16 },
    { wch: 12 },
    { wch: 12 },
    { wch: 14 },
    { wch: 16 },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "BASE COMPRAS");
  XLSX.writeFile(workbook, `base_compras_corporativo_${dateStamp()}.xlsx`);
}

function downloadTraceWorkbook() {
  if (!state.traceRows.length) {
    return;
  }

  const headers = [
    "TECH",
    "Incluido en suma",
    "Proveedor agrupado",
    "Total item",
    "Moneda",
    "Archivo",
    "Hoja PClientes",
    "Item",
    "Nro de Parte",
    "Descripcion",
    "Cant.",
    "Unit.",
    "Costo Asu Unitario",
    "Costo Asu sub-Total",
    "Calculo aplicado",
    "Fuente proveedor",
    "Hoja/criterio match",
  ];
  const rows = getVisibleTraceRows().map((row) => [
    row.tech,
    row.included,
    row.provider,
    row.total === "" ? "" : roundMoney(row.total),
    row.currency,
    row.sourceFile,
    row.clientSheet,
    row.item,
    row.part,
    row.description,
    row.quantity,
    row.unit,
    row.costAsuUnit,
    row.costAsuSubtotal,
    row.calculation,
    row.providerSource,
    row.providerMatchDetail,
  ]);
  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  worksheet["!cols"] = [
    { wch: 14 },
    { wch: 14 },
    { wch: 28 },
    { wch: 14 },
    { wch: 10 },
    { wch: 42 },
    { wch: 24 },
    { wch: 10 },
    { wch: 22 },
    { wch: 46 },
    { wch: 10 },
    { wch: 12 },
    { wch: 18 },
    { wch: 18 },
    { wch: 42 },
    { wch: 18 },
    { wch: 30 },
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "TRAZABILIDAD");
  XLSX.writeFile(workbook, `trazabilidad_corporativo_${dateStamp()}.xlsx`);
}

function render() {
  els.fileCount.textContent = String(state.files.length);
  els.rowCount.textContent = String(state.outputRows.length);
  els.warningCount.textContent = String(state.warnings.length);
  els.grandTotal.textContent = formatMoney(state.outputRows.reduce((sum, row) => sum + (Number(row.Total) || 0), 0));

  els.processButton.disabled = state.files.length === 0;
  els.clearButton.disabled = state.files.length === 0 && state.outputRows.length === 0 && state.traceRows.length === 0 && state.warnings.length === 0;
  els.downloadButton.disabled = state.outputRows.length === 0;
  els.downloadTraceButton.disabled = state.traceRows.length === 0;

  renderFiles();
  renderWarnings();
  renderFilterControls();
  renderTraceFilterControls();
  renderFilterState();
  renderTraceFilterState();
  renderResults();
  renderTrace();
}

function renderFiles() {
  els.fileList.innerHTML = "";
  if (state.files.length === 0) {
    const li = document.createElement("li");
    li.innerHTML = "<span>No hay archivos cargados.</span>";
    els.fileList.appendChild(li);
    return;
  }

  for (const file of state.files) {
    const li = document.createElement("li");
    const name = document.createElement("span");
    const size = document.createElement("small");
    name.textContent = file.name;
    size.textContent = formatBytes(file.size);
    li.append(name, size);
    els.fileList.appendChild(li);
  }
}

function renderWarnings() {
  els.warnings.innerHTML = "";
  els.warnings.classList.toggle("empty", state.warnings.length === 0);

  if (state.warnings.length === 0) {
    els.warnings.textContent = "Sin advertencias.";
    return;
  }

  for (const warning of state.warnings) {
    const div = document.createElement("div");
    div.className = "warning-item";
    div.textContent = warning;
    els.warnings.appendChild(div);
  }
}

function renderResults() {
  els.resultBody.innerHTML = "";
  const visibleRows = getVisibleRows();
  els.filteredCount.textContent = `${visibleRows.length} visibles`;

  if (state.outputRows.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.className = "empty-cell";
    cell.colSpan = OUTPUT_HEADERS.length;
    cell.textContent = "Todavia no hay datos procesados.";
    row.appendChild(cell);
    els.resultBody.appendChild(row);
    return;
  }

  if (visibleRows.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.className = "empty-cell";
    cell.colSpan = OUTPUT_HEADERS.length;
    cell.textContent = "No hay filas que coincidan con los filtros.";
    row.appendChild(cell);
    els.resultBody.appendChild(row);
    return;
  }

  for (const outputRow of visibleRows) {
    const tr = document.createElement("tr");
    for (const header of OUTPUT_HEADERS) {
      const td = document.createElement("td");
      const value = outputRow[header] ?? "";
      td.textContent = header === "Total" ? formatMoney(value) : value;
      if (header === "Total") {
        td.className = "number";
      }
      tr.appendChild(td);
    }
    els.resultBody.appendChild(tr);
  }
}

function renderTrace() {
  els.traceBody.innerHTML = "";
  const visibleRows = getVisibleTraceRows();
  els.traceCount.textContent = `${visibleRows.length} items fuente`;

  if (state.traceRows.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.className = "empty-cell";
    cell.colSpan = 17;
    cell.textContent = "Todavia no hay trazabilidad.";
    row.appendChild(cell);
    els.traceBody.appendChild(row);
    return;
  }

  if (visibleRows.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.className = "empty-cell";
    cell.colSpan = 17;
    cell.textContent = "No hay items fuente que coincidan con los filtros.";
    row.appendChild(cell);
    els.traceBody.appendChild(row);
    return;
  }

  for (const trace of visibleRows) {
    const tr = document.createElement("tr");
    const values = [
      trace.tech,
      trace.included,
      trace.provider,
      trace.total === "" ? "" : formatMoney(trace.total),
      trace.currency,
      trace.sourceFile,
      trace.clientSheet,
      trace.item,
      trace.part,
      trace.description,
      trace.quantity,
      trace.unit === "" ? "" : formatMoney(trace.unit),
      trace.costAsuUnit === "" ? "" : formatMoney(trace.costAsuUnit),
      trace.costAsuSubtotal === "" ? "" : formatMoney(trace.costAsuSubtotal),
      trace.calculation,
      trace.providerSource,
      trace.providerMatchDetail,
    ];
    values.forEach((value, index) => {
      const td = document.createElement("td");
      td.textContent = value ?? "";
      if ([3, 11, 12, 13].includes(index)) {
        td.className = "number";
      }
      tr.appendChild(td);
    });
    els.traceBody.appendChild(tr);
  }
}

function renderFilterControls() {
  const currentTech = state.filters.tech;
  const currentCurrency = state.filters.currency;

  replaceOptions(els.techFilter, uniqueSorted(state.outputRows.map((row) => row["Referencia del pedido - TECH"])), "Todos");
  replaceOptions(els.currencyFilter, uniqueSorted(state.outputRows.map((row) => row.Moneda || "SIN MONEDA")), "Todas");

  els.searchInput.value = state.filters.search;
  els.techFilter.value = currentTech;
  els.currencyFilter.value = currentCurrency;

  if (els.techFilter.value !== currentTech) {
    state.filters.tech = "";
  }
  if (els.currencyFilter.value !== currentCurrency) {
    state.filters.currency = "";
  }
}

function renderFilterState() {
  const hasFilters = Boolean(state.filters.search.trim() || state.filters.tech || state.filters.currency);
  els.resetFiltersButton.disabled = !hasFilters;
}

function renderTraceFilterControls() {
  const currentSource = state.traceFilters.providerSource;
  replaceOptions(
    els.traceProviderSourceFilter,
    uniqueSorted(state.traceRows.map((row) => row.providerSource || "SIN FUENTE")),
    "Todas",
  );
  els.traceSearchInput.value = state.traceFilters.search;
  els.traceIncludedFilter.value = state.traceFilters.included;
  els.traceProviderSourceFilter.value = currentSource;

  if (els.traceProviderSourceFilter.value !== currentSource) {
    state.traceFilters.providerSource = "";
  }
}

function renderTraceFilterState() {
  const hasFilters = Boolean(
    state.traceFilters.search.trim()
      || state.traceFilters.included
      || state.traceFilters.providerSource,
  );
  els.resetTraceFiltersButton.disabled = !hasFilters;
}

function replaceOptions(select, values, emptyLabel) {
  select.innerHTML = "";
  const empty = document.createElement("option");
  empty.value = "";
  empty.textContent = emptyLabel;
  select.appendChild(empty);

  for (const value of values) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  }
}

function getVisibleRows() {
  const search = normalizeText(state.filters.search);
  const tech = state.filters.tech;
  const currency = state.filters.currency;

  return state.outputRows.filter((row) => {
    const rowCurrency = row.Moneda || "SIN MONEDA";
    const textMatch = !search || normalizeText(OUTPUT_HEADERS.map((header) => row[header]).join(" ")).includes(search);
    const techMatch = !tech || row["Referencia del pedido - TECH"] === tech;
    const currencyMatch = !currency || rowCurrency === currency;
    return textMatch && techMatch && currencyMatch;
  });
}

function getVisibleTraceRows() {
  const search = normalizeText(state.filters.search);
  const tech = state.filters.tech;
  const currency = state.filters.currency;
  const traceSearch = normalizeText(state.traceFilters.search);
  const included = state.traceFilters.included;
  const providerSource = state.traceFilters.providerSource;

  return state.traceRows.filter((row) => {
    const rowCurrency = row.currency || "SIN MONEDA";
    const searchable = [
      row.tech,
      row.included,
      row.provider,
      row.currency,
      row.sourceFile,
      row.clientSheet,
      row.item,
      row.part,
      row.description,
      row.costAsuUnit,
      row.costAsuSubtotal,
      row.calculation,
      row.providerSource,
      row.providerMatchDetail,
    ].join(" ");
    const textMatch = !search || normalizeText(searchable).includes(search);
    const techMatch = !tech || row.tech === tech;
    const currencyMatch = !currency || rowCurrency === currency;
    const traceTextMatch = !traceSearch || normalizeText(searchable).includes(traceSearch);
    const includedMatch = !included || row.included === included;
    const sourceMatch = !providerSource || (row.providerSource || "SIN FUENTE") === providerSource;
    return textMatch && techMatch && currencyMatch && traceTextMatch && includedMatch && sourceMatch;
  });
}

function resetFilters() {
  state.filters.search = "";
  state.filters.tech = "";
  state.filters.currency = "";
}

function resetTraceFilters() {
  state.traceFilters.search = "";
  state.traceFilters.included = "";
  state.traceFilters.providerSource = "";
}

function uniqueSorted(values) {
  return Array.from(new Set(values.filter(Boolean))).sort((a, b) => String(a).localeCompare(String(b)));
}

function extractTechs(text) {
  const matches = [];
  const pattern = /TECH[-\s]?(\d+)/gi;
  let match = pattern.exec(text || "");
  while (match) {
    const tech = `TECH${match[1]}`;
    if (!matches.includes(tech)) {
      matches.push(tech);
    }
    match = pattern.exec(text || "");
  }
  return matches;
}

function detectClientName(fileName) {
  const withoutExtension = fileName.replace(/\.[^.]+$/, "");
  return withoutExtension
    .replace(/(^|[-_.\s])20\d{2}([-_.\s]|$)/g, " ")
    .replace(/TECH[-\s]?\d+/gi, " ")
    .replace(/[-_.]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function detectCurrency(text) {
  const normalized = normalizeText(text);
  if (/\busd\b|dolar|dolares/.test(normalized)) {
    return "USD";
  }
  if (/\bgs\b|guarani|guaranies|pyg/.test(normalized)) {
    return "GS";
  }
  return "";
}

function cleanCurrency(value) {
  const text = normalizeText(value).toUpperCase();
  if (!text) {
    return "";
  }
  if (text.includes("USD")) {
    return "USD";
  }
  if (text.includes("GS") || text.includes("PYG")) {
    return "GS";
  }
  return stringify(value).trim();
}

function cleanProvider(value) {
  const provider = stringify(value).trim();
  if (!provider || /^0+([.,]0+)?$/.test(provider)) {
    return "";
  }
  return provider.replace(/\s+/g, " ");
}

function normalizeProvider(value) {
  return normalizeText(value || "PROVEEDOR PENDIENTE");
}

function normalizeKey(value) {
  return normalizeText(value).replace(/[^a-z0-9]/g, "");
}

function normalizeText(value) {
  return stringify(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function stringify(value) {
  if (value === null || value === undefined) {
    return "";
  }
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  return String(value);
}

function valueAt(row, index) {
  return index === undefined ? null : row[index];
}

function toNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : NaN;
  }
  const text = stringify(value).trim();
  if (!text) {
    return NaN;
  }

  const cleaned = text
    .replace(/\s/g, "")
    .replace(/[^\d,.-]/g, "");

  if (!cleaned || cleaned === "-" || cleaned === ",") {
    return NaN;
  }

  const decimalComma = cleaned.includes(",") && (!cleaned.includes(".") || cleaned.lastIndexOf(",") > cleaned.lastIndexOf("."));
  const normalized = decimalComma
    ? cleaned.replace(/\./g, "").replace(",", ".")
    : cleaned.replace(/,/g, "");

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : NaN;
}

function roundMoney(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function formatMoney(value) {
  return new Intl.NumberFormat("es-PY", {
    maximumFractionDigits: 2,
  }).format(Number(value) || 0);
}

function formatBytes(bytes) {
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  if (bytes < 1024 * 1024) {
    return `${(bytes / 1024).toFixed(1)} KB`;
  }
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function dateStamp() {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  return `${yyyy}${mm}${dd}`;
}

render();
