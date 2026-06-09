const DB_NAME = "beihuo-demand-allocation-library";
const DB_VERSION = 1;
const STORE_NAME = "file-slots";

const slots = [
  { id: "file-1", label: "Dim-YL医疗器械商品分类" },
  { id: "file-2", label: "Dim-采购部分工明细" },
  { id: "file-3", label: "Dim-花名册（姓名&一级部门）" },
  { id: "file-4", label: "备货需求表" },
];

const TABLE_FILE_ACCEPT = ".xlsx,.xls,.xlsm,.csv";
const FILTER_DEFINITIONS = [
  { key: "businessUnit", label: "事业部" },
  { key: "productLine", label: "销售产品线" },
  { key: "purchaseGroup", label: "采购组" },
  { key: "buyer", label: "采购订单下单人" },
  { key: "supplierShortName", label: "供应商简称" },
];
const DETAIL_COLUMNS = [
  { key: "purchaseGroup", label: "采购组" },
  { key: "buyer", label: "采购单订单下单人" },
  { key: "applicant", label: "申请人" },
  { key: "supplierShortName", label: "供应商简称" },
  { key: "oaProcessNo", label: "OA备货流程号" },
  { key: "materialCode", label: "物料编码" },
  { key: "sku", label: "SKU" },
  { key: "materialName", label: "物料名称" },
  { key: "quantity", label: "数量" },
];
const FIELD_ALIASES = {
  businessUnit: ["事业部", "业务单元", "业务部门", "部门"],
  productLine: ["销售产品线", "产品线", "一级产品线", "销售线"],
  purchaseGroup: ["采购组", "采购分组", "采购组别", "采购分组名称"],
  buyer: ["采购单订单下单人", "采购订单下单人", "采购单下单人", "订单下单人", "下单人"],
  applicant: ["申请人", "OA申请人", "流程申请人"],
  supplierShortName: ["供应商简称", "供应商", "供应商名称", "供应商短名称"],
  oaProcessNo: ["OA备货流程号", "OA流程号", "备货流程号", "流程号", "OA编号"],
  materialCode: ["物料编码", "物料编号", "商品编码", "存货编码", "产品编码"],
  sku: ["SKU", "sku", "Sku"],
  materialName: ["物料名称", "物料名", "商品名称", "品名", "产品名称"],
  quantity: ["数量", "备货数量", "申请数量", "下单数量", "下单数量-备货需求OA申请为准"],
};

const els = {
  slotGrid: document.querySelector("#slotGrid"),
  applyAllButton: document.querySelector("#applyAllButton"),
  clearCacheButton: document.querySelector("#clearCacheButton"),
  libraryState: document.querySelector("#libraryState"),
  sourceNote: document.querySelector("#librarySourceNote"),
  slotCount: document.querySelector("#slotCount"),
  uploadedCount: document.querySelector("#uploadedCount"),
  appliedCount: document.querySelector("#appliedCount"),
  beihuoTotal: document.querySelector("#beihuoTotal"),
  supplierCount: document.querySelector("#supplierCount"),
  filterToolbar: document.querySelector("#filterToolbar"),
  clearFiltersButton: document.querySelector("#clearFiltersButton"),
  downloadDetailButton: document.querySelector("#downloadDetailButton"),
  detailState: document.querySelector("#detailState"),
  detailTableBody: document.querySelector("#detailTableBody"),
  detailEmpty: document.querySelector("#detailEmpty"),
};

const state = {
  records: new Map(),
  demandRows: [],
  filteredRows: [],
  filterOptions: Object.fromEntries(FILTER_DEFINITIONS.map((filter) => [filter.key, []])),
  filters: new Map(FILTER_DEFINITIONS.map((filter) => [filter.key, new Set()])),
};

async function init() {
  bindEvents();
  renderFilterControls();
  if (window.ensureSharedLibraryLoaded) {
    await window.ensureSharedLibraryLoaded();
  }
  await refresh();
}

function bindEvents() {
  els.slotGrid?.addEventListener("change", async (event) => {
    const input = event.target.closest("[data-upload]");
    if (!input) return;
    await savePendingFile(input.dataset.upload, input.files?.[0]);
    input.value = "";
  });

  els.slotGrid?.addEventListener("click", async (event) => {
    const applyButton = event.target.closest("[data-apply]");
    if (applyButton) {
      await applySlot(applyButton.dataset.apply);
      return;
    }

    const deleteButton = event.target.closest("[data-delete]");
    if (deleteButton) {
      await deleteSlot(deleteButton.dataset.delete);
    }
  });

  els.slotGrid?.addEventListener("dragover", (event) => {
    const card = event.target.closest("[data-drop]");
    if (!card) return;
    event.preventDefault();
    card.classList.add("drag-over");
  });

  els.slotGrid?.addEventListener("dragleave", (event) => {
    const card = event.target.closest("[data-drop]");
    if (!card || card.contains(event.relatedTarget)) return;
    card.classList.remove("drag-over");
  });

  els.slotGrid?.addEventListener("drop", async (event) => {
    const card = event.target.closest("[data-drop]");
    if (!card) return;
    event.preventDefault();
    card.classList.remove("drag-over");
    await savePendingFile(card.dataset.drop, event.dataTransfer?.files?.[0]);
  });

  els.applyAllButton?.addEventListener("click", applyAllSlots);
  els.clearCacheButton?.addEventListener("click", clearLibraryCache);
  els.filterToolbar?.addEventListener("click", handleFilterToolbarClick);
  els.filterToolbar?.addEventListener("change", handleFilterToolbarChange);
  els.clearFiltersButton?.addEventListener("click", clearAllFilters);
  els.downloadDetailButton?.addEventListener("click", downloadDetailWorkbook);
  document.addEventListener("click", (event) => {
    if (!event.target.closest(".multi-filter")) closeFilterMenus();
  });
}

async function refresh() {
  const db = await openDb();
  const entries = await Promise.all(slots.map(async (slot) => [slot.id, await getRecord(db, slot.id)]));
  db.close();
  state.records = new Map(entries);
  renderShell();
  await refreshDemandDashboard();
}

async function savePendingFile(slotId, file) {
  if (!file) return;
  const invalidMessage = getInvalidFileMessage(file);
  if (invalidMessage) {
    setLibraryState(invalidMessage);
    window.alert(invalidMessage);
    return;
  }

  const now = new Date().toISOString();
  const existing = state.records.get(slotId) || { id: slotId };
  const record = {
    ...existing,
    id: slotId,
    pendingFile: file,
    pendingName: file.name,
    pendingSize: file.size,
    pendingTypeLabel: getFileTypeLabel(file),
    pendingRefreshMonth: getRefreshMonth(file.name, now),
    pendingSavedAt: now,
  };

  const db = await openDb();
  await putRecord(db, record);
  db.close();
  await refresh();
}

function getInvalidFileMessage(file) {
  return isSupportedTableFile(file?.name) ? "" : "请上传 .xlsx、.xls、.xlsm 或 .csv 表格文件。";
}

async function applySlot(slotId, options = {}) {
  const record = state.records.get(slotId);
  if (!record) return false;

  const targetName = record.pendingName || record.name;
  const invalidMessage = getInvalidFileMessage({ name: targetName });
  if (invalidMessage) {
    setLibraryState(invalidMessage);
    window.alert(invalidMessage);
    return false;
  }

  const appliedAt = new Date().toISOString();
  const updatedRecord = record.pendingFile
    ? clearPendingFields({
        ...record,
        file: record.pendingFile,
        name: record.pendingName,
        size: record.pendingSize,
        typeLabel: record.pendingTypeLabel,
        refreshMonth: record.pendingRefreshMonth,
        savedAt: record.pendingSavedAt,
        applied: true,
        appliedAt,
      })
    : {
        ...record,
        applied: true,
        appliedAt,
      };

  const db = await openDb();
  await putRecord(db, updatedRecord);
  db.close();
  if (!options.skipRefresh) await refresh();
  return true;
}

async function applyAllSlots() {
  const targetSlotIds = slots
    .filter((slot) => {
      const record = state.records.get(slot.id);
      return record?.pendingFile || (record && !record.applied);
    })
    .map((slot) => slot.id);

  if (els.applyAllButton) els.applyAllButton.disabled = true;
  setLibraryState("刷新应用中");
  try {
    for (const slotId of targetSlotIds) {
      const applied = await applySlot(slotId, { skipRefresh: true });
      if (applied === false) throw new Error("应用刷新失败");
    }
    await refresh();
    setLibraryState(targetSlotIds.length ? "已刷新应用" : "暂无待应用文件");
  } catch (error) {
    console.warn("apply all failed", error);
    setLibraryState(error.message || "刷新应用失败");
  } finally {
    updateApplyAllButton();
  }
}

async function deleteSlot(slotId) {
  const db = await openDb();
  await deleteRecord(db, slotId);
  db.close();
  await refresh();
}

async function clearLibraryCache() {
  const confirmed = window.confirm("确认清除当前浏览器里的备货需求分配文件缓存吗？清除后需要重新上传并应用文件。");
  if (!confirmed) return;

  if (els.clearCacheButton) els.clearCacheButton.disabled = true;
  setLibraryState("清除中");
  try {
    await deleteLibraryDatabase();
    state.records = new Map();
    state.demandRows = [];
    state.filteredRows = [];
    resetAllFilterState();
    renderShell();
    setLibraryState("已清除缓存");
  } catch (error) {
    console.warn("clear library cache failed", error);
    setLibraryState("清除失败");
    window.alert(error.message || "清除缓存失败，请关闭其他备货需求分配页面后重试。");
  } finally {
    if (els.clearCacheButton) els.clearCacheButton.disabled = false;
  }
}

async function refreshDemandDashboard() {
  if (!window.XLSX) {
    state.demandRows = [];
    state.filteredRows = [];
    renderDemandDashboard("Excel 解析组件未加载，请刷新页面后重试");
    return;
  }

  const records = slots
    .map((slot) => state.records.get(slot.id))
    .filter((record) => record?.applied && record?.file && isSupportedTableFile(record.name));

  if (!records.length) {
    state.demandRows = [];
    state.filteredRows = [];
    resetAllFilterState();
    renderDemandDashboard("上传并应用表格后显示明细");
    return;
  }

  const results = await Promise.all(
    records.map(async (record) => {
      try {
        return await readDemandRowsFromRecord(record);
      } catch (error) {
        console.warn("read demand file failed", record?.name, error);
        return [];
      }
    })
  );

  state.demandRows = results.flat();
  rebuildFilterOptions();
  pruneFilterSelections();
  renderDemandDashboard();
}

async function readDemandRowsFromRecord(record) {
  const workbook = await readWorkbook(record.file);
  return (workbook.SheetNames || []).flatMap((sheetName) =>
    extractDemandRowsFromSheet(workbook.Sheets[sheetName], {
      fileName: record.name,
      slotId: record.id,
      sheetName,
    })
  );
}

async function readWorkbook(file) {
  const extension = String(file?.name || "").split(".").pop()?.toLowerCase();
  if (extension === "csv") {
    return window.XLSX.read(await file.text(), {
      type: "string",
      raw: false,
      cellText: true,
    });
  }
  return window.XLSX.read(await file.arrayBuffer(), {
    type: "array",
    cellText: true,
    cellStyles: true,
  });
}

function extractDemandRowsFromSheet(sheet, source) {
  if (!sheet) return [];
  const rows = window.XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    blankrows: false,
    raw: false,
  });
  const headerIndex = findDemandHeaderRowIndex(rows);
  if (headerIndex < 0) return [];

  const columnMap = buildDemandColumnMap(rows[headerIndex]);
  return rows.slice(headerIndex + 1).map((row, rowOffset) => {
    const item = Object.fromEntries(Object.keys(FIELD_ALIASES).map((key) => [key, getDemandCellValue(row, columnMap[key])]));
    const hasDetailValue = DETAIL_COLUMNS.some((column) => item[column.key]);
    const hasFilterValue = FILTER_DEFINITIONS.some((filter) => item[filter.key]);
    if (!hasDetailValue && !hasFilterValue) return null;
    item.quantityNumber = parseDemandQuantity(item.quantity);
    item.sourceFile = source.fileName;
    item.sourceSheet = source.sheetName;
    item.sourceSlotId = source.slotId;
    item.sourceRowNumber = headerIndex + rowOffset + 2;
    return item;
  }).filter(Boolean);
}

function findDemandHeaderRowIndex(rows) {
  let bestIndex = -1;
  let bestScore = 0;
  rows.slice(0, 30).forEach((row, index) => {
    const score = Object.keys(buildDemandColumnMap(row)).length;
    if (score > bestScore) {
      bestIndex = index;
      bestScore = score;
    }
  });
  return bestScore >= 2 ? bestIndex : -1;
}

function buildDemandColumnMap(headerRow = []) {
  const normalizedHeaders = headerRow.map(normalizeDemandHeader);
  const columnMap = {};
  Object.entries(FIELD_ALIASES).forEach(([key, aliases]) => {
    const normalizedAliases = aliases.map(normalizeDemandHeader).filter(Boolean);
    const exactIndex = normalizedHeaders.findIndex((header) => normalizedAliases.includes(header));
    if (exactIndex >= 0) {
      columnMap[key] = exactIndex;
      return;
    }

    const looseIndex = normalizedHeaders.findIndex((header) =>
      header && normalizedAliases.some((alias) => alias && (header.includes(alias) || alias.includes(header)))
    );
    if (looseIndex >= 0) columnMap[key] = looseIndex;
  });
  return columnMap;
}

function normalizeDemandHeader(value) {
  return String(value ?? "")
    .trim()
    .replace(/\s+/g, "")
    .replace(/[()[\]{}（）【】<>《》:：]/g, "")
    .toLowerCase();
}

function getDemandCellValue(row, columnIndex) {
  if (columnIndex === undefined || columnIndex < 0) return "";
  return String(row?.[columnIndex] ?? "").trim();
}

function parseDemandQuantity(value) {
  const match = String(value ?? "").replace(/,/g, "").match(/-?\d+(?:\.\d+)?/);
  if (!match) return NaN;
  const number = Number(match[0]);
  return Number.isFinite(number) ? number : NaN;
}

function isSupportedTableFile(fileName) {
  return /\.(xlsx|xlsm|xls|csv)$/i.test(String(fileName || ""));
}

function rebuildFilterOptions() {
  state.filterOptions = Object.fromEntries(
    FILTER_DEFINITIONS.map((filter) => [
      filter.key,
      uniqueSortedValues(state.demandRows.map((row) => row[filter.key]).filter(Boolean)),
    ])
  );
}

function uniqueSortedValues(values) {
  return [...new Set(values.map((value) => String(value ?? "").trim()).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, "zh-CN", { numeric: true }));
}

function pruneFilterSelections() {
  FILTER_DEFINITIONS.forEach((filter) => {
    const options = new Set(state.filterOptions[filter.key] || []);
    const selected = state.filters.get(filter.key) || new Set();
    state.filters.set(filter.key, new Set([...selected].filter((value) => options.has(value))));
  });
}

function resetAllFilterState() {
  FILTER_DEFINITIONS.forEach((filter) => state.filters.set(filter.key, new Set()));
  state.filterOptions = Object.fromEntries(FILTER_DEFINITIONS.map((filter) => [filter.key, []]));
}

function renderShell() {
  if (els.slotCount) els.slotCount.textContent = String(slots.length);
  if (els.uploadedCount) els.uploadedCount.textContent = String(countUploadedRecords());
  if (els.appliedCount) els.appliedCount.textContent = String(countAppliedRecords());
  if (els.sourceNote) els.sourceNote.textContent = `本地文件库｜引用时间：${getLatestAppliedTime()}`;
  if (els.slotGrid) els.slotGrid.innerHTML = slots.map(renderSlot).join("");
  renderDemandDashboard();
  setLibraryState("本地文件库");
  updateApplyAllButton();
}

function renderSlot(slot) {
  const record = state.records.get(slot.id);
  const display = getDisplayRecord(record);
  const hasPending = Boolean(record?.pendingFile);
  const isApplied = Boolean(record?.applied && !hasPending);
  return `
    <article class="slot-card ${isApplied ? "applied" : ""}" data-drop="${slot.id}">
      <div class="slot-head">
        <div>
          <p class="eyebrow">Slot</p>
          <h2>${escapeHtml(slot.label)}</h2>
        </div>
        <span class="slot-status ${isApplied ? "applied" : hasPending ? "pending" : ""}">
          ${isApplied ? "已引用" : hasPending ? "待应用" : "未上传"}
        </span>
      </div>
      <div class="file-meta">
        <span>${escapeHtml(display?.name || "未上传文件")}</span>
        <strong>${display ? `${escapeHtml(display.typeLabel)} / ${formatFileSize(display.size)}` : "--"}</strong>
        <small>刷新月份：${escapeHtml(display?.refreshMonth || "--")}</small>
        <small>更新时间：${formatDateTime(display?.savedAt)}</small>
        <small>引用时间：${formatDateTime(record?.appliedAt)}</small>
      </div>
      <div class="slot-actions">
        <label class="upload-button">
          <input type="file" accept="${TABLE_FILE_ACCEPT}" data-upload="${slot.id}" />
          上传/替换
        </label>
        <button type="button" data-apply="${slot.id}" ${hasPending || (record && !record.applied) ? "" : "disabled"}>应用刷新</button>
        <button class="danger-button" type="button" data-delete="${slot.id}" ${record ? "" : "disabled"}>删除</button>
      </div>
    </article>
  `;
}

function renderDemandDashboard(emptyMessage = "暂无匹配数据") {
  state.filteredRows = getFilteredDemandRows();
  updateDemandMetrics();
  renderFilterControls();
  renderDemandTable(emptyMessage);
}

function getFilteredDemandRows() {
  return state.demandRows.filter((row) =>
    FILTER_DEFINITIONS.every((filter) => {
      const selected = state.filters.get(filter.key);
      return !selected?.size || selected.has(row[filter.key]);
    })
  );
}

function updateDemandMetrics() {
  const quantityValues = state.filteredRows.map((row) => row.quantityNumber).filter(Number.isFinite);
  const total = quantityValues.length
    ? quantityValues.reduce((sum, value) => sum + value, 0)
    : state.filteredRows.length;
  const supplierCount = new Set(state.filteredRows.map((row) => row.supplierShortName).filter(Boolean)).size;
  if (els.beihuoTotal) els.beihuoTotal.textContent = formatMetricNumber(total);
  if (els.supplierCount) els.supplierCount.textContent = formatMetricNumber(supplierCount);
  if (els.downloadDetailButton) els.downloadDetailButton.disabled = !state.filteredRows.length || !window.XLSX;
}

function renderFilterControls() {
  FILTER_DEFINITIONS.forEach((filter) => {
    const container = els.filterToolbar?.querySelector(`[data-filter="${filter.key}"]`);
    if (!container) return;
    const options = state.filterOptions[filter.key] || [];
    const selected = state.filters.get(filter.key) || new Set();
    container.innerHTML = `
      <button class="multi-filter-button" type="button" data-filter-button="${filter.key}" ${options.length ? "" : "disabled"}>
        <span>${escapeHtml(getFilterButtonLabel(filter, selected))}</span>
        <i aria-hidden="true">▾</i>
      </button>
      <div class="multi-filter-menu">
        <label class="multi-filter-option">
          <input type="checkbox" data-filter-option data-filter-key="${filter.key}" value="__all__" ${selected.size ? "" : "checked"} />
          <span>全部</span>
        </label>
        ${
          options.length
            ? options.map((value) => renderFilterOption(filter.key, value, selected.has(value))).join("")
            : `<div class="multi-filter-empty">暂无选项</div>`
        }
      </div>
    `;
  });
}

function renderFilterOption(filterKey, value, checked) {
  const safeValue = escapeHtml(value);
  return `
    <label class="multi-filter-option ${checked ? "selected" : ""}">
      <input type="checkbox" data-filter-option data-filter-key="${filterKey}" value="${safeValue}" ${checked ? "checked" : ""} />
      <span>${safeValue}</span>
    </label>
  `;
}

function getFilterButtonLabel(filter, selected) {
  if (!selected?.size) return `全部${filter.label}`;
  const values = [...selected];
  if (values.length === 1) return values[0];
  if (values.length === 2) return values.join("、");
  return `已选${values.length}项`;
}

function handleFilterToolbarClick(event) {
  const button = event.target.closest("[data-filter-button]");
  if (!button) return;
  event.stopPropagation();
  const container = button.closest(".multi-filter");
  const shouldOpen = !container.classList.contains("open");
  closeFilterMenus();
  if (shouldOpen) container.classList.add("open");
}

function handleFilterToolbarChange(event) {
  const input = event.target.closest("[data-filter-option]");
  if (!input) return;
  const selected = new Set(state.filters.get(input.dataset.filterKey) || []);
  if (input.value === "__all__") {
    selected.clear();
  } else if (input.checked) {
    selected.add(input.value);
  } else {
    selected.delete(input.value);
  }
  state.filters.set(input.dataset.filterKey, selected);
  renderDemandDashboard();
}

function clearAllFilters() {
  FILTER_DEFINITIONS.forEach((filter) => state.filters.set(filter.key, new Set()));
  closeFilterMenus();
  renderDemandDashboard();
}

function closeFilterMenus() {
  els.filterToolbar?.querySelectorAll(".multi-filter.open").forEach((container) => container.classList.remove("open"));
}

function renderDemandTable(emptyMessage) {
  if (!els.detailTableBody || !els.detailEmpty) return;
  const rows = state.filteredRows;
  els.detailTableBody.innerHTML = rows.map(renderDemandTableRow).join("");
  els.detailEmpty.hidden = rows.length > 0;
  els.detailEmpty.textContent = state.demandRows.length ? emptyMessage : "上传并应用表格后显示明细";
  if (els.detailState) {
    els.detailState.textContent = rows.length
      ? `显示 ${formatMetricNumber(rows.length)} 条`
      : state.demandRows.length ? "无匹配数据" : "暂无数据";
  }
}

function renderDemandTableRow(row) {
  return `
    <tr>
      ${DETAIL_COLUMNS.map((column) => `<td>${escapeHtml(row[column.key] || "")}</td>`).join("")}
    </tr>
  `;
}

function downloadDetailWorkbook() {
  if (!state.filteredRows.length) {
    window.alert("暂无可下载的明细数据。");
    return;
  }
  if (!window.XLSX) {
    window.alert("Excel 导出组件未加载，请刷新页面后重试。");
    return;
  }

  const workbook = window.XLSX.utils.book_new();
  const usedSheetNames = new Set();
  groupRowsByBuyer(state.filteredRows).forEach(([buyer, rows]) => {
    const sheetName = makeUniqueSheetName(buyer, usedSheetNames);
    const sheetRows = [
      DETAIL_COLUMNS.map((column) => column.label),
      ...rows.map((row) => DETAIL_COLUMNS.map((column) => row[column.key] || "")),
    ];
    const worksheet = window.XLSX.utils.aoa_to_sheet(sheetRows);
    worksheet["!cols"] = DETAIL_COLUMNS.map((column) => ({ wch: getExportColumnWidth(column.key) }));
    window.XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  });

  window.XLSX.writeFile(workbook, `备货需求分配明细_${formatFileTimestamp(new Date())}.xlsx`);
}

function groupRowsByBuyer(rows) {
  const groups = new Map();
  rows.forEach((row) => {
    const buyer = String(row.buyer || "").trim() || "未填写下单人";
    if (!groups.has(buyer)) groups.set(buyer, []);
    groups.get(buyer).push(row);
  });
  return [...groups.entries()].sort(([left], [right]) => left.localeCompare(right, "zh-CN", { numeric: true }));
}

function makeUniqueSheetName(rawName, usedSheetNames) {
  const baseName = sanitizeSheetName(rawName) || "未填写下单人";
  let sheetName = baseName.slice(0, 31);
  let index = 2;
  while (usedSheetNames.has(sheetName)) {
    const suffix = `_${index}`;
    sheetName = `${baseName.slice(0, 31 - suffix.length)}${suffix}`;
    index += 1;
  }
  usedSheetNames.add(sheetName);
  return sheetName;
}

function sanitizeSheetName(value) {
  return String(value || "")
    .replace(/[\[\]:*?/\\]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function getExportColumnWidth(key) {
  return {
    purchaseGroup: 14,
    buyer: 18,
    applicant: 14,
    supplierShortName: 20,
    oaProcessNo: 20,
    materialCode: 18,
    sku: 16,
    materialName: 28,
    quantity: 12,
  }[key] || 14;
}

function getDisplayRecord(record) {
  if (!record) return null;
  if (record.pendingFile) {
    return {
      name: record.pendingName,
      size: record.pendingSize,
      typeLabel: record.pendingTypeLabel,
      refreshMonth: record.pendingRefreshMonth,
      savedAt: record.pendingSavedAt,
    };
  }
  return record;
}

function countUploadedRecords() {
  return slots.filter((slot) => {
    const record = state.records.get(slot.id);
    return record?.file || record?.pendingFile;
  }).length;
}

function countAppliedRecords() {
  return slots.filter((slot) => state.records.get(slot.id)?.applied).length;
}

function updateApplyAllButton() {
  const hasAnyRecordToApply = slots.some((slot) => {
    const record = state.records.get(slot.id);
    return record?.pendingFile || (record && !record.applied);
  });
  if (els.applyAllButton) els.applyAllButton.disabled = !hasAnyRecordToApply;
}

function setLibraryState(text) {
  if (els.libraryState) els.libraryState.textContent = text;
}

function clearPendingFields(record) {
  const nextRecord = { ...record };
  delete nextRecord.pendingFile;
  delete nextRecord.pendingName;
  delete nextRecord.pendingSize;
  delete nextRecord.pendingTypeLabel;
  delete nextRecord.pendingRefreshMonth;
  delete nextRecord.pendingSavedAt;
  return nextRecord;
}

function getRefreshMonth(fileName, fallbackTime) {
  const name = String(fileName || "");
  const compact = name.match(/(20\d{2})(0[1-9]|1[0-2])/);
  if (compact) return `${compact[1]}-${compact[2]}`;
  const separated = name.match(/(20\d{2})[-_.年]+(0?[1-9]|1[0-2])/);
  if (separated) return `${separated[1]}-${String(separated[2]).padStart(2, "0")}`;
  return getMonthFromDate(fallbackTime);
}

function getMonthFromDate(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "--";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function getLatestAppliedTime() {
  const times = slots
    .map((slot) => state.records.get(slot.id)?.appliedAt)
    .filter(Boolean)
    .sort();
  if (!times.length) return "--";
  return formatDateTime(times.at(-1));
}

function getFileTypeLabel(file) {
  const extension = String(file?.name || "").split(".").pop()?.toLowerCase();
  if (extension === "xlsx" || extension === "xls" || extension === "xlsm") return "Excel 工作簿";
  if (extension === "csv") return "CSV 文件";
  return file?.type || "表格文件";
}

function formatDateTime(value) {
  if (!value) return "--";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "--";
  return new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function formatFileSize(size = 0) {
  if (!Number.isFinite(size) || size <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  let value = size;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  return `${value.toFixed(value >= 10 || unitIndex === 0 ? 0 : 1)} ${units[unitIndex]}`;
}

function formatFileTimestamp(date) {
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
    String(date.getHours()).padStart(2, "0"),
    String(date.getMinutes()).padStart(2, "0"),
  ].join("");
}

function formatMetricNumber(value) {
  return new Intl.NumberFormat("zh-CN", {
    maximumFractionDigits: Number.isInteger(value) ? 0 : 2,
  }).format(value || 0);
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  })[char]);
}

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function getRecord(db, key) {
  return runStoreRequest(db, "readonly", (store) => store.get(key));
}

function putRecord(db, record) {
  return runStoreRequest(db, "readwrite", (store) => store.put(record));
}

function deleteRecord(db, key) {
  return runStoreRequest(db, "readwrite", (store) => store.delete(key));
}

function runStoreRequest(db, mode, createRequest) {
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(STORE_NAME, mode);
    const request = createRequest(transaction.objectStore(STORE_NAME));
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    transaction.onerror = () => reject(transaction.error);
  });
}

function deleteLibraryDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.deleteDatabase(DB_NAME);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
    request.onblocked = () => reject(new Error("文件库正被其他备货需求分配页面占用，请关闭其他备货需求分配页面后重试。"));
  });
}

init().catch((error) => {
  console.error(error);
  setLibraryState("读取失败");
});
