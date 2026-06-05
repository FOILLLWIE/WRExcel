const qs = (selector, root = document) => root.querySelector(selector);
const qsa = (selector, root = document) => [...root.querySelectorAll(selector)];
const byType = (type) => qs(`[data-type="${type}"]`);
const columnField = (type) => qs(`.column-field[data-type="${type}"]`);
const columnInput = (type) => qs(`[data-role="column-input"][data-type="${type}"]`);
const optionInput = (name) => qs(`[data-role="option"][data-option="${name}"]`);
const selectInput = (target) => qs(`[data-role="native-select"][data-target="${target}"]`);
const actionButton = (action) => qs(`[data-role="action"][data-action="${action}"]`);

const fileInput = qs('[data-role="file-input"]');
const fileDrop = qs(".file-drop");
const statusCard = qs('[data-role="status"]');
const mappingSection = qs('[data-section="mapping"]');
const resultSection = qs('[data-section="results"]');
const sheetSelect = selectInput("sheet");
const previewTable = qs('[data-role="preview-table"]');
const productColumn = columnInput("product");
const productColorChoices = qs('[data-role="color-choices"]', columnField("product"));
const productColorPicker = qs('[data-role="color-picker"]', columnField("product"));
const productColorTrigger = qs('[data-role="color-trigger"]', columnField("product"));
const originalColumn = columnInput("original");
const originalColorChoices = qs('[data-role="color-choices"]', columnField("original"));
const originalColorPicker = qs('[data-role="color-picker"]', columnField("original"));
const originalColorTrigger = qs('[data-role="color-trigger"]', columnField("original"));
const finalPriceColumn = columnInput("finalPrice");
const categoryColumn = columnInput("category");
const finalColorChoices = qs('[data-role="color-choices"]', columnField("finalPrice"));
const finalColorPicker = qs('[data-role="color-picker"]', columnField("finalPrice"));
const finalColorTrigger = qs('[data-role="color-trigger"]', columnField("finalPrice"));
const startRow = qs('[data-role="row-input"][data-type="startRow"]');
const endRow = qs('[data-role="row-input"][data-type="endRow"]');
const currentPickLabel = qs('[data-role="current-pick-label"]');
const calculateButton = actionButton("calculate");
const itemCount = qs('[data-role="item-count"]');
const resultBody = qs('[data-role="result-body"]');
const editResultsButton = actionButton("toggle-edit-results");
const cancelEditResultsButton = actionButton("cancel-edit-results");
const addResultRowButton = actionButton("add-result-row");
const editAddRowWrap = qs('[data-role="edit-add-row-wrap"]');
const expandResultsButton = actionButton("toggle-expand-results");
const productSearch = qs('[data-role="product-search"]');
const productFilterButton = actionButton("toggle-product-filter");
const productFilterPanel = qs('[data-role="product-filter-panel"]');
const productFilterList = qs('[data-role="product-filter-list"]');
const addProductFilterButton = actionButton("add-product-filter");
const removeParenthesesText = optionInput("removeParenthesesText");
const removeBracketsText = optionInput("removeBracketsText");
const removeAfterDelimiter = optionInput("removeAfterDelimiter");
const removeBeforeDelimiter = optionInput("removeBeforeDelimiter");
const removeLeadingText = optionInput("removeLeadingText");
const leadingTextOptions = qs('[data-role="option-panel"][data-option-panel="removeLeadingText"]');
const leadingTextValue = qs('[data-role="option-value"][data-option="removeLeadingText"]');
const delimiterOptions = qs('[data-role="option-panel"][data-option-panel="removeAfterDelimiter"]');
const delimiterValue = qs('[data-role="option-value"][data-option="removeAfterDelimiter"]');
const visibleColumnsButton = actionButton("toggle-visible-columns");
const visibleColumnsPanel = qs('[data-role="visible-columns-panel"]');
const showWonSuffix = optionInput("showWonSuffix");
const showDiscountMinus = optionInput("showDiscountMinus");
const highlightFinalPrices = optionInput("highlightFinalPrices");
const titleCaseProductName = optionInput("titleCaseProductName");
const autoLoadColorFilters = optionInput("autoLoadColorFilters");
const groupSimilarProducts = optionInput("groupSimilarProducts");
const colorLoadOptionText = autoLoadColorFilters?.closest("label")?.querySelector(".option-label");
const sortField = selectInput("sortField");
const sortDirection = selectInput("sortDirection");
const emptyResultMessage = qs('[data-role="empty-result-message"]');
const productNameHeader = qs('[data-column="product_name"]');
const productNameResizer = qs('[data-role="product-name-resizer"]');
const appDialog = qs('[data-role="dialog"]');
const dialogTitle = qs('[data-role="dialog-title"]', appDialog);
const dialogMessage = qs('[data-role="dialog-message"]', appDialog);
const dialogActions = qs('[data-role="dialog-actions"]', appDialog);
const addExtraFieldButton = actionButton("add-extra-field");
const extraFieldsList = qs('[data-role="dynamic-fields-list"][data-type="extra"]');
const resultHeaderRow = qs('[data-role="result-header-row"]');
const tableWrap = qs(".table-wrap");
const extraColumnsAnchor = qs('[data-column="total_reward_amount"]', resultHeaderRow);
const addRewardFieldButton = actionButton("add-reward-field");
const rewardFieldsList = qs('[data-role="dynamic-fields-list"][data-type="reward"]');
const rewardColumnsAnchor = qs('[data-column="total_reward_amount"]', resultHeaderRow);

const currencyFormatter = new Intl.NumberFormat("ko-KR");
let uploadedFile = null;
let excelWorkbook = null;
let workbookPreview = null;
let activePick = "product";
let pendingPayload = null;
let currentRows = [];
let isEditMode = false;
let selectedRowId = null;
const hiddenColumns = new Set();
let productNameFilters = [];
let extraFields = [];
let rewardFields = [];
const collapsedCategories = new Set();
let activeExtraFieldTarget = null;
let pendingRestoreConfig = null;
let activeFileSettingsConfig = null;
let activeResultSnapshotConfig = null;
let lastResultMapping = null;
let resultSnapshotDirty = false;
let currentFileKey = "";
let currentFileContentHash = "";
let settingsSaveTimer = null;
let floatingHeader = null;
let floatingHeaderTable = null;
let floatingCategory = null;
let activeInlineEdit = null;
let editRowsSnapshot = null;
let editSortSnapshot = null;
let draggedResultRowId = null;
let lastActiveCategoryKey = "";
const DB_NAME = "discountCalculatorDB";
const DB_VERSION = 1;
const SETTINGS_STORE_NAME = "fileSettings";
const SETTINGS_RETENTION_DAYS = 30;
const LOCAL_SETTINGS_KEY = "discountCalculatorFileSettingsV1";
const LAST_SETTINGS_KEY = "WRExcelLastFileSettings";
const fallbackSettingsStore = new Map();
const extraFieldNameSuggestions = ["선택쿠폰", "중복쿠폰", "카드할인"];
const GROUPED_PRODUCT_SEPARATOR = "\uFF5C";
const GROUPED_PRODUCT_JOINER = ` ${GROUPED_PRODUCT_SEPARATOR} `;
const rewardFieldNameSuggestions = ["\uC2A4\uB9C8\uC77C\uCE74\uB4DC", "\uBA38\uB2C8\uCDA9\uC804", "\uAF2D\uBA64\uBC841", "\uAF2D\uBA64\uBC842"];
const worksheetMergeRangeCache = new WeakMap();
const worksheetRowBoundsCache = new WeakMap();
const worksheetConditionalFillCache = new WeakMap();
const worksheetStatusFillCache = new WeakMap();
const CONDITIONAL_STATUS_RED = "FF0000";
const columnFilters = {
  product: { input: productColumn, picker: productColorPicker, trigger: productColorTrigger, choices: productColorChoices, selected: "", mode: "include" },
  original: { input: originalColumn, picker: originalColorPicker, trigger: originalColorTrigger, choices: originalColorChoices, selected: "", mode: "include" },
  final: { input: finalPriceColumn, picker: finalColorPicker, trigger: finalColorTrigger, choices: finalColorChoices, selected: "", mode: "include" },
};

if (window.location.protocol === "file:") {
  setStatus("이 페이지는 파일로 직접 열려 있습니다. start.bat을 실행해 접속해주세요.", "error");
}
let productColumnWidth = null;
let productColumnDragWidth = null;
let productColumnUserResized = false;
let appliedProductColumnWidth = null;
let appliedResultColumnWidths = [];

initializeCustomSelects();
cleanupStoredSettings();
window.addEventListener("beforeunload", () => {
  if (uploadedFile && currentFileKey) saveCurrentSettingsToLocalMirror();
});
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden" && uploadedFile && currentFileKey) saveCurrentSettingsToLocalMirror();
});
window.addEventListener("scroll", updateFloatingResultHeader, { passive: true });
window.addEventListener("resize", updateFloatingResultHeader);
tableWrap?.addEventListener("scroll", updateFloatingResultHeader, { passive: true });

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  uploadedFile = file;
  currentFileKey = createFileKey(file);
  currentFileContentHash = await createFileContentHash(file);
  activeFileSettingsConfig = null;
  activeResultSnapshotConfig = null;
  lastResultMapping = null;
  resultSnapshotDirty = false;
  fileDrop.classList.add("uploaded");
  hideResults();
  extraFields = [];
  rewardFields = [];
  renderExtraFieldInputs();
  renderRewardFieldInputs();
  setStatus("엑셀 내용을 읽고 있습니다...", "success");
  try {
    workbookPreview = await postFile("/api/inspect");
    if (!workbookPreview.sheets?.length) {
      throw new Error("이 파일은 구형 .xls 형식이라 바로 읽을 수 없습니다. Excel에서 .xlsx 형식으로 다시 저장한 뒤 업로드해주세요.");
    }
    populateSheetSelect(workbookPreview.sheets);
    renderSheet(workbookPreview.sheets[0].name);
    setElementHidden(mappingSection, false);
    const savedConfig = await loadSettingsForUploadedFile(file);
    if (savedConfig) {
      pendingRestoreConfig = savedConfig;
      activeFileSettingsConfig = savedConfig;
      if (hasMatchingResultSnapshot(savedConfig)) activeResultSnapshotConfig = savedConfig;
      showRestoreDialog();
    }
    setStatus("파일을 읽었습니다. 계산에 사용할 위치를 선택하세요.", "success");
  } catch (error) {
    setStatus(error.message, "error");
  }
});

sheetSelect.addEventListener("change", () => {
  renderSheet(sheetSelect.value);
  queueSaveCurrentSettings();
});
productColumn.addEventListener("focus", () => setActivePick("product"));
originalColumn.addEventListener("focus", () => setActivePick("original"));
finalPriceColumn.addEventListener("focus", () => setActivePick("final"));
categoryColumn.addEventListener("focus", () => setActivePick("category"));
addExtraFieldButton.addEventListener("click", addExtraField);
addRewardFieldButton.addEventListener("click", addRewardField);

calculateButton.addEventListener("click", async () => {
  syncTypedColumns();
  syncDynamicFieldsFromDom();
  if (!productColumn.dataset.index || !originalColumn.dataset.index || !finalPriceColumn.dataset.index) {
    setStatus("제품명 열, 기존금액 열, 최종가 열을 모두 선택해주세요.", "error");
    return;
  }
  const extraFieldValidationMessage = validateExtraFields();
  if (extraFieldValidationMessage) {
    showMessageDialog(extraFieldValidationMessage);
    return;
  }
  if (restoreResultSnapshotIfSameFile(activeResultSnapshotConfig)) {
    queueSaveCurrentSettings();
    resultSection.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }
  try {
    setCalculating(true);
    const payload = await postFile("/api/calculate", {
      sheet_name: sheetSelect.value,
      product_col: productColumn.dataset.index,
      product_color: columnFilters.product.selected,
      product_color_mode: columnFilters.product.mode,
      original_col: originalColumn.dataset.index,
      original_color: columnFilters.original.selected,
      original_color_mode: columnFilters.original.mode,
      final_price_col: finalPriceColumn.dataset.index,
      final_price_color: columnFilters.final.selected,
      final_price_color_mode: columnFilters.final.mode,
      category_col: categoryColumn.dataset.index ?? "",
      extra_fields: JSON.stringify(serializeExtraFields()),
      reward_fields: JSON.stringify(serializeRewardFields()),
      start_row: startRow.value,
      end_row: endRow.value,
    });
    const negativeRows = payload.rows.filter((row) => row.total_discount_amount < 0);
    if (negativeRows.length > 0) {
      pendingPayload = payload;
      showConfirmDialog(
        `최종가가 기존금액보다 큰 상품이 ${negativeRows.length}개 있습니다. 계속하시겠습니까?`,
      );
      return;
    }
    await finishRender(payload);
  } catch (error) {
    showMessageDialog(error.message);
  } finally {
    setCalculating(false);
  }
});


productSearch.addEventListener("input", renderFilteredResults);
sortField.addEventListener("change", renderFilteredResults);
sortDirection.addEventListener("change", renderFilteredResults);
resultSection.addEventListener("click", (event) => {
  if (isEditMode) return;
  if (!selectedRowId) return;
  const isResultRow = event.target.closest("tbody tr:not(.category-group-row)");
  const isInteractive = event.target.closest("button, input, select, textarea, [data-role='select'], [data-role='select-menu']");
  if (isResultRow || isInteractive) return;
  selectedRowId = null;
  updateSelectedResultRow();
});
productFilterButton.addEventListener("click", (event) => {
  event.stopPropagation();
  setElementHidden(productFilterPanel, !productFilterPanel.hidden ? true : false);
});
productFilterPanel.addEventListener("click", (event) => {
  event.stopPropagation();
});
addProductFilterButton.addEventListener("click", () => {
  productNameFilters.push("");
  renderProductFilterInputs();
  queueSaveCurrentSettings();
});
removeParenthesesText.addEventListener("change", () => { renderFilteredResults(); queueSaveCurrentSettings(); });
removeBracketsText.addEventListener("change", () => { renderFilteredResults(); queueSaveCurrentSettings(); });
removeAfterDelimiter.addEventListener("change", () => {
  if (removeAfterDelimiter.checked) removeBeforeDelimiter.checked = false;
  updateDelimiterOptionsVisibility();
  renderFilteredResults();
  queueSaveCurrentSettings();
});
removeBeforeDelimiter.addEventListener("change", () => {
  if (removeBeforeDelimiter.checked) removeAfterDelimiter.checked = false;
  updateDelimiterOptionsVisibility();
  renderFilteredResults();
  queueSaveCurrentSettings();
});
removeLeadingText.addEventListener("change", () => {
  setElementHidden(leadingTextOptions, !removeLeadingText.checked);
  renderFilteredResults();
  queueSaveCurrentSettings();
});
leadingTextValue.addEventListener("input", () => { renderFilteredResults(); queueSaveCurrentSettings(); });
delimiterValue.addEventListener("input", () => { renderFilteredResults(); queueSaveCurrentSettings(); });
function updateDelimiterOptionsVisibility() {
  setElementHidden(delimiterOptions, !(removeAfterDelimiter.checked || removeBeforeDelimiter.checked));
}
document.addEventListener("click", (event) => {
  if (!event.target.closest(".search-with-filter")) {
    setElementHidden(productFilterPanel, true);
  }
  if (!event.target.closest(".visible-columns-wrap")) {
    setElementHidden(visibleColumnsPanel, true);
  }
});
visibleColumnsButton.addEventListener("click", (event) => {
  event.stopPropagation();
  setElementHidden(visibleColumnsPanel, !visibleColumnsPanel.hidden ? true : false);
});
visibleColumnsPanel.querySelectorAll("[data-visible-column]").forEach((checkbox) => {
  checkbox.addEventListener("change", () => {
    const column = checkbox.dataset.visibleColumn;
    if (checkbox.checked) hiddenColumns.delete(column);
    else hiddenColumns.add(column);
    applyVisibleColumns();
    updateFloatingResultHeader();
  });
});
showWonSuffix.addEventListener("change", () => {
  renderFilteredResults();
  queueSaveCurrentSettings();
});
showDiscountMinus.addEventListener("change", () => {
  renderFilteredResults();
  queueSaveCurrentSettings();
});
highlightFinalPrices.addEventListener("change", () => {
  applyPriceHighlighting();
  queueSaveCurrentSettings();
});
titleCaseProductName.addEventListener("change", () => { renderFilteredResults(); queueSaveCurrentSettings(); });
autoLoadColorFilters.addEventListener("change", () => {
  if (autoLoadColorFilters.checked) loadAllColumnColors();
  else Object.keys(columnFilters).forEach(resetColumnColors);
  queueSaveCurrentSettings();
});
groupSimilarProducts.addEventListener("change", () => {
  renderFilteredResults();
  queueSaveCurrentSettings();
});
productNameResizer.addEventListener("pointerdown", startProductColumnResize);
editResultsButton.addEventListener("click", () => {
  if (isEditMode) {
    commitInlineEdit(false);
    setEditMode(false);
    editRowsSnapshot = null;
    editSortSnapshot = null;
    renderFilteredResults();
    queueSaveCurrentSettings();
    return;
  }
  editRowsSnapshot = cloneRows(currentRows);
  editSortSnapshot = { field: sortField.value, direction: sortDirection.value };
  setEditMode(true);
  renderFilteredResults();
});
cancelEditResultsButton.addEventListener("click", () => {
  activeInlineEdit = null;
  if (editRowsSnapshot) currentRows = cloneRows(editRowsSnapshot);
  if (editSortSnapshot) {
    sortField.value = editSortSnapshot.field;
    sortDirection.value = editSortSnapshot.direction;
    syncCustomSelect(sortField);
    syncCustomSelect(sortDirection);
  }
  editRowsSnapshot = null;
  editSortSnapshot = null;
  setEditMode(false);
  renderFilteredResults();
});
if (addResultRowButton) {
  addResultRowButton.addEventListener("click", () => {
    commitInlineEdit(false);
    const row = createEmptyResultRow();
    currentRows.push(row);
    selectedRowId = String(row.row_id);
    setSortFieldToCustomOrder();
    renderFilteredResults();
    requestAnimationFrame(() => {
      const rowElement = resultBody.querySelector(`tr[data-row-id="${cssEscape(String(row.row_id))}"]`);
      const productButton = rowElement?.querySelector('[data-column="product_name"] [data-copy-value]');
      if (productButton) startInlineEdit(productButton, row);
    });
  });
}
expandResultsButton.addEventListener("click", () => {
  resultSection.classList.toggle("expanded-results");
  const expanded = resultSection.classList.contains("expanded-results");
  expandResultsButton.textContent = expanded ? "축소" : "확장";
  scheduleResultColumnLayoutSync();
});
Object.values(columnFilters).forEach((filter) => {
  filter.trigger.addEventListener("click", () => {
    closeAllColorMenus(filter.choices);
    setElementHidden(filter.choices, !filter.choices.hidden ? true : false);
  });
});
document.addEventListener("click", (event) => {
  if (!event.target.closest(".inline-color-picker")) closeAllColorMenus();
});

function populateSheetSelect(sheets) {
  sheetSelect.innerHTML = "";
  sheets.forEach((sheet) => {
    const option = document.createElement("option");
    option.value = sheet.name;
    option.textContent = sheet.name;
    sheetSelect.appendChild(option);
  });
  syncCustomSelect(sheetSelect);
}

function renderSheet(sheetName) {
  const sheet = workbookPreview.sheets.find((item) => item.name === sheetName);
  previewTable.innerHTML = "";
  [productColumn, originalColumn, finalPriceColumn, categoryColumn].forEach((input) => {
    input.value = "";
    delete input.dataset.index;
  });
  Object.keys(columnFilters).forEach(resetColumnColors);
  startRow.value = sheet.suggested_start_row || 1;
  endRow.value = sheet.suggested_end_row || sheet.total_rows;
  setActivePick("product");
  updateCalculateButtonState();

  const header = document.createElement("thead");
  const headerRow = document.createElement("tr");
  headerRow.appendChild(document.createElement("th"));
  sheet.preview[0]?.forEach((_, visibleColIndex) => {
    const colIndex = sheet.preview_cols?.[visibleColIndex] ?? visibleColIndex;
    const th = document.createElement("th");
    th.textContent = columnLabel(colIndex);
    headerRow.appendChild(th);
  });
  header.appendChild(headerRow);
  previewTable.appendChild(header);

  const body = document.createElement("tbody");
  const mergeMaps = getPreviewMergeMaps(sheet);
  sheet.preview.forEach((row, visibleRowIndex) => {
    const rowNumber = sheet.preview_rows?.[visibleRowIndex] ?? visibleRowIndex + 1;
    const tr = document.createElement("tr");
    const rowHeader = document.createElement("th");
    rowHeader.textContent = String(rowNumber);
    tr.appendChild(rowHeader);
    row.forEach((cell, visibleColIndex) => {
      const colIndex = sheet.preview_cols?.[visibleColIndex] ?? visibleColIndex;
      const mergeKey = `${rowNumber}:${colIndex}`;
      if (mergeMaps.skip.has(mergeKey)) return;
      const mergeCell = mergeMaps.render.get(mergeKey);
      const td = document.createElement("td");
      td.textContent = mergeCell?.value ?? cell ?? "";
      td.dataset.row = String(mergeCell?.source_row ?? rowNumber);
      td.dataset.col = String(mergeCell?.source_col ?? colIndex);
      if (mergeCell?.rowspan > 1) td.rowSpan = mergeCell.rowspan;
      if (mergeCell?.colspan > 1) td.colSpan = mergeCell.colspan;
      if (mergeCell) td.classList.add("merged-cell");
      td.addEventListener("click", () => selectCell(td));
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
  previewTable.appendChild(body);
}

function selectCell(cell) {
  const colIndex = Number(cell.dataset.col);
  const rowNumber = Number(cell.dataset.row);
  const target =
    activePick === "extra"
      ? activeExtraFieldTarget
      : activePick === "product"
      ? productColumn
      : activePick === "original"
        ? originalColumn
        : activePick === "final"
          ? finalPriceColumn
          : categoryColumn;

  if (!target) return;
  target.value = `${columnLabel(colIndex)}열`;
  target.dataset.index = String(colIndex);
  target.dispatchEvent(new Event("input", { bubbles: true }));
  if (activePick === "extra") persistDynamicColumnInput(target, colIndex);
  if (activePick === "product") {
    startRow.value = rowNumber;
  }
  if (activePick !== "extra" && autoLoadColorFilters.checked) loadColumnColors(activePick);

  previewTable.querySelectorAll(`td[data-col="${colIndex}"]`).forEach((node) => {
    node.classList.add(
      activePick === "product"
        ? "selected-product"
        : activePick === "original"
          ? "selected-original"
          : "selected-discount",
    );
  });

  if (activePick === "product") setActivePick("original");
  else if (activePick === "original") setActivePick("final");
  updateCalculateButtonState();
  queueSaveCurrentSettings();
}

[productColumn, originalColumn, finalPriceColumn, categoryColumn].forEach((input) => {
  input.addEventListener("input", () => {
    const parsed = parseColumnInput(input.value);
    if (parsed === null) {
      delete input.dataset.index;
      resetColumnColors(columnKeyForInput(input));
      return;
    }
    input.dataset.index = String(parsed);
    if (autoLoadColorFilters.checked) loadColumnColors(columnKeyForInput(input));
    else resetColumnColors(columnKeyForInput(input));
    updateCalculateButtonState();
    queueSaveCurrentSettings();
  });
});

startRow.addEventListener("change", () => {
  loadAllColumnColors();
  queueSaveCurrentSettings();
});
endRow.addEventListener("change", () => {
  loadAllColumnColors();
  queueSaveCurrentSettings();
});
sheetSelect.addEventListener("change", () => {
  Object.keys(columnFilters).forEach(resetColumnColors);
});

function setActivePick(type) {
  activePick = type;
  currentPickLabel.textContent =
    type === "product"
      ? "제품명"
      : type === "original"
        ? "기존금액"
        : type === "final"
          ? "최종가"
          : type === "category"
            ? "카테고리"
            : "추가 항목";
}

function columnKeyForInput(input) {
  return Object.keys(columnFilters).find((key) => columnFilters[key].input === input);
}

async function loadAllColumnColors() {
  if (!autoLoadColorFilters.checked) {
    Object.keys(columnFilters).forEach(resetColumnColors);
    return;
  }
  setColorLoading(true);
  try {
    await Promise.all(Object.keys(columnFilters).map(loadColumnColors));
  } finally {
    setColorLoading(false);
  }
}

function setColorLoading(isLoading) {
  if (!colorLoadOptionText) return;
  colorLoadOptionText.textContent = isLoading ? "불러오는 중..." : "열 색상값 불러오기";
}

async function loadColumnColors(key) {
  syncTypedColumns();
  const filter = columnFilters[key];
  if (!filter?.input.dataset.index || !uploadedFile) return;

  try {
    const previousSelection = filter.selected;
    const previousMode = filter.mode;
    const payload = await postFile("/api/column-colors", {
      sheet_name: sheetSelect.value,
      column_index: filter.input.dataset.index,
      start_row: startRow.value,
      end_row: endRow.value,
    });
    const usableColors = payload.colors.filter((color) => color.value !== "FFFFFF");
    const availableColors = usableColors.map((color) => color.value);
    filter.selected = availableColors.includes(previousSelection) ? previousSelection : "";
    filter.mode = filter.selected ? previousMode || "include" : "include";
    filter.choices.innerHTML = buildColorModeControls(filter) + `<button class="color-choice${filter.selected ? "" : " selected"}" type="button" data-color="">\uc804\uccb4</button>`;
    usableColors.forEach((color) => {
      const button = document.createElement("button");
      button.className = `color-choice${filter.selected === color.value ? " selected" : ""}`;
      button.type = "button";
      button.dataset.color = color.value;
      button.innerHTML = `
        <span class="color-swatch" style="background:#${color.value}"></span>
        <span>${color.count}개</span>
        <span class="color-tooltip">#${color.value}</span>
      `;
      filter.choices.appendChild(button);
    });
    setElementHidden(filter.picker, usableColors.length === 0);
    bindColorChoiceEvents(key);
    updateColorTrigger(key);
  } catch {
    resetColumnColors(key);
  }
}

function resetColumnColors(key) {
  const filter = columnFilters[key];
  filter.selected = "";
  filter.mode = "include";
  setElementHidden(filter.picker, true);
  setElementHidden(filter.choices, true);
  filter.choices.innerHTML = buildColorModeControls(filter) + `<button class="color-choice selected" type="button" data-color="">\uc804\uccb4</button>`;
  updateColorTrigger(key);
}

function closeAllColorMenus(exceptChoices = null) {
  Object.values(columnFilters).forEach((filter) => {
    if (filter.choices !== exceptChoices) setElementHidden(filter.choices, true);
  });
}

function buildColorModeControls(filter) {
  return `
    <div class="color-mode-controls" role="group" aria-label="\uc0c9\uc0c1 \ud544\ud130 \ubc29\uc2dd">
      <button class="color-mode-button${filter.mode !== "exclude" ? " selected" : ""}" type="button" data-color-mode="include">\uc120\ud0dd \uc0c9\uc0c1\ub9cc \ubcf4\uae30</button>
      <button class="color-mode-button${filter.mode === "exclude" ? " selected" : ""}" type="button" data-color-mode="exclude">\uc120\ud0dd \uc0c9\uc0c1 \uc81c\uc678</button>
    </div>
  `;
}

function bindColorChoiceEvents(key) {
  const filter = columnFilters[key];
  filter.choices.querySelectorAll(".color-mode-button").forEach((button) => {
    button.addEventListener("click", (event) => {
      event.stopPropagation();
      filter.mode = button.dataset.colorMode === "exclude" ? "exclude" : "include";
      filter.choices.querySelectorAll(".color-mode-button").forEach((node) => {
        node.classList.toggle("selected", node === button);
      });
      clearRestoredResultSnapshot();
      updateColorTrigger(key);
      queueSaveCurrentSettings();
    });
  });
  filter.choices.querySelectorAll(".color-choice").forEach((button) => {
    button.addEventListener("click", () => {
      const nextColor = button.dataset.color;
      if (!nextColor) filter.mode = "include";
      if (filter.selected !== nextColor) clearRestoredResultSnapshot();
      filter.selected = nextColor;
      filter.choices.querySelectorAll(".color-choice").forEach((node) => {
        node.classList.toggle("selected", node === button);
      });
      updateColorTrigger(key);
      setElementHidden(filter.choices, true);
      queueSaveCurrentSettings();
    });
  });
}

function updateColorTrigger(key) {
  const filter = columnFilters[key];
  if (!filter.selected) {
    filter.trigger.textContent = "+";
    filter.trigger.classList.remove("has-selection");
    return;
  }
  const selectedButton = filter.choices.querySelector(`[data-color="${filter.selected}"]`);
  filter.trigger.innerHTML = selectedButton
    ? `<span class="color-swatch" style="background:#${filter.selected}"></span>${filter.mode === "exclude" ? '<span class="color-exclude-mark">\u00d7</span>' : ""}`
    : "+";
  filter.trigger.title = filter.selected ? `#${filter.selected} ${filter.mode === "exclude" ? "\uc81c\uc678" : "\ud3ec\ud568"}` : "";
  filter.trigger.classList.add("has-selection");
}

async function postFile(path, extraFields = {}) {
  if (path === "/api/inspect") return inspectExcelInBrowser();
  if (path === "/api/column-colors") return columnColorsInBrowser(extraFields);
  if (path === "/api/calculate") return calculateExcelInBrowser(extraFields);
  throw new Error("지원하지 않는 처리 요청입니다.");
}

function renderResults(payload) {
  itemCount.textContent = `${payload.rows.length}개`;
  currentRows = payload.rows;
  lastResultMapping = getCurrentMappingSnapshot();
  resultSnapshotDirty = false;
  editRowsSnapshot = null;
  activeInlineEdit = null;
  setEditMode(false);
  extraFields = payload.extra_fields ?? extraFields;
  rewardFields = payload.reward_fields ?? rewardFields;
  collapsedCategories.clear();
  resultSection.classList.toggle("single-extra-field", extraFields.length === 1);
  resultSection.classList.toggle("multiple-extra-fields", extraFields.length > 1);
  resultSection.classList.toggle("no-reward-fields", rewardFields.length === 0);
  resultSection.classList.toggle("has-reward-fields", rewardFields.length > 0);
  productSearch.value = "";
  renderProductFilterInputs();
  sortField.value = "none";
  sortDirection.value = "desc";
  removeCustomSortOption();
  syncCustomSelect(sortField);
  syncCustomSelect(sortDirection);
  setElementHidden(resultSection, false);
  renderFilteredResults();
  refreshResultLayoutAfterRender();
}

function refreshResultLayoutAfterRender() {
  normalizeInitialExtraTableWidth();
  requestAnimationFrame(() => {
    normalizeInitialExtraTableWidth();
    applyProductColumnWidth();
    updateFloatingResultHeader();
    requestAnimationFrame(() => {
      applyProductColumnWidth();
      updateFloatingResultHeader();
    });
  });
}

function renderFilteredResults() {
  resultBody.innerHTML = "";
  itemCount.textContent = `${currentRows.length}개`;

  const keyword = productSearch.value.trim().toLowerCase();
  const rows = currentRows
    .filter((row) => row.product_name.toLowerCase().includes(keyword))
    .sort((a, b) => {
      if (sortField.value === "none" || sortField.value === "custom") return 0;
      const diff = a[sortField.value] - b[sortField.value];
      return sortDirection.value === "asc" ? diff : -diff;
    });

  const displayRows = groupSimilarProducts.checked && !isEditMode ? groupSimilarResultRows(rows) : rows;
  const groupedRows = categoryColumn.dataset.index ? groupRowsByCategory(displayRows) : [["", displayRows]];

  groupedRows.forEach(([category, categoryRows]) => {
    const categoryKey = category || "미분류";
    const isCollapsed = collapsedCategories.has(categoryKey);
    if (categoryColumn.dataset.index) {
      const groupRow = document.createElement("tr");
      groupRow.className = `category-group-row${isCollapsed ? " collapsed" : ""}`;
      groupRow.dataset.category = categoryKey;
      groupRow.innerHTML = `
        <td colspan="${getResultColumnSpan()}">
          <button class="category-toggle" type="button" aria-expanded="${!isCollapsed}">
            <span class="category-caret">${isCollapsed ? ">" : "⌄"}</span>
            <span>${escapeHtml(categoryKey)}</span>
            <span class="category-count">${categoryRows.length}개</span>
          </button>
        </td>`;
      groupRow.querySelector(".category-toggle").addEventListener("click", (event) => {
        event.stopPropagation();
        if (collapsedCategories.has(categoryKey)) collapsedCategories.delete(categoryKey);
        else collapsedCategories.add(categoryKey);
        renderFilteredResults();
      });
      resultBody.appendChild(groupRow);
    }
    if (isCollapsed) return;
    categoryRows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.dataset.rowId = row.row_id;
    tr.classList.toggle("selected-result-row", String(selectedRowId) === String(row.row_id));
    const extraCells = extraFields
      .map((field) => `<td class="extra-column" data-column="extra_${field.id}">${buildCopyCell(formatExtraValue(row.extra_values?.[field.id]))}</td>`)
      .join("");
    const rewardCells = rewardFields
      .map((field, index) => `<td class="reward-column${index === 0 ? " reward-group-start" : ""}" data-column="reward_${field.id}">${buildCopyCell(formatCurrency(row.reward_values?.[field.id] ?? 0))}</td>`)
      .join("");
    const dragControl = isEditMode ? `<button class="row-drag-handle" type="button" draggable="true" data-drag-row="${escapeAttribute(row.row_id)}" aria-label="행 이동">⋮⋮</button>` : "";
    const deleteControl = isEditMode ? `<button class="row-delete-button" type="button" data-delete-row="${escapeAttribute(row.row_id)}" aria-label="행 삭제">×</button>` : "";
    tr.innerHTML = `
      <td data-column="product_name"><div class="product-cell-content">${dragControl}${deleteControl}${buildProductNameCopyCell(displayProductName(row.product_name ?? ""))}</div></td>
      <td data-column="original_price">${buildCopyCell(formatCurrency(row.original_price))}</td>
      <td data-column="discount_amount">${buildCopyCell(formatCurrency(row.discount_amount))}</td>
      ${extraCells}
      ${rewardCells}
      <td class="reward-column${rewardFields.length === 0 ? " reward-group-start" : ""}" data-column="total_reward_amount">${buildCopyCell(formatCurrency(row.total_reward_amount ?? 0))}</td>
      <td class="reward-column reward-group-end" data-column="effective_price">${buildCopyCell(formatCurrency(row.effective_price ?? row.discount_amount))}</td>
      <td data-column="total_discount_amount">${buildCopyCell(formatTotalDiscount(row.total_discount_amount))}</td>
      <td data-column="discount_rate">${buildCopyCell(formatRate(row.discount_rate))}</td>
    `;
    tr.querySelectorAll("[data-copy-value]").forEach((button) => {
      button.addEventListener("pointerdown", (event) => {
        if (!isEditMode) return;
        event.preventDefault();
        event.stopPropagation();
        startInlineEdit(event.currentTarget, row);
      });
      button.addEventListener("click", async (event) => {
        event.stopPropagation();
        if (isEditMode) {
          return;
        }
        selectedRowId = String(row.row_id);
        updateSelectedResultRow();
        await copyText(event.currentTarget.dataset.copyValue);
      });
    });
    tr.querySelector("[data-delete-row]")?.addEventListener("click", (event) => {
      event.stopPropagation();
      commitInlineEdit(false);
      currentRows = currentRows.filter((item) => String(item.row_id) !== String(row.row_id));
      if (String(selectedRowId) === String(row.row_id)) selectedRowId = null;
      renderFilteredResults();
    });
    tr.querySelector("[data-drag-row]")?.addEventListener("dragstart", (event) => {
      if (!isEditMode) {
        event.preventDefault();
        return;
      }
      commitInlineEdit(false);
      draggedResultRowId = String(row.row_id);
      tr.classList.add("dragging-result-row");
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", draggedResultRowId);
    });
    tr.addEventListener("dragover", (event) => {
      if (!isEditMode || !draggedResultRowId || draggedResultRowId === String(row.row_id)) return;
      event.preventDefault();
      tr.classList.add("drag-over-result-row");
      event.dataTransfer.dropEffect = "move";
    });
    tr.addEventListener("dragleave", () => {
      tr.classList.remove("drag-over-result-row");
    });
    tr.addEventListener("drop", (event) => {
      if (!isEditMode || !draggedResultRowId) return;
      event.preventDefault();
      tr.classList.remove("drag-over-result-row");
      moveResultRow(draggedResultRowId, String(row.row_id));
    });
    tr.addEventListener("dragend", () => {
      draggedResultRowId = null;
      tr.classList.remove("dragging-result-row", "drag-over-result-row");
    });
    tr.addEventListener("click", () => {
      if (isEditMode) return;
      if (!selectedRowId) return;
      selectedRowId = null;
      updateSelectedResultRow();
    });
      resultBody.appendChild(tr);
    });
  });
  renderExtraResultHeaders();
  renderRewardResultHeaders();
  applyRewardColumnVisibility();
  applyVisibleColumns();
  applyPriceHighlighting();
  applyProductColumnWidth();
  setElementHidden(emptyResultMessage, rows.length > 0);
  normalizeInitialExtraTableWidth();
}

function normalizeInitialExtraTableWidth() {
  if (productColumnWidth !== null || resultSection.hidden) return;
  requestAnimationFrame(() => {
    if (productColumnWidth !== null || resultSection.hidden) return;
    const productWidth = getInitialProductColumnWidth();
    if (productWidth <= 0) return;
    productColumnWidth = productWidth;
    productColumnDragWidth = productWidth;
    productColumnUserResized = false;
    applyProductColumnWidth();
  });
}

function getInitialProductColumnWidth() {
  const table = tableWrap?.querySelector("table");
  const availableWidth = tableWrap?.clientWidth || 0;
  if (!table || !availableWidth) return 0;
  const otherColumnsWidth = getVisibleResultHeaders()
    .filter((header) => header.dataset.column !== "product_name")
    .reduce((sum, header) => sum + getResultColumnPixelWidth(header), 0);
  const minWidth = 180;
  return Math.max(minWidth, Math.floor(availableWidth - otherColumnsWidth));
}

function getResultColumnSpan() {
  const baseColumns = ["product_name", "original_price", "discount_amount", "total_discount_amount", "discount_rate"]
    .filter((column) => !hiddenColumns.has(column)).length;
  const extraColumns = extraFields.filter((field) => !hiddenColumns.has(`extra_${field.id}`)).length;
  const rewardColumns = rewardFields.length > 0
    ? rewardFields.filter((field) => !hiddenColumns.has(`reward_${field.id}`)).length
      + ["total_reward_amount", "effective_price"].filter((column) => !hiddenColumns.has(column)).length
    : 0;
  return Math.max(1, baseColumns + extraColumns + rewardColumns);
}

function updateSelectedResultRow() {
  resultBody.querySelectorAll("tr[data-row-id]").forEach((row) => {
    row.classList.toggle("selected-result-row", String(row.dataset.rowId) === String(selectedRowId));
  });
}

function applyRewardColumnVisibility() {
  const shouldShowRewards = rewardFields.length > 0;
  ["total_reward_amount", "effective_price"].forEach((column) => {
    document.querySelectorAll(`[data-column="${column}"]`).forEach((node) => {
      node.classList.toggle("hidden-result-column", !shouldShowRewards || hiddenColumns.has(column));
    });
  });
  visibleColumnsPanel.querySelectorAll('[data-visible-column="total_reward_amount"], [data-visible-column="effective_price"]').forEach((checkbox) => {
    setElementHidden(checkbox.closest("label"), !shouldShowRewards);
  });
  document.querySelectorAll('[data-column="total_reward_amount"]').forEach((node) => {
    node.classList.toggle("reward-group-start", rewardFields.length === 0);
  });
}

async function finishRender(payload) {
  renderResults(payload);
  await saveCurrentSettings();
  setStatus(`계산 완료: ${payload.rows.length}개 상품을 찾았습니다.`, "success");
  resultSection.scrollIntoView({ behavior: "smooth", block: "start" });
}

function buildCopyCell(value, displayHtml = null) {
  const displayValue = displayHtml ?? escapeHtml(value);
  const copyValue = escapeAttribute(value);
  return `<button class="copy-value" type="button" data-copy-value="${copyValue}" aria-label="${copyValue}">${displayValue}</button>`;
}

function buildProductNameCopyCell(value) {
  return buildCopyCell(value, buildGroupedProductDisplayHtml(value));
}

function buildGroupedProductDisplayHtml(value) {
  const text = String(value ?? "");
  if (!text.includes(GROUPED_PRODUCT_JOINER)) return escapeHtml(text);
  return text
    .split(GROUPED_PRODUCT_JOINER)
    .map((part) => escapeHtml(part))
    .join('<span class="grouped-product-separator" aria-hidden="true"></span>');
}

function removeCustomSortOption() {
  if (sortField.value === "custom") sortField.value = "none";
  [...sortField.options].forEach((option) => {
    if (option.value === "custom") option.remove();
  });
}
function setSortFieldToCustomOrder() {
  if (![...sortField.options].some((option) => option.value === "custom")) {
    const option = document.createElement("option");
    option.value = "custom";
    option.textContent = "사용자 설정";
    sortField.appendChild(option);
  }
  sortField.value = "custom";
  syncCustomSelect(sortField);
}

function moveResultRow(draggedId, targetId) {
  if (draggedId === targetId) return;
  const fromIndex = currentRows.findIndex((row) => String(row.row_id) === String(draggedId));
  const targetIndex = currentRows.findIndex((row) => String(row.row_id) === String(targetId));
  if (fromIndex < 0 || targetIndex < 0) return;

  const targetRow = currentRows[targetIndex];
  const [movedRow] = currentRows.splice(fromIndex, 1);
  movedRow.category_name = targetRow.category_name ?? movedRow.category_name ?? "";

  let targetIndexAfterRemoval = currentRows.findIndex((row) => String(row.row_id) === String(targetId));
  if (targetIndexAfterRemoval < 0) targetIndexAfterRemoval = currentRows.length;

  const insertIndex = fromIndex < targetIndex ? targetIndexAfterRemoval + 1 : targetIndexAfterRemoval;
  currentRows.splice(insertIndex, 0, movedRow);

  selectedRowId = String(movedRow.row_id);
  draggedResultRowId = null;
  setSortFieldToCustomOrder();
  renderFilteredResults();
  queueSaveCurrentSettings();
}

function getCategoryForNewResultRow() {
  if (!categoryColumn.dataset.index) return "";
  if (selectedRowId) {
    const selectedRow = currentRows.find((row) => String(row.row_id) === String(selectedRowId));
    if (selectedRow?.category_name) return selectedRow.category_name;
  }
  if (lastActiveCategoryKey && lastActiveCategoryKey !== "???") return lastActiveCategoryKey;
  const groups = groupRowsByCategory(currentRows);
  const lastGroup = groups[groups.length - 1];
  return lastGroup?.[0] === "???" ? "" : (lastGroup?.[0] ?? "");
}

function createEmptyResultRow(categoryName = "") {
  const extraValues = Object.fromEntries(extraFields.map((field) => [field.id, 0]));
  const rewardValues = Object.fromEntries(rewardFields.map((field) => [field.id, 0]));
  const row = {
    row_id: `manual_${Date.now()}_${Math.random().toString(36).slice(2)}`,
    product_name: "새 상품",
    category_name: categoryName || "",
    base_original_price: 0,
    has_extra_price_formula: extraFields.length > 0,
    original_price: 0,
    discount_amount: 0,
    extra_values: extraValues,
    reward_values: rewardValues,
    total_reward_amount: 0,
    effective_price: 0,
    total_discount_amount: 0,
    discount_rate: 0,
  };
  recalculateDerivedValues(row);
  return row;
}

function setEditMode(enabled) {
  isEditMode = Boolean(enabled);
  draggedResultRowId = null;
  editResultsButton.classList.toggle("active", isEditMode);
  editResultsButton.textContent = isEditMode ? "저장하기" : "편집";
  resultSection.classList.toggle("edit-mode", isEditMode);
  setElementHidden(cancelEditResultsButton, !isEditMode);
  setElementHidden(editAddRowWrap, !isEditMode);
}

function cloneRows(rows) {
  return JSON.parse(JSON.stringify(rows ?? []));
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function escapeAttribute(value) {
  return escapeHtml(value);
}

async function getExcelWorkbook() {
  if (excelWorkbook) return excelWorkbook;
  if (isLegacyXlsFile(uploadedFile)) {
    throw new Error("이 파일은 구형 .xls 형식이라 바로 읽을 수 없습니다. Excel에서 .xlsx 형식으로 다시 저장한 뒤 업로드해주세요.");
  }
  if (!window.ExcelJS) {
    throw new Error("엑셀 처리 모듈을 불러오지 못했습니다. 인터넷 연결을 확인한 뒤 다시 열어주세요.");
  }
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await uploadedFile.arrayBuffer());
  excelWorkbook = workbook;
  return workbook;
}

async function inspectExcelInBrowser() {
  excelWorkbook = null;
  const workbook = await getExcelWorkbook();
  const sheets = workbook.worksheets.map((worksheet) => {
    const visibleRows = [];
    for (let rowNumber = 1; rowNumber <= worksheet.rowCount && visibleRows.length < 40; rowNumber += 1) {
      if (!isHiddenRow(worksheet, rowNumber)) visibleRows.push(rowNumber);
    }
    const visibleCols = [];
    const maxColumn = Math.max(worksheet.columnCount, 1);
    for (let colNumber = 1; colNumber <= maxColumn && visibleCols.length < 24; colNumber += 1) {
      if (!isHiddenColumn(worksheet, colNumber - 1)) visibleCols.push(colNumber - 1);
    }
    const preview = visibleRows.map((rowNumber) =>
      visibleCols.map((colIndex) => cleanExcelCellValue(getDisplayCell(worksheet, rowNumber, colIndex).value)),
    );
    const previewMerges = getPreviewMerges(worksheet, visibleRows, visibleCols);
    const [suggestedStart, suggestedEnd] = suggestBrowserRowBounds(worksheet);
    return {
      name: worksheet.name,
      preview,
      preview_rows: visibleRows,
      preview_cols: visibleCols,
      preview_merges: previewMerges,
      total_rows: suggestedEnd,
      total_cols: worksheet.columnCount,
      suggested_start_row: suggestedStart,
      suggested_end_row: suggestedEnd,
    };
  });
  return { sheets };
}

function getWorksheetByName(name) {
  const worksheet = excelWorkbook?.getWorksheet(name);
  if (!worksheet) throw new Error("선택한 시트를 찾지 못했습니다.");
  return worksheet;
}

function isHiddenRow(worksheet, rowNumber) {
  const row = worksheet.getRow(rowNumber);
  return Boolean(row.hidden || row.height === 0);
}

function isHiddenColumn(worksheet, columnIndex) {
  const column = worksheet.getColumn(columnIndex + 1);
  return Boolean(column.hidden || column.width === 0);
}

function getWorksheetMergeRanges(worksheet) {
  if (worksheetMergeRangeCache.has(worksheet)) return worksheetMergeRangeCache.get(worksheet);
  const rawMerges = [];
  if (worksheet?._merges) rawMerges.push(...Object.values(worksheet._merges));
  if (Array.isArray(worksheet?.model?.merges)) rawMerges.push(...worksheet.model.merges);
  const seen = new Set();
  const ranges = rawMerges
    .map(parseMergeRange)
    .filter(Boolean)
    .filter((range) => {
      const key = `${range.top}:${range.left}:${range.bottom}:${range.right}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return range.bottom > range.top || range.right > range.left;
    });
  worksheetMergeRangeCache.set(worksheet, ranges);
  return ranges;
}

function parseMergeRange(merge) {
  if (!merge) return null;
  if (merge.model) return parseMergeRange(merge.model);
  if (typeof merge === "string") {
    const [start, end = start] = merge.split(":");
    const startCell = parseCellAddress(start);
    const endCell = parseCellAddress(end);
    if (!startCell || !endCell) return null;
    return {
      top: Math.min(startCell.row, endCell.row),
      left: Math.min(startCell.col, endCell.col),
      bottom: Math.max(startCell.row, endCell.row),
      right: Math.max(startCell.col, endCell.col),
    };
  }
  const top = merge.top ?? merge.tl?.nativeRow ?? merge.tl?.row;
  const left = merge.left ?? merge.tl?.nativeCol ?? merge.tl?.col;
  const bottom = merge.bottom ?? merge.br?.nativeRow ?? merge.br?.row;
  const right = merge.right ?? merge.br?.nativeCol ?? merge.br?.col;
  if ([top, left, bottom, right].some((value) => !Number.isFinite(Number(value)))) return null;
  return { top: Number(top), left: Number(left), bottom: Number(bottom), right: Number(right) };
}

function parseCellAddress(address) {
  const match = String(address).match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!match) return null;
  return { row: Number(match[2]), col: columnNameToNumber(match[1]) };
}

function columnNameToNumber(name) {
  return String(name).toUpperCase().split("").reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0);
}

function getMergeRangeForCell(worksheet, rowNumber, columnIndex) {
  const colNumber = columnIndex + 1;
  return getWorksheetMergeRanges(worksheet).find(
    (range) => rowNumber >= range.top && rowNumber <= range.bottom && colNumber >= range.left && colNumber <= range.right,
  );
}

function getDisplayCell(worksheet, rowNumber, columnIndex) {
  const range = getMergeRangeForCell(worksheet, rowNumber, columnIndex);
  if (range) return worksheet.getRow(range.top).getCell(range.left);
  return worksheet.getRow(rowNumber).getCell(columnIndex + 1);
}

function getPreviewMerges(worksheet, visibleRows, visibleCols) {
  return getWorksheetMergeRanges(worksheet).map((range) => {
    const rows = visibleRows.filter((rowNumber) => rowNumber >= range.top && rowNumber <= range.bottom);
    const cols = visibleCols.filter((colIndex) => colIndex + 1 >= range.left && colIndex + 1 <= range.right);
    if (rows.length === 0 || cols.length === 0) return null;
    return {
      render_row: rows[0],
      render_col: cols[0],
      source_row: range.top,
      source_col: range.left - 1,
      rowspan: rows.length,
      colspan: cols.length,
      covered: rows.flatMap((rowNumber) => cols.map((colIndex) => `${rowNumber}:${colIndex}`)).filter((key) => key !== `${rows[0]}:${cols[0]}`),
      value: cleanExcelCellValue(worksheet.getRow(range.top).getCell(range.left).value),
    };
  }).filter(Boolean);
}

function getPreviewMergeMaps(sheet) {
  const render = new Map();
  const skip = new Set();
  (sheet.preview_merges ?? []).forEach((merge) => {
    render.set(`${merge.render_row}:${merge.render_col}`, merge);
    (merge.covered ?? []).forEach((key) => skip.add(key));
  });
  return { render, skip };
}

function ensureVisibleBrowserColumn(worksheet, columnIndex, label) {
  if (isHiddenColumn(worksheet, columnIndex)) {
    throw new Error(`${label}로 선택한 열이 엑셀에서 숨김 처리되어 있습니다. 숨김 해제 후 다시 선택해주세요.`);
  }
}

function cleanExcelCellValue(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "object") {
    if ("result" in value) return cleanExcelCellValue(value.result);
    if ("text" in value) return cleanExcelCellValue(value.text);
    if ("richText" in value) return value.richText.map((part) => part.text).join("");
    if ("hyperlink" in value && "text" in value) return cleanExcelCellValue(value.text);
    if ("formula" in value || "sharedFormula" in value) return 0;
    if ("error" in value) return "";
    if (Object.keys(value).length === 0) return "";
    return "";
  }
  return value;
}

function parseBrowserNumber(value) {
  const cleanedValue = cleanExcelCellValue(value);
  if (typeof cleanedValue === "number") return cleanedValue;
  const text = String(cleanedValue).trim();
  if (text === "") return NaN;
  const cleaned = text.replace(/[^\d.-]/g, "");
  if (cleaned === "" || cleaned === "-" || cleaned === "." || cleaned === "-.") return NaN;
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : NaN;
}

function suggestBrowserRowBounds(worksheet) {
  if (worksheetRowBoundsCache.has(worksheet)) return worksheetRowBoundsCache.get(worksheet);
  let first = null;
  let last = null;
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (isHiddenRow(worksheet, rowNumber)) return;
    let count = 0;
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      if (isHiddenColumn(worksheet, colNumber - 1)) return;
      if (cleanExcelCellValue(cell.value) !== "") count += 1;
    });
    if (count >= 2) {
      if (first === null) first = rowNumber;
      last = rowNumber;
    }
  });
  const bounds = [first ?? 1, last ?? first ?? 1];
  worksheetRowBoundsCache.set(worksheet, bounds);
  return bounds;
}

function normalizeBrowserFillColor(cell) {
  const fill = cell?.fill ?? cell?.style?.fill;
  if (!fill || fill.type === "pattern" && fill.pattern === "none") return null;
  const candidates = [fill.fgColor, fill.bgColor].filter(Boolean);
  for (const candidate of candidates) {
    const rgb = normalizeExcelColor(candidate);
    if (rgb) return rgb;
  }
  return null;
}

function getEffectiveBrowserFillColor(worksheet, rowNumber, columnIndex) {
  const conditionalColor = inferConditionalFormattingFillColor(worksheet, rowNumber, columnIndex)
    || inferStatusConditionalFillColor(worksheet, rowNumber);
  if (conditionalColor) return conditionalColor;
  return normalizeBrowserFillColor(getDisplayCell(worksheet, rowNumber, columnIndex));
}

function inferConditionalFormattingFillColor(worksheet, rowNumber, columnIndex) {
  const colNumber = columnIndex + 1;
  for (const item of getWorksheetConditionalFillRules(worksheet)) {
    const matchedRange = item.ranges.find((range) => isCellInRange(rowNumber, colNumber, range));
    if (!matchedRange) continue;
    if (evaluateConditionalFillRule(worksheet, item.rule, rowNumber, columnIndex, item.anchor)) return item.color;
  }
  return null;
}

function getWorksheetConditionalFillRules(worksheet) {
  if (worksheetConditionalFillCache.has(worksheet)) return worksheetConditionalFillCache.get(worksheet);
  const groups = worksheet?.conditionalFormattings ?? worksheet?.model?.conditionalFormattings ?? [];
  const items = [];
  groups.forEach((group) => {
    const ranges = parseRangeList(group.ref).filter(Boolean);
    if (!ranges.length) return;
    const anchor = ranges[0];
    (group.rules ?? []).forEach((rule) => {
      const color = normalizeExcelColor(rule?.style?.fill?.fgColor) || normalizeExcelColor(rule?.style?.fill?.bgColor);
      if (!color) return;
      items.push({ ranges, anchor, rule, color, priority: Number(rule.priority ?? 9999) });
    });
  });
  items.sort((a, b) => a.priority - b.priority);
  worksheetConditionalFillCache.set(worksheet, items);
  return items;
}

function parseRangeList(ref) {
  return String(ref || "").split(/\s+/).map(parseMergeRange).filter(Boolean);
}

function isCellInRange(rowNumber, colNumber, range) {
  return rowNumber >= range.top && rowNumber <= range.bottom && colNumber >= range.left && colNumber <= range.right;
}

function evaluateConditionalFillRule(worksheet, rule, rowNumber, columnIndex, anchor) {
  if (rule.type === "expression") {
    return (rule.formulae ?? []).some((formula) => evaluateConditionalExpression(worksheet, formula, rowNumber, columnIndex, anchor));
  }
  if (rule.type === "cellIs") return evaluateCellIsRule(worksheet, rule, rowNumber, columnIndex);
  return false;
}

function evaluateConditionalExpression(worksheet, formula, rowNumber, columnIndex, anchor) {
  const text = String(formula || "").trim();
  const orArgs = getFunctionArguments(text, "OR");
  if (orArgs) return orArgs.some((part) => evaluateConditionalExpression(worksheet, part, rowNumber, columnIndex, anchor));
  const andArgs = getFunctionArguments(text, "AND");
  if (andArgs) return andArgs.every((part) => evaluateConditionalExpression(worksheet, part, rowNumber, columnIndex, anchor));
  return evaluateConditionalTerm(worksheet, text, rowNumber, columnIndex, anchor);
}

function getFunctionArguments(text, name) {
  const pattern = new RegExp(`^${name}\s*\((.*)\)$`, "i");
  const match = text.match(pattern);
  return match ? splitFormulaArguments(match[1]) : null;
}

function splitFormulaArguments(text) {
  const result = [];
  let depth = 0;
  let quote = false;
  let start = 0;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    if (char === '"') quote = !quote;
    else if (!quote && char === "(") depth += 1;
    else if (!quote && char === ")") depth -= 1;
    else if (!quote && depth === 0 && char === ",") {
      result.push(text.slice(start, index).trim());
      start = index + 1;
    }
  }
  result.push(text.slice(start).trim());
  return result.filter(Boolean);
}

function evaluateConditionalTerm(worksheet, term, rowNumber, columnIndex, anchor) {
  const searchMatch = term.match(/^ISNUMBER\s*\(\s*SEARCH\s*\(\s*"([^"]*)"\s*,\s*(\$?[A-Z]+\$?\d+)\s*\)\s*\)$/i);
  if (searchMatch) {
    const value = getFormulaReferenceText(worksheet, searchMatch[2], rowNumber, columnIndex, anchor);
    return value.includes(searchMatch[1]);
  }
  const equalMatch = term.match(/^(\$?[A-Z]+\$?\d+)\s*=\s*"([^"]*)"$/i);
  if (equalMatch) return getFormulaReferenceText(worksheet, equalMatch[1], rowNumber, columnIndex, anchor) === equalMatch[2];
  const notEqualMatch = term.match(/^(\$?[A-Z]+\$?\d+)\s*<>\s*"([^"]*)"$/i);
  if (notEqualMatch) return getFormulaReferenceText(worksheet, notEqualMatch[1], rowNumber, columnIndex, anchor) !== notEqualMatch[2];
  return false;
}

function getFormulaReferenceText(worksheet, reference, rowNumber, columnIndex, anchor) {
  const cellRef = resolveFormulaCellReference(reference, rowNumber, columnIndex, anchor);
  if (!cellRef) return "";
  return String(cleanExcelCellValue(getDisplayCell(worksheet, cellRef.row, cellRef.col - 1).value) ?? "").trim();
}

function resolveFormulaCellReference(reference, rowNumber, columnIndex, anchor) {
  const match = String(reference).match(/^(\$?)([A-Z]+)(\$?)(\d+)$/i);
  if (!match) return null;
  const formulaCol = columnNameToNumber(match[2]);
  const formulaRow = Number(match[4]);
  const col = match[1] ? formulaCol : formulaCol + ((columnIndex + 1) - anchor.left);
  const row = match[3] ? formulaRow : formulaRow + (rowNumber - anchor.top);
  return { row, col };
}

function evaluateCellIsRule(worksheet, rule, rowNumber, columnIndex) {
  const cellValue = parseConditionalNumber(cleanExcelCellValue(getDisplayCell(worksheet, rowNumber, columnIndex).value));
  const targetValue = parseConditionalNumber(rule.formulae?.[0]);
  if (!Number.isFinite(cellValue) || !Number.isFinite(targetValue)) return false;
  if (rule.operator === "greaterThan") return cellValue > targetValue;
  if (rule.operator === "greaterThanOrEqual") return cellValue >= targetValue;
  if (rule.operator === "lessThan") return cellValue < targetValue;
  if (rule.operator === "lessThanOrEqual") return cellValue <= targetValue;
  if (rule.operator === "equal") return cellValue === targetValue;
  if (rule.operator === "notEqual") return cellValue !== targetValue;
  return false;
}

function parseConditionalNumber(value) {
  const text = String(value ?? "").trim().replace(/^"|"$/g, "");
  if (!text) return NaN;
  const isPercent = text.endsWith("%");
  const parsed = parseBrowserNumber(text.replace(/%$/, ""));
  return Number.isFinite(parsed) ? (isPercent ? parsed / 100 : parsed) : NaN;
}

function inferStatusConditionalFillColor(worksheet, rowNumber) {
  if (!hasStatusConditionalFillHint(worksheet)) return null;
  const productText = String(cleanExcelCellValue(getDisplayCell(worksheet, rowNumber, 0).value) ?? "").trim();
  const statusText = String(cleanExcelCellValue(getDisplayCell(worksheet, rowNumber, 4).value) ?? "").trim().toUpperCase();
  if (!productText) return null;
  if (productText.includes("(\uD488\uC808)") || productText.includes("(\uB2E8\uC885)") || statusText === "X") return CONDITIONAL_STATUS_RED;
  return null;
}

function hasStatusConditionalFillHint(worksheet) {
  if (worksheetStatusFillCache.has(worksheet)) return worksheetStatusFillCache.get(worksheet);
  let hasHint = false;
  for (let rowNumber = 1; rowNumber <= Math.min(worksheet.rowCount, 500); rowNumber += 1) {
    const text = String(cleanExcelCellValue(getDisplayCell(worksheet, rowNumber, 0).value) ?? "");
    if (text.includes("(\uD488\uC808)") || text.includes("(\uB2E8\uC885)")) {
      hasHint = true;
      break;
    }
  }
  worksheetStatusFillCache.set(worksheet, hasHint);
  return hasHint;
}

function normalizeExcelColor(color) {
  if (!color) return null;
  if (color.argb) {
    const value = String(color.argb).slice(-6).toUpperCase();
    if (value === "000000" && /^0{8}$/i.test(String(color.argb))) return null;
    return value;
  }
  if (color.rgb) return String(color.rgb).slice(-6).toUpperCase();
  if (Number.isInteger(color.indexed)) return indexedColorToRgb(color.indexed);
  if (Number.isInteger(color.theme)) return applyTint(themeColorToRgb(color.theme), color.tint ?? 0);
  return null;
}

function indexedColorToRgb(index) {
  const table = {
    0: "000000", 1: "FFFFFF", 2: "FF0000", 3: "00FF00", 4: "0000FF", 5: "FFFF00", 6: "FF00FF", 7: "00FFFF",
    8: "000000", 9: "FFFFFF", 10: "FF0000", 11: "00FF00", 12: "0000FF", 13: "FFFF00", 14: "FF00FF", 15: "00FFFF",
    16: "800000", 17: "008000", 18: "000080", 19: "808000", 20: "800080", 21: "008080", 22: "C0C0C0", 23: "808080",
    24: "9999FF", 25: "993366", 26: "FFFFCC", 27: "CCFFFF", 28: "660066", 29: "FF8080", 30: "0066CC", 31: "CCCCFF",
    32: "000080", 33: "FF00FF", 34: "FFFF00", 35: "00FFFF", 36: "800080", 37: "800000", 38: "008080", 39: "0000FF",
    40: "00CCFF", 41: "CCFFFF", 42: "CCFFCC", 43: "FFFF99", 44: "99CCFF", 45: "FF99CC", 46: "CC99FF", 47: "FFCC99",
    48: "3366FF", 49: "33CCCC", 50: "99CC00", 51: "FFCC00", 52: "FF9900", 53: "FF6600", 54: "666699", 55: "969696",
    56: "003366", 57: "339966", 58: "003300", 59: "333300", 60: "993300", 61: "993366", 62: "333399", 63: "333333",
  };
  return table[index] ?? null;
}

function themeColorToRgb(theme) {
  const themeColors = ["FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646"];
  return themeColors[theme] ?? null;
}

function applyTint(rgb, tint) {
  if (!rgb || !tint) return rgb;
  const channels = [0, 2, 4].map((index) => parseInt(rgb.slice(index, index + 2), 16));
  const tinted = channels.map((channel) => {
    const next = tint < 0 ? channel * (1 + tint) : channel + (255 - channel) * tint;
    return Math.max(0, Math.min(255, Math.round(next))).toString(16).padStart(2, "0").toUpperCase();
  });
  return tinted.join("");
}

async function columnColorsInBrowser(fields) {
  await getExcelWorkbook();
  const worksheet = getWorksheetByName(fields.sheet_name);
  const columnIndex = Number(fields.column_index);
  if (isHiddenColumn(worksheet, columnIndex)) return { colors: [] };
  const counts = new Map();
  const [, effectiveLastRow] = suggestBrowserRowBounds(worksheet);
  const start = Math.max(Number(fields.start_row || 1), 1);
  const end = Math.min(Number(fields.end_row || effectiveLastRow), effectiveLastRow, worksheet.rowCount);
  for (let rowNumber = start; rowNumber <= end; rowNumber += 1) {
    if (isHiddenRow(worksheet, rowNumber)) continue;
    const cell = getDisplayCell(worksheet, rowNumber, columnIndex);
    if (cleanExcelCellValue(cell.value) === "") continue;
    const color = getEffectiveBrowserFillColor(worksheet, rowNumber, columnIndex);
    if (color) counts.set(color, (counts.get(color) ?? 0) + 1);
  }
  return {
    colors: [...counts.entries()].map(([value, count]) => ({ value, label: `#${value}`, count })),
  };
}

async function calculateExcelInBrowser(fields) {
  await getExcelWorkbook();
  const worksheet = getWorksheetByName(fields.sheet_name);
  const productCol = Number(fields.product_col);
  const originalCol = Number(fields.original_col);
  const finalCol = Number(fields.final_price_col);
  const categoryCol = fields.category_col === "" ? null : Number(fields.category_col);
  const extraFieldList = JSON.parse(fields.extra_fields || "[]");
  const rewardFieldList = JSON.parse(fields.reward_fields || "[]");

  ensureVisibleBrowserColumn(worksheet, productCol, "제품명");
  ensureVisibleBrowserColumn(worksheet, originalCol, "기존금액");
  ensureVisibleBrowserColumn(worksheet, finalCol, "최종가");
  if (categoryCol !== null) ensureVisibleBrowserColumn(worksheet, categoryCol, "카테고리");
  extraFieldList.forEach((field) => {
    ["source_col", "operand_col", "left_col", "right_col"].forEach((key) => {
      if (Number.isInteger(field[key])) ensureVisibleBrowserColumn(worksheet, field[key], "추가 항목");
    });
  });
  rewardFieldList.forEach((field) => {
    if (Number.isInteger(field.source_col)) ensureVisibleBrowserColumn(worksheet, field.source_col, "적립금액");
  });

  const rows = [];
  let activeCategoryName = "";
  const [, effectiveLastRow] = suggestBrowserRowBounds(worksheet);
  const start = Math.max(Number(fields.start_row || 1), 1);
  const end = Math.min(Number(fields.end_row || effectiveLastRow), effectiveLastRow, worksheet.rowCount);
  for (let rowNumber = start; rowNumber <= end; rowNumber += 1) {
    if (isHiddenRow(worksheet, rowNumber)) continue;
    const row = worksheet.getRow(rowNumber);
    const categoryValue = categoryCol !== null ? String(cleanExcelCellValue(getDisplayCell(worksheet, rowNumber, categoryCol).value)).trim() : "";
    if (categoryValue) activeCategoryName = categoryValue;
    if (!passesColorFilter(worksheet, rowNumber, productCol, fields.product_color, fields.product_color_mode)) continue;
    if (!passesColorFilter(worksheet, rowNumber, originalCol, fields.original_color, fields.original_color_mode)) continue;
    if (!passesColorFilter(worksheet, rowNumber, finalCol, fields.final_price_color, fields.final_price_color_mode)) continue;

    const productValue = cleanExcelCellValue(getDisplayCell(worksheet, rowNumber, productCol).value);
    const originalCellValue = cleanExcelCellValue(getDisplayCell(worksheet, rowNumber, originalCol).value);
    const finalCellValue = cleanExcelCellValue(getDisplayCell(worksheet, rowNumber, finalCol).value);
    if (String(productValue).trim() === "" || String(originalCellValue).trim() === "" || String(finalCellValue).trim() === "") continue;
    const originalPrice = parseBrowserNumber(originalCellValue);
    const finalPrice = parseBrowserNumber(finalCellValue);
    if (!Number.isFinite(originalPrice) || !Number.isFinite(finalPrice) || originalPrice <= 0) continue;

    const extraValues = {};
    extraFieldList.forEach((field) => {
      if (field.mode === "raw") {
        extraValues[field.id] = cleanExcelCellValue(getDisplayCell(worksheet, rowNumber, field.source_col).value);
      } else if (field.mode === "original_minus") {
        const value = parseBrowserNumber(getDisplayCell(worksheet, rowNumber, field.operand_col).value);
        extraValues[field.id] = Number.isFinite(value) ? Math.round(originalPrice - value) : "";
      } else if (field.mode === "column_minus") {
        const left = parseBrowserNumber(getDisplayCell(worksheet, rowNumber, field.left_col).value);
        const right = parseBrowserNumber(getDisplayCell(worksheet, rowNumber, field.right_col).value);
        extraValues[field.id] = Number.isFinite(left) && Number.isFinite(right) ? Math.round(left - right) : "";
      }
    });

    const hasExtraPriceFormula = extraFieldList.length > 0;
    const calculatedFinalPrice = hasExtraPriceFormula
      ? calculateFinalPriceFromExtraFields(originalPrice, extraValues, extraFieldList)
      : Math.round(finalPrice);

    const rewardValues = {};
    let totalRewardAmount = 0;
    rewardFieldList.forEach((field) => {
      const value = parseBrowserNumber(getDisplayCell(worksheet, rowNumber, field.source_col).value);
      const amount = Number.isFinite(value) ? Math.round(value) : 0;
      rewardValues[field.id] = amount;
      totalRewardAmount += amount;
    });

    const totalDiscount = Math.round(originalPrice - calculatedFinalPrice);
    if (totalDiscount === 0) continue;
    rows.push({
      row_id: rowNumber,
      product_name: String(productValue),
      category_name: categoryCol !== null ? (categoryValue || activeCategoryName) : "",
      base_original_price: Math.round(originalPrice),
      has_extra_price_formula: hasExtraPriceFormula,
      original_price: Math.round(originalPrice),
      discount_amount: calculatedFinalPrice,
      total_discount_amount: totalDiscount,
      discount_rate: originalPrice > 0 ? Math.round((totalDiscount / originalPrice) * 1000) / 10 : 0,
      extra_values: extraValues,
      reward_values: rewardValues,
      total_reward_amount: totalRewardAmount,
      effective_price: Math.round(calculatedFinalPrice - totalRewardAmount),
    });
  }
  if (rows.length === 0) throw new Error("선택한 범위에서 출력할 상품을 찾지 못했습니다.");
  return { rows, extra_fields: extraFieldList, reward_fields: rewardFieldList };
}

function passesColorFilter(worksheet, rowNumber, columnIndex, selectedColor, mode = "include") {
  if (!selectedColor) return true;
  const cellColor = getEffectiveBrowserFillColor(worksheet, rowNumber, columnIndex);
  return mode === "exclude" ? cellColor !== selectedColor : cellColor === selectedColor;
}

function columnLabel(index) {
  let n = index + 1;
  let label = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
}

function groupSimilarResultRows(rows) {
  const usedRowIds = new Set();
  const result = [];
  rows.forEach((row, index) => {
    if (usedRowIds.has(row.row_id)) return;
    const group = [row];
    for (let nextIndex = index + 1; nextIndex < rows.length; nextIndex += 1) {
      const candidate = rows[nextIndex];
      if (usedRowIds.has(candidate.row_id)) continue;
      if (!canGroupSimilarRows(row, candidate)) continue;
      group.push(candidate);
      usedRowIds.add(candidate.row_id);
    }
    if (group.length > 1) {
      result.push({
        ...row,
        product_name: buildGroupedProductName(group.map((item) => displayProductName(item.product_name ?? ""))),
        grouped_row_ids: group.map((item) => item.row_id),
      });
    } else {
      result.push(row);
    }
    usedRowIds.add(row.row_id);
  });
  return result;
}

function canGroupSimilarRows(rowA, rowB) {
  if ((rowA.category_name || "") !== (rowB.category_name || "")) return false;
  if (!hasSameVisibleResultValues(rowA, rowB)) return false;
  return Boolean(buildGroupedProductName([displayProductName(rowA.product_name ?? ""), displayProductName(rowB.product_name ?? "")]));
}

function hasSameVisibleResultValues(rowA, rowB) {
  const columns = ["original_price", "discount_amount", "total_discount_amount", "discount_rate"];
  if (rewardFields.length > 0) columns.push("total_reward_amount", "effective_price");
  const baseSame = columns.every((column) => normalizeSignatureValue(rowA[column]) === normalizeSignatureValue(rowB[column]));
  if (!baseSame) return false;
  const extraSame = extraFields.every((field) => normalizeSignatureValue(rowA.extra_values?.[field.id]) === normalizeSignatureValue(rowB.extra_values?.[field.id]));
  if (!extraSame) return false;
  const rewardSame = rewardFields.every((field) => normalizeSignatureValue(rowA.reward_values?.[field.id]) === normalizeSignatureValue(rowB.reward_values?.[field.id]));
  return rewardSame;
}

function buildResultSignature(row) {
  const extraSignature = extraFields.map((field) => normalizeSignatureValue(row.extra_values?.[field.id])).join("|");
  const rewardSignature = rewardFields.map((field) => normalizeSignatureValue(row.reward_values?.[field.id])).join("|");
  return [
    normalizeSignatureValue(row.original_price),
    normalizeSignatureValue(row.discount_amount),
    normalizeSignatureValue(row.total_discount_amount),
    normalizeSignatureValue(row.discount_rate),
    extraSignature,
    rewardSignature,
    normalizeSignatureValue(row.total_reward_amount),
    normalizeSignatureValue(row.effective_price),
  ].join("::");
}

function normalizeSignatureValue(value) {
  if (value === null || value === undefined || value === "") return "0";
  const cleaned = String(value).replace(/[^\d.-]/g, "");
  const numeric = Number(cleaned || value);
  return Number.isFinite(numeric) ? String(Math.round(numeric * 1000) / 1000) : String(value).trim();
}

function buildGroupedProductName(names) {
  const cleanedNames = names.map((name) => String(name ?? "").replace(/\s+/g, " ").trim()).filter(Boolean);
  if (cleanedNames.length < 2) return "";

  const directName = buildGroupedNameFromCleanedNames(cleanedNames);
  if (directName) return directName;

  const brandlessNames = removeSharedLeadingBrand(cleanedNames);
  if (brandlessNames.join("||") !== cleanedNames.join("||")) {
    const brandlessName = buildGroupedNameFromCleanedNames(brandlessNames);
    if (brandlessName) {
      const brand = cleanedNames[0].split(" ")[0];
      return `${brand} ${brandlessName}`.trim();
    }
  }
  return "";
}

function buildGroupedNameFromCleanedNames(cleanedNames) {
  const firstName = cleanedNames[0];
  const commonPrefix = getCommonWordPrefix(cleanedNames);
  if (!commonPrefix) return "";
  const commonWords = commonPrefix.trim().split(/\s+/).filter(Boolean);
  if (commonWords.length < 2) return "";
  const rawSuffixes = cleanedNames.map((name) => name.slice(commonPrefix.length).trim());
  const nonEmptySuffixes = rawSuffixes.filter(Boolean);
  if (nonEmptySuffixes.length === 0) return "";
  const maxSuffixLength = Math.max(...nonEmptySuffixes.map((suffix) => suffix.length));
  const commonLength = commonPrefix.trim().length;
  if (commonLength < 6 || maxSuffixLength > commonLength * 1.2) return "";
  if (rawSuffixes.some((suffix) => !suffix)) {
    return `${commonPrefix.trim()} ${GROUPED_PRODUCT_SEPARATOR} ${nonEmptySuffixes.join(GROUPED_PRODUCT_JOINER)}`;
  }
  return `${firstName} ${GROUPED_PRODUCT_SEPARATOR} ${rawSuffixes.slice(1).join(GROUPED_PRODUCT_JOINER)}`;
}

function removeSharedLeadingBrand(names) {
  const firstWords = names[0].split(" ");
  if (firstWords.length < 2) return names;
  const brand = firstWords[0].toLowerCase();
  if (!names.every((name) => name.split(" ")[0]?.toLowerCase() === brand)) return names;
  return names.map((name) => name.split(" ").slice(1).join(" ").trim()).filter(Boolean);
}

function getCommonWordPrefix(names) {
  const firstWords = names[0].split(" ");
  let commonWords = [...firstWords];
  names.slice(1).forEach((name) => {
    const words = name.split(" ");
    let count = 0;
    while (count < commonWords.length && count < words.length && commonWords[count].toLowerCase() === words[count].toLowerCase()) {
      count += 1;
    }
    commonWords = commonWords.slice(0, count);
  });
  return commonWords.join(" ");
}

function groupRowsByCategory(rows) {
  const groups = new Map();
  rows.forEach((row) => {
    const key = row.category_name || "미분류";
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  });
  return [...groups.entries()];
}

function clearRestoredResultSnapshot() {
  activeResultSnapshotConfig = null;
  resultSnapshotDirty = true;
}

function syncDynamicFieldsFromDom() {
  extraFieldsList.querySelectorAll(".extra-field-card").forEach((card) => {
    const field = extraFields.find((item) => item.id === card.dataset.fieldId);
    if (field) updateExtraFieldFromCard(field, card);
  });
  rewardFieldsList.querySelectorAll(".reward-field-card").forEach((card) => {
    const field = rewardFields.find((item) => item.id === card.dataset.fieldId);
    if (field) updateRewardFieldFromCard(field, card);
  });
}

function syncTypedColumns() {
  [productColumn, originalColumn, finalPriceColumn, categoryColumn].forEach((input) => {
    const parsed = parseColumnInput(input.value);
    if (parsed === null) {
      delete input.dataset.index;
    } else {
      input.dataset.index = String(parsed);
    }
  });
}

function parseColumnInput(value) {
  const normalized = String(value).trim().toUpperCase().replace(/열$/, "");
  if (!/^[A-Z]+$/.test(normalized)) return null;
  let result = 0;
  for (const char of normalized) {
    result = result * 26 + (char.charCodeAt(0) - 64);
  }
  return result - 1;
}

function formatCurrency(value) {
  return `${currencyFormatter.format(value)}${showWonSuffix.checked ? "원" : ""}`;
}

function formatTotalDiscount(value) {
  const prefix = showDiscountMinus.checked ? "-" : "";
  return `${prefix}${formatCurrency(value)}`;
}

function formatRate(value) {
  return `${Math.round(value)}%`;
}

function formatExtraValue(value) {
  const normalized = normalizeDisplayValue(value);
  if (normalized === null || normalized === undefined || normalized === "") return "";
  return typeof normalized === "number" ? formatCurrency(normalized) : String(normalized);
}

function normalizeDisplayValue(value) {
  const cleaned = cleanExcelCellValue(value);
  if (cleaned === null || cleaned === undefined) return "";
  return cleaned;
}

function displayProductName(value) {
  let baseValue = String(value);

  if (removeParenthesesText.checked) {
    baseValue = baseValue.replace(/\([^)]*\)/g, "");
  }
  if (removeBracketsText.checked) {
    baseValue = baseValue.replace(/\[[^\]]*\]/g, "");
  }
  if (removeLeadingText.checked && leadingTextValue.value.trim()) {
    const escapedLeadingText = escapeRegExp(leadingTextValue.value.trim());
    baseValue = baseValue.replace(new RegExp(`^${escapedLeadingText}\\s*`, "i"), "");
  }
  if (delimiterValue.value && (removeAfterDelimiter.checked || removeBeforeDelimiter.checked)) {
    const delimiterIndex = baseValue.indexOf(delimiterValue.value);
    if (delimiterIndex >= 0) {
      baseValue = removeBeforeDelimiter.checked
        ? baseValue.slice(delimiterIndex + delimiterValue.value.length)
        : baseValue.slice(0, delimiterIndex);
    }
  }

  return productNameFilters
    .filter((text) => text.trim())
    .reduce((current, text) => current.replaceAll(text.trim(), ""), baseValue)
    .replace(/\s+/g, " ")
    .trim()
    .split(" ")
    .map((word) => (titleCaseProductName.checked ? toTitleCaseWord(word) : word))
    .join(" ");
}

function toTitleCaseWord(word) {
  if (!word) return word;
  return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function applyVisibleColumns() {
  productNameHeader.classList.toggle("hidden-result-column", hiddenColumns.has("product_name"));
  resultBody.querySelectorAll("[data-column]").forEach((cell) => {
    cell.classList.toggle("hidden-result-column", hiddenColumns.has(cell.dataset.column));
  });
  document.querySelectorAll("th[data-column]").forEach((header) => {
    header.classList.toggle("hidden-result-column", hiddenColumns.has(header.dataset.column));
  });
  resultBody.querySelectorAll("tr").forEach((row) => {
    const firstCell = row.children[0];
    if (row.classList.contains("category-group-row")) {
      firstCell?.setAttribute("colspan", String(getResultColumnSpan()));
    } else if (firstCell) {
      firstCell.classList.toggle("hidden-result-column", hiddenColumns.has("product_name"));
    }
  });
  applyRewardColumnVisibility();
  refreshResultColumnLayout();
}

function refreshResultColumnLayout() {
  applyProductColumnWidth();
  updateFloatingResultHeader();
  requestAnimationFrame(() => {
    applyProductColumnWidth();
    updateFloatingResultHeader();
  });
}

function applyPriceHighlighting() {
  const shouldHighlight = highlightFinalPrices.checked;
  document.querySelectorAll('[data-column="discount_amount"]').forEach((node) => {
    node.classList.toggle("highlight-price-column", shouldHighlight);
  });
  document.querySelectorAll('[data-column="effective_price"]').forEach((node) => {
    node.classList.toggle("highlight-effective-price-column", shouldHighlight && rewardFields.length > 0);
  });
}

function addExtraField() {
  clearRestoredResultSnapshot();
  extraFields.push({
    id: crypto.randomUUID(),
    name: "",
    name_choice: "",
    custom_name: "",
    mode: "raw",
    source_col: "",
    operand_col: "",
    left_col: "",
    right_col: "",
  });
  renderExtraFieldInputs();
  queueSaveCurrentSettings();
}

function addRewardField() {
  clearRestoredResultSnapshot();
  rewardFields.push({
    id: crypto.randomUUID(),
    name: "",
    name_choice: "",
    custom_name: "",
    source_col: "",
  });
  renderRewardFieldInputs();
  queueSaveCurrentSettings();
}

function renderExtraFieldInputs() {
  extraFieldsList.innerHTML = "";
  extraFields.forEach((field) => {
    field.name_choice ??= extraFieldNameSuggestions.includes(field.name) ? field.name : "custom";
    field.custom_name ??= field.name_choice === "custom" ? field.name : "";
    const nameSelectId = `extraNameSelect_${field.id}`;
    const modeSelectId = `extraModeSelect_${field.id}`;
    const card = document.createElement("div");
    card.className = "extra-field-card";
    card.dataset.fieldId = field.id;
    card.dataset.mode = field.mode || "raw";
    card.innerHTML = `
      <label>
        <span>항목명</span>
        <div class="custom-select extra-name-select" data-select-target="${nameSelectId}">
          <button class="custom-select-trigger" type="button"></button>
          <div class="custom-select-menu hidden"></div>
        </div>
        <select id="${nameSelectId}" data-role="name_choice" class="native-select-hidden" aria-hidden="true" tabindex="-1">
          <option value="">항목 선택</option>
          <option value="선택쿠폰">선택쿠폰</option>
          <option value="중복쿠폰">중복쿠폰</option>
          <option value="카드할인">카드할인</option>
          <option value="custom">직접입력</option>
        </select>
        <div class="extra-custom-name-wrap hidden-field">
          <input class="extra-custom-name" data-role="custom_name" value="${escapeAttribute(field.custom_name)}" placeholder="항목명을 입력하세요" />
          <button class="extra-custom-select-toggle" type="button" aria-label="항목명 선택 열기"></button>
          <div class="extra-custom-select-menu hidden">
            <button type="button" data-name-choice="">항목 선택</button>
            <button type="button" data-name-choice="선택쿠폰">선택쿠폰</button>
            <button type="button" data-name-choice="중복쿠폰">중복쿠폰</button>
            <button type="button" data-name-choice="카드할인">카드할인</button>
          </div>
        </div>
      </label>
      <label>
        <span>출력 방식</span>
        <div class="custom-select" data-select-target="${modeSelectId}">
          <button class="custom-select-trigger" type="button"></button>
          <div class="custom-select-menu hidden"></div>
        </div>
        <select id="${modeSelectId}" data-role="mode" class="native-select-hidden" aria-hidden="true" tabindex="-1">
          <option value="raw">그대로 출력</option>
          <option value="original_minus">기존금액 - 선택 열</option>
          <option value="column_minus">두 열 차이</option>
        </select>
      </label>
      <label data-field="source_col">
        <span>출력 열</span>
        <input data-role="source_col" placeholder="예: H" value="${columnValue(field.source_col)}" />
      </label>
      <label data-field="operand_col">
        <span>뺄 열</span>
        <input data-role="operand_col" placeholder="예: H" value="${columnValue(field.operand_col)}" />
      </label>
      <label data-field="left_col">
        <span>왼쪽 열</span>
        <input data-role="left_col" placeholder="예: F" value="${columnValue(field.left_col)}" />
      </label>
      <label data-field="right_col">
        <span>오른쪽 열</span>
        <input data-role="right_col" placeholder="예: H" value="${columnValue(field.right_col)}" />
      </label>
      <button class="ghost-button extra-field-remove" type="button">삭제</button>
    `;
    const nameSelect = card.querySelector('[data-role="name_choice"]');
    const customNameInput = card.querySelector('[data-role="custom_name"]');
    const modeSelect = card.querySelector('[data-role="mode"]');
    nameSelect.value = field.name_choice;
    modeSelect.value = field.mode;
    extraFieldsList.appendChild(card);
    mountCustomSelect(card.querySelector(`[data-select-target="${nameSelectId}"]`), nameSelect);
    mountCustomSelect(card.querySelector(`[data-select-target="${modeSelectId}"]`), modeSelect);
    card.querySelectorAll("input").forEach((input) => {
      input.addEventListener("focus", () => {
        if (input.dataset.role === "custom_name") return;
        activeExtraFieldTarget = input;
        setActivePick("extra");
      });
      input.addEventListener("input", () => {
        updateExtraFieldFromCard(field, card);
        persistDynamicColumnInput(input);
        clearRestoredResultSnapshot();
        queueSaveCurrentSettings();
      });
    });
    nameSelect.addEventListener("change", () => {
      updateExtraFieldFromCard(field, card);
      updateExtraNameVisibility(card, field.name_choice);
      activateDynamicFieldColumnTarget(card);
      clearRestoredResultSnapshot();
      queueSaveCurrentSettings();
    });
    const customToggle = card.querySelector(".extra-custom-select-toggle");
    const customMenu = card.querySelector(".extra-custom-select-menu");
    customToggle.addEventListener("click", (event) => {
      event.stopPropagation();
      customMenu.classList.toggle("hidden");
    });
    customMenu.querySelectorAll("button").forEach((button) => {
      button.addEventListener("click", () => {
        field.name_choice = button.dataset.nameChoice;
        nameSelect.value = field.name_choice;
        updateExtraFieldFromCard(field, card);
        updateExtraNameVisibility(card, field.name_choice);
        syncCustomSelect(nameSelect);
        activateDynamicFieldColumnTarget(card);
        customMenu.classList.add("hidden");
        clearRestoredResultSnapshot();
        queueSaveCurrentSettings();
      });
    });
    modeSelect.addEventListener("change", () => {
      updateExtraFieldFromCard(field, card);
      updateExtraFieldVisibility(card, field.mode);
      activateDynamicFieldColumnTarget(card);
      clearRestoredResultSnapshot();
      queueSaveCurrentSettings();
    });
    card.querySelector(".extra-field-remove").addEventListener("click", () => {
      extraFields = extraFields.filter((item) => item.id !== field.id);
      clearRestoredResultSnapshot();
      renderExtraFieldInputs();
      queueSaveCurrentSettings();
    });
    updateExtraNameVisibility(card, field.name_choice);
    updateExtraFieldVisibility(card, field.mode);
  });
}

function activateDynamicFieldColumnTarget(card) {
  const mode = card.querySelector('[data-role="mode"]')?.value ?? "raw";
  const role = mode === "raw" ? "source_col" : mode === "original_minus" ? "operand_col" : "left_col";
  const input = card.querySelector(`[data-role="${role}"]`) || card.querySelector('[data-role="source_col"]');
  if (!input) return;
  activeExtraFieldTarget = input;
  setActivePick("extra");
}

function updateExtraFieldFromCard(field, card) {
  field.name_choice = card.querySelector('[data-role="name_choice"]').value;
  field.custom_name = card.querySelector('[data-role="custom_name"]').value;
  field.name = field.name_choice === "custom" ? field.custom_name || "추가 항목" : field.name_choice || "추가 항목";
  field.mode = card.querySelector('[data-role="mode"]').value;
  ["source_col", "operand_col", "left_col", "right_col"].forEach((key) => {
    const input = card.querySelector(`[data-role="${key}"]`);
    const parsed = normalizeColumnIndexValue(input.value);
    field[key] = parsed === null ? "" : parsed;
  });
}

function updateExtraNameVisibility(card, nameChoice) {
  card.querySelector(".extra-name-select").classList.toggle("hidden-field", nameChoice === "custom");
  card.querySelector(".extra-custom-name-wrap").classList.toggle("hidden-field", nameChoice !== "custom");
}

function updateExtraFieldVisibility(card, mode) {
  card.dataset.mode = mode || "raw";
  card.querySelector('[data-field="source_col"]').classList.toggle("hidden-field", mode !== "raw");
  card.querySelector('[data-field="operand_col"]').classList.toggle("hidden-field", mode !== "original_minus");
  card.querySelector('[data-field="left_col"]').classList.toggle("hidden-field", mode !== "column_minus");
  card.querySelector('[data-field="right_col"]').classList.toggle("hidden-field", mode !== "column_minus");
}

function serializeExtraFields() {
  return extraFields.map((field) => ({ ...field }));
}

function validateExtraFields() {
  for (const field of extraFields) {
    const name = field.name || "추가 항목";
    if (field.mode === "raw" && field.source_col === "") {
      return `"${name}"의 출력 열을 선택해주세요.`;
    }
    if (field.mode === "original_minus" && field.operand_col === "") {
      return `"${name}"의 뺄 열을 선택해주세요.`;
    }
    if (field.mode === "column_minus" && (field.left_col === "" || field.right_col === "")) {
      return `"${name}"의 왼쪽 열과 오른쪽 열을 모두 선택해주세요.`;
    }
  }
  for (const field of rewardFields) {
    const name = field.name || "적립금액";
    if (field.source_col === "") {
      return `"${name}"의 적립금액 열을 선택해주세요.`;
    }
  }
  return "";
}

function serializeRewardFields() {
  return rewardFields.map((field) => ({ ...field }));
}

function isLegacyXlsFile(file) {
  const name = String(file?.name ?? "").toLowerCase();
  return name.endsWith(".xls") && !name.endsWith(".xlsx");
}

function createFileSignature(file) {
  return `${file.name}_${file.size}`;
}

function createFileSignatureStorageKey(file) {
  return `signature:${createFileSignature(file)}`;
}

function createFileNameStorageKey(fileName) {
  return `name:${fileName}`;
}

function getSettingsStorageKeys(fileKey, data) {
  const keys = [fileKey];
  if (data?.fileSignature) keys.push(`signature:${data.fileSignature}`);
  if (data?.fileName) keys.push(createFileNameStorageKey(data.fileName));
  return [...new Set(keys.filter(Boolean))];
}

function createFileKey(file) {
  return `${createFileSignature(file)}_${file.lastModified}`;
}

async function createFileContentHash(file) {
  if (!file?.arrayBuffer || !window.crypto?.subtle) return "";
  try {
    const buffer = await file.arrayBuffer();
    const digest = await crypto.subtle.digest("SHA-256", buffer);
    return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
  } catch {
    return "";
  }
}

function getColumnLetterFromInput(input) {
  const parsed = normalizeColumnIndexValue(input.dataset.index ?? input.value);
  return parsed === null ? "" : columnLabel(parsed);
}

function getCurrentMappingSnapshot() {
  syncTypedColumns();
  return {
    product: getColumnLetterFromInput(productColumn),
    original: getColumnLetterFromInput(originalColumn),
    finalPrice: getColumnLetterFromInput(finalPriceColumn),
    category: getColumnLetterFromInput(categoryColumn),
  };
}

function isSameMapping(mappingA = {}, mappingB = {}) {
  return ["product", "original", "finalPrice", "category"].every((key) => (mappingA[key] ?? "") === (mappingB[key] ?? ""));
}

function hasActiveColorFilter() {
  return Object.values(columnFilters).some((filter) => Boolean(filter.selected));
}

function getCurrentCalculationFingerprint() {
  return {
    sheetName: sheetSelect.value,
    mapping: getCurrentMappingSnapshot(),
    range: {
      startRow: Number(startRow.value || 1),
      endRow: Number(endRow.value || 1),
    },
    colorFilters: {
      product: { color: columnFilters.product.selected, mode: columnFilters.product.mode },
      original: { color: columnFilters.original.selected, mode: columnFilters.original.mode },
      final: { color: columnFilters.final.selected, mode: columnFilters.final.mode },
    },
    extraFields: serializeExtraFields(),
    rewardFields: serializeRewardFields(),
  };
}

function stableSettingsString(value) {
  if (Array.isArray(value)) return `[${value.map(stableSettingsString).join(",")}]`;
  if (value && typeof value === "object") {
    return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${stableSettingsString(value[key])}`).join(",")}}`;
  }
  return JSON.stringify(value ?? null);
}

function canRestoreResultSnapshot(snapshot, config) {
  if (!snapshot?.rows?.length) return false;
  if (!currentFileContentHash || config.fileContentHash !== currentFileContentHash) return false;
  if (snapshot.calculation) {
    return stableSettingsString(snapshot.calculation) === stableSettingsString(getCurrentCalculationFingerprint());
  }
  if (hasActiveColorFilter()) return false;
  const snapshotMapping = snapshot.mapping ?? config.mapping ?? {};
  const snapshotExtraFields = (snapshot.extraFields ?? config.extraFields ?? []).map(normalizeExtraFieldConfig);
  const snapshotRewardFields = (snapshot.rewardFields ?? config.rewardFields ?? []).map(normalizeRewardFieldConfig);
  return isSameMapping(snapshotMapping, getCurrentMappingSnapshot())
    && stableSettingsString(snapshotExtraFields) === stableSettingsString(serializeExtraFields())
    && stableSettingsString(snapshotRewardFields) === stableSettingsString(serializeRewardFields());
}

function getProductCleanupOptions() {
  return {
    filters: [...productNameFilters],
    removeParenthesesText: Boolean(removeParenthesesText.checked),
    removeBracketsText: Boolean(removeBracketsText.checked),
    removeAfterDelimiter: Boolean(removeAfterDelimiter.checked),
    removeBeforeDelimiter: Boolean(removeBeforeDelimiter.checked),
    removeLeadingText: Boolean(removeLeadingText.checked),
    leadingTextValue: leadingTextValue.value,
    delimiterValue: delimiterValue.value,
    titleCaseProductName: Boolean(titleCaseProductName.checked),
  };
}

function applyProductCleanupOptions(options = {}) {
  productNameFilters = Array.isArray(options.filters) ? [...options.filters] : [];
  removeParenthesesText.checked = Boolean(options.removeParenthesesText);
  removeBracketsText.checked = Boolean(options.removeBracketsText);
  removeAfterDelimiter.checked = Boolean(options.removeAfterDelimiter);
  removeBeforeDelimiter.checked = Boolean(options.removeBeforeDelimiter);
  if (removeAfterDelimiter.checked && removeBeforeDelimiter.checked) removeBeforeDelimiter.checked = false;
  removeLeadingText.checked = Boolean(options.removeLeadingText);
  leadingTextValue.value = options.leadingTextValue ?? "";
  delimiterValue.value = options.delimiterValue ?? "";
  titleCaseProductName.checked = Boolean(options.titleCaseProductName);
  setElementHidden(leadingTextOptions, !removeLeadingText.checked);
  updateDelimiterOptionsVisibility();
  renderProductFilterInputs();
}

function buildCurrentSettingsData() {
  if (!uploadedFile || !currentFileKey) return null;
  const resultSnapshot = buildResultSnapshot();
  return {
    fileKey: currentFileKey,
    fileSignature: createFileSignature(uploadedFile),
    fileName: uploadedFile.name,
    fileSize: uploadedFile.size,
    fileLastModified: uploadedFile.lastModified,
    fileContentHash: currentFileContentHash,
    savedAt: Date.now(),
    sheetName: sheetSelect.value,
    mapping: getCurrentMappingSnapshot(),
    range: {
      startRow: Number(startRow.value || 1),
      endRow: Number(endRow.value || 1),
    },
    options: {
      autoLoadColorFilters: Boolean(autoLoadColorFilters.checked),
      showWonSuffix: Boolean(showWonSuffix.checked),
      showDiscountMinus: Boolean(showDiscountMinus.checked),
      highlightFinalPrices: Boolean(highlightFinalPrices.checked),
      groupSimilarProducts: Boolean(groupSimilarProducts.checked),
      productColor: columnFilters.product.selected,
      productColorMode: columnFilters.product.mode,
      originalColor: columnFilters.original.selected,
      originalColorMode: columnFilters.original.mode,
      finalPriceColor: columnFilters.final.selected,
      finalPriceColorMode: columnFilters.final.mode,
      productCleanup: getProductCleanupOptions(),
    },
    extraFields: serializeExtraFields(),
    rewardFields: serializeRewardFields(),
    resultSnapshot,
  };
}

function buildResultSnapshot() {
  if (resultSnapshotDirty) return null;
  if (resultSection.hidden || !currentRows.length) return null;
  const currentMapping = getCurrentMappingSnapshot();
  if (lastResultMapping && !isSameMapping(lastResultMapping, currentMapping)) return null;
  return {
    version: 1,
    savedAt: Date.now(),
    rows: cloneRows(currentRows),
    extraFields: serializeExtraFields(),
    rewardFields: serializeRewardFields(),
    mapping: currentMapping,
    calculation: getCurrentCalculationFingerprint(),
    sort: {
      field: sortField.value,
      direction: sortDirection.value,
    },
  };
}

function restoreResultSnapshotIfSameFile(config) {
  const snapshot = config?.resultSnapshot;
  if (!canRestoreResultSnapshot(snapshot, config)) return false;
  extraFields = (snapshot.extraFields ?? config.extraFields ?? extraFields).map(normalizeExtraFieldConfig);
  rewardFields = (snapshot.rewardFields ?? config.rewardFields ?? rewardFields).map(normalizeRewardFieldConfig);
  renderExtraFieldInputs();
  renderRewardFieldInputs();
  renderResults({
    rows: cloneRows(snapshot.rows),
    extra_fields: extraFields,
    reward_fields: rewardFields,
  });
  lastResultMapping = getCurrentMappingSnapshot();
  resultSnapshotDirty = false;
  applyProductCleanupOptions(config.options?.productCleanup ?? config.productCleanup ?? {});
  if (snapshot.sort?.field === "custom") setSortFieldToCustomOrder();
  if (snapshot.sort?.field) sortField.value = snapshot.sort.field;
  if (snapshot.sort?.direction) sortDirection.value = snapshot.sort.direction;
  syncCustomSelect(sortField);
  syncCustomSelect(sortDirection);
  renderFilteredResults();
  setStatus("저장된 설정과 편집했던 결과를 함께 불러왔습니다.", "success");
  return true;
}

function queueSaveCurrentSettings() {
  if (!uploadedFile || !currentFileKey) return;
  saveCurrentSettingsToLocalMirror();
  window.clearTimeout(settingsSaveTimer);
  settingsSaveTimer = window.setTimeout(() => {
    saveCurrentSettings();
  }, 250);
}

async function saveCurrentSettings() {
  const data = buildCurrentSettingsData();
  if (!data) return;
  activeFileSettingsConfig = data;
  if (hasMatchingResultSnapshot(data)) activeResultSnapshotConfig = data;
  try {
    await saveFileSettings(data.fileKey, data);
  } catch (error) {
    console.warn("Failed to save file settings.", error);
  }
}

function saveCurrentSettingsToLocalMirror() {
  const data = buildCurrentSettingsData();
  if (!data) return;
  activeFileSettingsConfig = data;
  if (hasMatchingResultSnapshot(data)) activeResultSnapshotConfig = data;
  saveLocalFileSettings(data.fileKey, data);
}

function normalizeStoredConfig(config) {
  if (!config) return null;
  if (config.mapping || config.range || config.options) {
    return {
      sheet_name: config.sheetName ?? config.sheet_name ?? sheetSelect.value,
      product_col: config.mapping?.product ?? "",
      product_color: config.options?.productColor ?? "",
      product_color_mode: config.options?.productColorMode ?? "include",
      original_col: config.mapping?.original ?? "",
      original_color: config.options?.originalColor ?? "",
      original_color_mode: config.options?.originalColorMode ?? "include",
      final_price_col: config.mapping?.finalPrice ?? "",
      final_price_color: config.options?.finalPriceColor ?? "",
      final_price_color_mode: config.options?.finalPriceColorMode ?? "include",
      category_col: config.mapping?.category ?? "",
      start_row: config.range?.startRow ?? "",
      end_row: config.range?.endRow ?? "",
      auto_load_color_filters: config.options?.autoLoadColorFilters,
      show_won_suffix: config.options?.showWonSuffix,
      show_discount_minus: config.options?.showDiscountMinus,
      highlight_final_prices: config.options?.highlightFinalPrices,
      group_similar_products: config.options?.groupSimilarProducts,
      product_cleanup: config.options?.productCleanup ?? config.productCleanup ?? {},
      extra_fields: config.extraFields ?? [],
      reward_fields: config.rewardFields ?? [],
    };
  }
  return config;
}

function restoreSettings(config) {
  const normalizedConfig = normalizeStoredConfig(config);
  if (!normalizedConfig) return;
  if (normalizedConfig.sheet_name && [...sheetSelect.options].some((option) => option.value === normalizedConfig.sheet_name)) {
    sheetSelect.value = normalizedConfig.sheet_name;
    sheetSelect.dispatchEvent(new Event("change", { bubbles: true }));
  }
  productColumn.value = normalizedConfig.product_col ?? "";
  originalColumn.value = normalizedConfig.original_col ?? "";
  finalPriceColumn.value = normalizedConfig.final_price_col ?? "";
  categoryColumn.value = normalizedConfig.category_col ?? "";
  startRow.value = normalizedConfig.start_row ?? startRow.value;
  endRow.value = normalizedConfig.end_row ?? endRow.value;
  if (normalizedConfig.auto_load_color_filters !== undefined) autoLoadColorFilters.checked = Boolean(normalizedConfig.auto_load_color_filters);
  if (normalizedConfig.show_won_suffix !== undefined) showWonSuffix.checked = Boolean(normalizedConfig.show_won_suffix);
  if (normalizedConfig.show_discount_minus !== undefined) showDiscountMinus.checked = Boolean(normalizedConfig.show_discount_minus);
  if (normalizedConfig.highlight_final_prices !== undefined) highlightFinalPrices.checked = Boolean(normalizedConfig.highlight_final_prices);
  if (normalizedConfig.group_similar_products !== undefined) groupSimilarProducts.checked = Boolean(normalizedConfig.group_similar_products);
  applyProductCleanupOptions(normalizedConfig.product_cleanup ?? {});
  syncTypedColumns();
  columnFilters.product.selected = normalizedConfig.product_color ?? "";
  columnFilters.product.mode = normalizedConfig.product_color_mode === "exclude" ? "exclude" : "include";
  columnFilters.original.selected = normalizedConfig.original_color ?? "";
  columnFilters.original.mode = normalizedConfig.original_color_mode === "exclude" ? "exclude" : "include";
  columnFilters.final.selected = normalizedConfig.final_price_color ?? "";
  columnFilters.final.mode = normalizedConfig.final_price_color_mode === "exclude" ? "exclude" : "include";
  extraFields = (normalizedConfig.extra_fields ?? []).map(normalizeExtraFieldConfig);
  rewardFields = (normalizedConfig.reward_fields ?? []).map(normalizeRewardFieldConfig);
  renderExtraFieldInputs();
  renderRewardFieldInputs();
  loadAllColumnColors();
  updateCalculateButtonState();
  applyPriceHighlighting();
}

function normalizeExtraFieldConfig(field) {
  return {
    ...field,
    source_col: normalizeColumnIndexValue(field.source_col) ?? "",
    operand_col: normalizeColumnIndexValue(field.operand_col) ?? "",
    left_col: normalizeColumnIndexValue(field.left_col) ?? "",
    right_col: normalizeColumnIndexValue(field.right_col) ?? "",
  };
}

function normalizeRewardFieldConfig(field) {
  return {
    ...field,
    source_col: normalizeColumnIndexValue(field.source_col) ?? "",
  };
}

function openDB() {
  if (!window.indexedDB) return Promise.resolve(null);
  return new Promise((resolve) => {
    let request;
    try {
      request = indexedDB.open(DB_NAME, DB_VERSION);
    } catch {
      resolve(null);
      return;
    }
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(SETTINGS_STORE_NAME)) {
        db.createObjectStore(SETTINGS_STORE_NAME, { keyPath: "fileKey" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => resolve(null);
    request.onblocked = () => resolve(null);
  });
}

async function withSettingsStore(mode, callback) {
  const db = await openDB();
  if (!db) return null;
  return new Promise((resolve, reject) => {
    let transaction;
    try {
      transaction = db.transaction(SETTINGS_STORE_NAME, mode);
    } catch (error) {
      db.close();
      reject(error);
      return;
    }
    const store = transaction.objectStore(SETTINGS_STORE_NAME);
    let settled = false;
    transaction.oncomplete = () => {
      db.close();
      if (!settled) resolve(undefined);
    };
    transaction.onerror = () => {
      db.close();
      reject(transaction.error);
    };
    transaction.onabort = () => {
      db.close();
      reject(transaction.error);
    };
    try {
      callback(store, (value) => {
        settled = true;
        resolve(value);
      });
    } catch (error) {
      transaction.abort();
      reject(error);
    }
  });
}

async function requestServerSettings(path, payload = null) {
  if (!/^https?:$/.test(location.protocol)) return null;
  try {
    const response = await fetch(path, payload
      ? {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      }
      : { method: "GET" });
    if (!response.ok) return null;
    return await response.json();
  } catch {
    return null;
  }
}

async function saveServerFileSettings(fileKey, data) {
  if (!fileKey || !data) return;
  await requestServerSettings("/api/settings/save", { fileKey, data: { ...data, fileKey } });
}

async function loadServerFileSettings(fileKey) {
  if (!fileKey) return null;
  const result = await requestServerSettings(`/api/settings/load?file_key=${encodeURIComponent(fileKey)}`);
  return result?.settings ?? null;
}

async function loadLatestServerSettingsByName(fileName) {
  if (!fileName) return null;
  const result = await requestServerSettings(`/api/settings/latest?file_name=${encodeURIComponent(fileName)}`);
  return result?.settings ?? null;
}

async function deleteServerFileSettings(fileKey) {
  if (!fileKey) return;
  await requestServerSettings("/api/settings/delete", { fileKey });
}

async function saveFileSettings(fileKey, data) {
  if (!fileKey || !data) return;
  saveLocalFileSettings(fileKey, data);
  saveServerFileSettings(fileKey, data).catch(() => {});
  const storageKeys = getSettingsStorageKeys(fileKey, data);
  try {
    const saved = await withSettingsStore("readwrite", (store) => {
      storageKeys.forEach((key) => store.put({ ...data, fileKey: key, sourceFileKey: fileKey }));
    });
    if (saved === null) storageKeys.forEach((key) => fallbackSettingsStore.set(key, { ...data, fileKey: key, sourceFileKey: fileKey }));
  } catch {
    storageKeys.forEach((key) => fallbackSettingsStore.set(key, { ...data, fileKey: key, sourceFileKey: fileKey }));
  }
}

async function loadSettingsForUploadedFile(file) {
  const candidates = [
    await loadFileSettings(createFileKey(file)),
    await loadFileSettings(createFileSignatureStorageKey(file)),
    await loadFileSettings(createFileNameStorageKey(file.name)),
    await loadLatestFileSettingsBySignature(createFileSignature(file)),
    await loadLatestFileSettingsByName(file.name),
  ];
  const lastSettings = readLastLocalSettings();
  if (isLikelySameFileSetting(lastSettings, file)) candidates.push(lastSettings);
  return pickBestSettingsForCurrentFile(...candidates);
}

async function loadLatestFileSettingsBySignature(fileSignature) {
  if (!fileSignature) return null;
  const localLatest = loadLatestLocalSettingsBySignature(fileSignature);
  try {
    const result = await withSettingsStore("readonly", (store, done) => {
      const request = store.openCursor();
      let latest = null;
      request.onsuccess = () => {
        const cursor = request.result;
        if (!cursor) {
          done(latest);
          return;
        }
        const value = cursor.value;
        if (value?.fileSignature === fileSignature && (!latest || (value.savedAt ?? 0) > (latest.savedAt ?? 0))) {
          latest = value;
        }
        cursor.continue();
      };
      request.onerror = () => done(null);
    });
    return pickLatestSettings(result, localLatest);
  } catch {
    let latest = null;
    fallbackSettingsStore.forEach((value) => {
      if (value?.fileSignature === fileSignature && (!latest || (value.savedAt ?? 0) > (latest.savedAt ?? 0))) latest = value;
    });
    return pickLatestSettings(latest, localLatest);
  }
}
async function loadLatestFileSettingsByName(fileName) {
  if (!fileName) return null;
  const localLatest = loadLatestLocalSettingsByName(fileName);
  const serverLatest = await loadLatestServerSettingsByName(fileName);
  try {
    const result = await withSettingsStore("readonly", (store, done) => {
      const request = store.openCursor();
      let latest = null;
      request.onsuccess = () => {
        const cursor = request.result;
        if (!cursor) {
          done(latest);
          return;
        }
        const value = cursor.value;
        if (value?.fileName === fileName && (!latest || (value.savedAt ?? 0) > (latest.savedAt ?? 0))) {
          latest = value;
        }
        cursor.continue();
      };
      request.onerror = () => done(null);
    });
    return pickLatestSettings(result, localLatest, serverLatest);
  } catch {
    let latest = null;
    fallbackSettingsStore.forEach((value) => {
      if (value?.fileName === fileName && (!latest || (value.savedAt ?? 0) > (latest.savedAt ?? 0))) latest = value;
    });
    return pickLatestSettings(latest, localLatest, serverLatest);
  }
}

async function loadFileSettings(fileKey) {
  if (!fileKey) return null;
  const server = await loadServerFileSettings(fileKey);
  const local = loadLocalFileSettings(fileKey);
  let indexed = null;
  try {
    indexed = await withSettingsStore("readonly", (store, done) => {
      const request = store.get(fileKey);
      request.onsuccess = () => done(request.result ?? null);
      request.onerror = () => done(null);
    });
  } catch {
    indexed = null;
  }
  return pickBestSettingsForCurrentFile(server, indexed, local, fallbackSettingsStore.get(fileKey));
}

async function deleteFileSettings(fileKey) {
  if (!fileKey) return;
  fallbackSettingsStore.delete(fileKey);
  deleteLocalFileSettings(fileKey);
  deleteServerFileSettings(fileKey).catch(() => {});
  try {
    await withSettingsStore("readwrite", (store) => {
      store.delete(fileKey);
    });
  } catch {
    // Keep the app running even when IndexedDB is unavailable.
  }
}

async function cleanupStoredSettings() {
  const cutoff = Date.now() - SETTINGS_RETENTION_DAYS * 24 * 60 * 60 * 1000;
  cleanupLocalFileSettings(cutoff);
  try {
    await withSettingsStore("readwrite", (store) => {
      const request = store.openCursor();
      request.onsuccess = () => {
        const cursor = request.result;
        if (!cursor) return;
        if (!cursor.value?.savedAt || cursor.value.savedAt < cutoff) cursor.delete();
        cursor.continue();
      };
    });
  } catch {
    // Cleanup failures should not affect calculator behavior.
  }
}

function markStorageUsage() {
  try {
    localStorage.setItem("WRExcelStorageMarker", JSON.stringify({ savedAt: Date.now(), version: "20260527-3" }));
  } catch {
    // Storage can be blocked by browser settings. Settings saving will fall back where possible.
  }
}
async function requestPersistentStorage() {
  try {
    if (!navigator.storage?.persisted || !navigator.storage?.persist) return;
    const alreadyPersisted = await navigator.storage.persisted();
    if (!alreadyPersisted) await navigator.storage.persist();
  } catch {
    // Some browsers do not support persistent storage requests. The app still uses normal browser storage.
  }
}
function readLocalSettingsMap() {
  try {
    return JSON.parse(localStorage.getItem(LOCAL_SETTINGS_KEY) || "{}");
  } catch {
    return {};
  }
}

function writeLocalSettingsMap(map) {
  try {
    localStorage.setItem(LOCAL_SETTINGS_KEY, JSON.stringify(map));
  } catch {
    // localStorage can be blocked in private mode. IndexedDB/fallback map will still be attempted.
  }
}

function saveLocalFileSettings(fileKey, data) {
  if (!fileKey || !data) return;
  const savedAt = Date.now();
  const normalizedData = { ...data, savedAt };
  const map = readLocalSettingsMap();
  getSettingsStorageKeys(fileKey, normalizedData).forEach((key) => {
    map[key] = { ...normalizedData, fileKey: key, sourceFileKey: fileKey };
  });
  writeLocalSettingsMap(map);
  writeLastLocalSettings({ ...normalizedData, fileKey, sourceFileKey: fileKey });
}

function writeLastLocalSettings(data) {
  try {
    localStorage.setItem(LAST_SETTINGS_KEY, JSON.stringify(data));
  } catch {
    // Ignore blocked storage.
  }
}

function readLastLocalSettings() {
  try {
    return JSON.parse(localStorage.getItem(LAST_SETTINGS_KEY) || "null");
  } catch {
    return null;
  }
}

function isLikelySameFileSetting(setting, file) {
  if (!setting || !file) return false;
  if (setting.fileSignature && setting.fileSignature === createFileSignature(file)) return true;
  if (setting.fileName && setting.fileName === file.name) return true;
  return false;
}
function loadLocalFileSettings(fileKey) {
  if (!fileKey) return null;
  return readLocalSettingsMap()[fileKey] ?? null;
}

function deleteLocalFileSettings(fileKey) {
  const map = readLocalSettingsMap();
  if (!(fileKey in map)) return;
  delete map[fileKey];
  writeLocalSettingsMap(map);
}

function loadLatestLocalSettingsBySignature(fileSignature) {
  const values = Object.values(readLocalSettingsMap());
  return values.reduce((latest, value) => {
    if (value?.fileSignature !== fileSignature) return latest;
    if (!latest || (value.savedAt ?? 0) > (latest.savedAt ?? 0)) return value;
    return latest;
  }, null);
}
function loadLatestLocalSettingsByName(fileName) {
  const values = Object.values(readLocalSettingsMap());
  return values.reduce((latest, value) => {
    if (value?.fileName !== fileName) return latest;
    if (!latest || (value.savedAt ?? 0) > (latest.savedAt ?? 0)) return value;
    return latest;
  }, null);
}

function cleanupLocalFileSettings(cutoff) {
  const map = readLocalSettingsMap();
  let changed = false;
  Object.entries(map).forEach(([key, value]) => {
    if (!value?.savedAt || value.savedAt < cutoff) {
      delete map[key];
      changed = true;
    }
  });
  if (changed) writeLocalSettingsMap(map);
}

function hasMatchingResultSnapshot(setting) {
  return Boolean(
    setting?.resultSnapshot?.rows?.length
    && currentFileContentHash
    && setting.fileContentHash === currentFileContentHash
  );
}

function pickBestSettingsForCurrentFile(...items) {
  const validItems = items.filter(Boolean);
  const resultSnapshots = validItems.filter(hasMatchingResultSnapshot);
  if (resultSnapshots.length) return pickLatestSettings(...resultSnapshots);
  return pickLatestSettings(...validItems);
}

function pickLatestSettings(...items) {
  return items
    .filter(Boolean)
    .sort((a, b) => (b.savedAt ?? 0) - (a.savedAt ?? 0))[0] ?? null;
}

function columnValue(index) {
  const parsed = normalizeColumnIndexValue(index);
  return parsed === null ? "" : columnLabel(parsed);
}

function normalizeColumnIndexValue(value) {
  if (value === "" || value === null || value === undefined) return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const text = String(value).trim();
  if (text === "") return null;
  if (/^\d+$/.test(text)) return Number(text);
  return parseColumnInput(text);
}

function persistDynamicColumnInput(input, colIndex = null) {
  const card = input.closest(".extra-field-card, .reward-field-card");
  if (!card) return;
  clearRestoredResultSnapshot();
  const parsed = colIndex ?? normalizeColumnIndexValue(input.value);
  const value = parsed === null ? "" : parsed;
  const role = input.dataset.role;

  if (card.classList.contains("extra-field-card")) {
    const field = extraFields.find((item) => item.id === card.dataset.fieldId);
    if (field && role) field[role] = value;
    return;
  }

  if (card.classList.contains("reward-field-card")) {
    const field = rewardFields.find((item) => item.id === card.dataset.fieldId);
    if (field && role) field[role] = value;
  }
}

function renderExtraResultHeaders() {
  document.querySelectorAll("th[data-extra-column]").forEach((node) => node.remove());
  const panel = visibleColumnsPanel;
  panel.querySelectorAll("[data-extra-visible-column]").forEach((node) => node.closest("label").remove());
  const rewardStartLabel = panel.querySelector('[data-visible-column="total_reward_amount"]')?.closest("label");
  extraFields.forEach((field) => {
    const header = document.createElement("th");
    header.dataset.column = `extra_${field.id}`;
    header.dataset.extraColumn = field.id;
    header.className = "extra-column";
    header.textContent = field.name;
    extraColumnsAnchor.before(header);

    const label = document.createElement("label");
    label.innerHTML = `<input type="checkbox" data-visible-column="extra_${field.id}" data-extra-visible-column="${field.id}" checked /> ${escapeHtml(field.name)}`;
    const checkbox = label.querySelector("input");
    checkbox.addEventListener("change", () => {
      const column = checkbox.dataset.visibleColumn;
      if (checkbox.checked) hiddenColumns.delete(column);
      else hiddenColumns.add(column);
      applyVisibleColumns();
      updateFloatingResultHeader();
    });
    panel.insertBefore(label, rewardStartLabel);
  });
}




function renderRewardFieldInputs() {
  rewardFieldsList.innerHTML = "";
  rewardFields.forEach((field) => {
    field.name_choice ??= rewardFieldNameSuggestions.includes(field.name) ? field.name : "";
    field.custom_name ??= field.name_choice === "custom" ? field.name : "";
    const nameSelectId = `rewardNameSelect_${field.id}`;
    const card = document.createElement("div");
    card.className = "reward-field-card";
    card.dataset.fieldId = field.id;
    card.innerHTML = `
      <label>
        <span>\uD56D\uBAA9\uBA85</span>
        <div class="custom-select reward-name-select" data-select-target="${nameSelectId}">
          <button class="custom-select-trigger" type="button"></button>
          <div class="custom-select-menu hidden"></div>
        </div>
        <select id="${nameSelectId}" data-role="name_choice" class="native-select-hidden" aria-hidden="true" tabindex="-1">
          <option value="">\uD56D\uBAA9 \uC120\uD0DD</option>
          <option value="\uC2A4\uB9C8\uC77C\uCE74\uB4DC">\uC2A4\uB9C8\uC77C\uCE74\uB4DC</option>
          <option value="\uBA38\uB2C8\uCDA9\uC804">\uBA38\uB2C8\uCDA9\uC804</option>
          <option value="\uAF2D\uBA64\uBC841">\uAF2D\uBA64\uBC841</option>
          <option value="\uAF2D\uBA64\uBC842">\uAF2D\uBA64\uBC842</option>
          <option value="custom">\uC9C1\uC811\uC785\uB825</option>
        </select>
        <div class="extra-custom-name-wrap hidden-field">
          <input class="extra-custom-name" data-role="custom_name" value="${escapeAttribute(field.custom_name)}" placeholder="\uD56D\uBAA9\uBA85\uC744 \uC785\uB825\uD558\uC138\uC694" />
          <button class="extra-custom-select-toggle" type="button" aria-label="\uD56D\uBAA9\uBA85 \uC120\uD0DD \uC5F4\uAE30"></button>
          <div class="extra-custom-select-menu hidden">
            <button type="button" data-name-choice="">\uD56D\uBAA9 \uC120\uD0DD</button>
            <button type="button" data-name-choice="\uC2A4\uB9C8\uC77C\uCE74\uB4DC">\uC2A4\uB9C8\uC77C\uCE74\uB4DC</button>
            <button type="button" data-name-choice="\uBA38\uB2C8\uCDA9\uC804">\uBA38\uB2C8\uCDA9\uC804</button>
            <button type="button" data-name-choice="\uAF2D\uBA64\uBC841">\uAF2D\uBA64\uBC841</button>
            <button type="button" data-name-choice="\uAF2D\uBA64\uBC842">\uAF2D\uBA64\uBC842</button>
          </div>
        </div>
      </label>
      <label>
        <span>\uC801\uB9BD\uAE08\uC561 \uC5F4</span>
        <input data-role="source_col" placeholder="\uC608: H" value="${columnValue(field.source_col)}" />
      </label>
      <button class="ghost-button extra-field-remove" type="button">\uC0AD\uC81C</button>
    `;
    const nameSelect = card.querySelector('[data-role="name_choice"]');
    nameSelect.value = field.name_choice;
    rewardFieldsList.appendChild(card);
    mountCustomSelect(card.querySelector(`[data-select-target="${nameSelectId}"]`), nameSelect);

    card.querySelectorAll("input").forEach((input) => {
      input.addEventListener("focus", () => {
        if (input.dataset.role === "custom_name") return;
        activeExtraFieldTarget = input;
        setActivePick("extra");
      });
      input.addEventListener("input", () => {
        updateRewardFieldFromCard(field, card);
        persistDynamicColumnInput(input);
        clearRestoredResultSnapshot();
        queueSaveCurrentSettings();
      });
    });
    nameSelect.addEventListener("change", () => {
      updateRewardFieldFromCard(field, card);
      updateRewardNameVisibility(card, field.name_choice);
      activateDynamicFieldColumnTarget(card);
      clearRestoredResultSnapshot();
      queueSaveCurrentSettings();
    });
    const customToggle = card.querySelector(".extra-custom-select-toggle");
    const customMenu = card.querySelector(".extra-custom-select-menu");
    customToggle.addEventListener("click", (event) => {
      event.stopPropagation();
      customMenu.classList.toggle("hidden");
    });
    customMenu.querySelectorAll("button").forEach((button) => {
      button.addEventListener("click", () => {
        field.name_choice = button.dataset.nameChoice;
        nameSelect.value = field.name_choice;
        updateRewardFieldFromCard(field, card);
        updateRewardNameVisibility(card, field.name_choice);
        syncCustomSelect(nameSelect);
        activateDynamicFieldColumnTarget(card);
        customMenu.classList.add("hidden");
        clearRestoredResultSnapshot();
        queueSaveCurrentSettings();
      });
    });
    card.querySelector(".extra-field-remove").addEventListener("click", () => {
      rewardFields = rewardFields.filter((item) => item.id !== field.id);
      clearRestoredResultSnapshot();
      renderRewardFieldInputs();
      queueSaveCurrentSettings();
    });
    updateRewardNameVisibility(card, field.name_choice);
  });
}

function updateRewardFieldFromCard(field, card) {
  field.name_choice = card.querySelector('[data-role="name_choice"]').value;
  field.custom_name = card.querySelector('[data-role="custom_name"]').value;
  field.name = field.name_choice === "custom" ? field.custom_name || "\uC801\uB9BD\uAE08\uC561" : field.name_choice || "\uC801\uB9BD\uAE08\uC561";
  const parsed = normalizeColumnIndexValue(card.querySelector('[data-role="source_col"]').value);
  field.source_col = parsed === null ? "" : parsed;
}

function updateRewardNameVisibility(card, nameChoice) {
  card.querySelector(".reward-name-select").classList.toggle("hidden-field", nameChoice === "custom");
  card.querySelector(".extra-custom-name-wrap").classList.toggle("hidden-field", nameChoice !== "custom");
}

function renderRewardResultHeaders() {
  document.querySelectorAll("th[data-reward-column]").forEach((node) => node.remove());
  const panel = visibleColumnsPanel;
  panel.querySelectorAll("[data-reward-visible-column]").forEach((node) => node.closest("label").remove());
  const totalRewardLabel = panel.querySelector('[data-visible-column="total_reward_amount"]')?.closest("label");
  rewardFields.forEach((field) => {
    const header = document.createElement("th");
    header.dataset.column = `reward_${field.id}`;
    header.dataset.rewardColumn = field.id;
    header.className = `reward-column${rewardFields.indexOf(field) === 0 ? " reward-group-start" : ""}`;
    header.textContent = field.name;
    rewardColumnsAnchor.before(header);

    const label = document.createElement("label");
    label.innerHTML = `<input type="checkbox" data-visible-column="reward_${field.id}" data-reward-visible-column="${field.id}" checked /> ${escapeHtml(field.name)}`;
    const checkbox = label.querySelector("input");
    checkbox.addEventListener("change", () => {
      const column = checkbox.dataset.visibleColumn;
      if (checkbox.checked) hiddenColumns.delete(column);
      else hiddenColumns.add(column);
      applyVisibleColumns();
      updateFloatingResultHeader();
    });
    panel.insertBefore(label, totalRewardLabel);
  });
}

function startInlineEdit(button, row) {
  let cell = button.closest("td");
  const column = cell?.dataset.column ?? "product_name";
  if (!cell || column === "total_discount_amount" || column === "discount_rate" || column === "total_reward_amount" || column === "effective_price") {
    return;
  }
  if (activeInlineEdit?.input === button) return;
  if (activeInlineEdit) {
    commitInlineEdit(true);
    const selector = `tr[data-row-id="${row.row_id}"] td[data-column="${cssEscape(column)}"] [data-copy-value]`;
    button = resultBody.querySelector(selector);
    cell = button?.closest("td");
    if (!button || !cell) return;
  }

  const input = document.createElement("input");
  input.className = "inline-edit-input";
  input.value = getEditableValue(row, column);
  button.replaceWith(input);
  activeInlineEdit = { input, row, column };
  input.focus();
  input.select();

  input.addEventListener("blur", () => {
    window.setTimeout(() => {
      if (activeInlineEdit?.input === input) commitInlineEdit(true);
    }, 0);
  }, { once: true });
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") commitInlineEdit(true);
    if (event.key === "Escape") {
      activeInlineEdit = null;
      renderFilteredResults();
    }
  });
}

function commitInlineEdit(shouldRender = true) {
  if (!activeInlineEdit) return;
  const { input, row, column } = activeInlineEdit;
  activeInlineEdit = null;
  applyEditedValue(row, column, input.value);
  recalculateDerivedValues(row);
  if (shouldRender) renderFilteredResults();
}

function getEditableValue(row, column) {
  if (column === "product_name") return row.product_name ?? "";
  if (column === "original_price") return row.original_price ?? "";
  if (column === "discount_amount") return row.discount_amount ?? "";
  if (column.startsWith("extra_")) return row.extra_values?.[column.replace("extra_", "")] ?? "";
  if (column.startsWith("reward_")) return row.reward_values?.[column.replace("reward_", "")] ?? "";
  return "";
}

function applyEditedValue(row, column, value) {
  if (column === "product_name") {
    row.product_name = value;
    return;
  }
  const numericValue = parseEditedNumber(value);
  if (column === "original_price") {
    row.base_original_price = numericValue;
    row.original_price = numericValue;
  } else if (column === "discount_amount") {
    row.discount_amount = numericValue;
    row.has_extra_price_formula = false;
  } else if (column.startsWith("extra_")) {
    row.extra_values = row.extra_values ?? {};
    row.extra_values[column.replace("extra_", "")] = numericValue;
    row.has_extra_price_formula = true;
  } else if (column.startsWith("reward_")) {
    row.reward_values = row.reward_values ?? {};
    row.reward_values[column.replace("reward_", "")] = numericValue;
  }
}

function parseEditedNumber(value) {
  const cleaned = String(value).replace(/[^\d.-]/g, "");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function recalculateDerivedValues(row) {
  const baseOriginalPrice = Number.isFinite(Number(row.base_original_price))
    ? Number(row.base_original_price)
    : Number(row.original_price);
  row.base_original_price = baseOriginalPrice;
  row.original_price = Math.round(baseOriginalPrice);
  if (row.has_extra_price_formula) {
    row.discount_amount = calculateFinalPriceFromExtraFields(baseOriginalPrice, row.extra_values ?? {}, extraFields);
  }
  row.total_discount_amount = Math.round((row.original_price ?? 0) - (row.discount_amount ?? 0));
  row.discount_rate = row.original_price > 0 ? Math.round((row.total_discount_amount / row.original_price) * 1000) / 10 : 0;
  row.total_reward_amount = sumNumericObjectValues(row.reward_values ?? {});
  row.effective_price = Math.round((row.discount_amount ?? 0) - row.total_reward_amount);
}

function calculateFinalPriceFromExtraFields(originalPrice, values, fields) {
  const orderedFields = fields ?? [];
  if (orderedFields.length === 0) return Math.round(Number(originalPrice) || 0);
  return Math.round(orderedFields.reduce((total, field) => {
    const amount = parseEditedNumber(values?.[field.id]);
    return total - amount;
  }, Number(originalPrice) || 0));
}

function sumNumericObjectValues(values) {
  return Object.values(values ?? {}).reduce((sum, value) => {
    const numericValue = parseEditedNumber(value);
    return sum + (Number.isFinite(numericValue) ? numericValue : 0);
  }, 0);
}

function renderProductFilterInputs() {
  productFilterList.innerHTML = "";
  productNameFilters.forEach((value, index) => {
    const row = document.createElement("div");
    row.className = "product-filter-row";
    row.innerHTML = `
      <input type="text" value="${escapeAttribute(value)}" placeholder="제외할 문구" />
      <button type="button" aria-label="삭제">×</button>
    `;
    const input = row.querySelector("input");
    const removeButton = row.querySelector("button");
    input.addEventListener("input", () => {
      productNameFilters[index] = input.value;
      renderFilteredResults();
      queueSaveCurrentSettings();
    });
    removeButton.addEventListener("click", (event) => {
      event.stopPropagation();
      productNameFilters.splice(index, 1);
      renderProductFilterInputs();
      renderFilteredResults();
      queueSaveCurrentSettings();
      setElementHidden(productFilterPanel, false);
    });
    productFilterList.appendChild(row);
  });
}

function setStatus(message, type) {
  statusCard.hidden = false;
  statusCard.textContent = message;
  statusCard.className = `panel status-panel ${type}`;
}

function hideResults() {
  productColumnDragWidth = null;
  productColumnUserResized = false;
  setElementHidden(resultSection, true);
}

function showConfirmDialog(message, onConfirm = null, onCancel = null) {
  showAppDialog({
    type: "confirm",
    title: "\uD655\uC778\uC774 \uD544\uC694\uD569\uB2C8\uB2E4",
    message,
    actions: [
      {
        label: "아니오",
        className: "ghost-button",
        handler: () => {
          if (onCancel) onCancel();
          else {
            pendingPayload = null;
            hideResults();
            setStatus("\uACC4\uC0B0\uC744 \uCDE8\uC18C\uD588\uC2B5\uB2C8\uB2E4. \uC120\uD0DD\uD55C \uC5F4\uC744 \uB2E4\uC2DC \uD655\uC778\uD574\uC8FC\uC138\uC694.", "error");
          }
          hideConfirmDialog();
        },
      },
      {
        label: "예",
        className: "primary-button",
        handler: async () => {
          if (onConfirm) await onConfirm();
          else if (pendingPayload) {
            const payload = pendingPayload;
            pendingPayload = null;
            await finishRender(payload);
          }
          hideConfirmDialog();
        },
      },
    ],
  });
}

function hideConfirmDialog() {
  hideDialog(appDialog);
}

function showMessageDialog(message) {
  showAppDialog({
    type: "message",
    title: "\uC548\uB0B4",
    message,
    actions: [
      {
        label: "확인",
        className: "primary-button",
        handler: hideMessageDialog,
      },
    ],
  });
}

function hideMessageDialog() {
  hideDialog(appDialog);
}

function showRestoreDialog() {
  showAppDialog({
    type: "restore",
    title: "\uC800\uC7A5\uB41C \uB0B4\uC5ED\uC774 \uC788\uC2B5\uB2C8\uB2E4",
    message: "\uC774 \uD30C\uC77C\uC5D0 \uB300\uD574 \uC774\uC804\uC5D0 \uC800\uC7A5\uD55C \uC124\uC815\uC774 \uC788\uC2B5\uB2C8\uB2E4. \uBD88\uB7EC\uC624\uC2DC\uACA0\uC2B5\uB2C8\uAE4C?",
    actions: [
      {
        label: "아니오",
        className: "ghost-button",
        handler: () => {
          pendingRestoreConfig = null;
          activeResultSnapshotConfig = null;
          resultSnapshotDirty = true;
          hideRestoreDialog();
        },
      },
      {
        label: "예",
        className: "primary-button",
        handler: () => {
          if (pendingRestoreConfig) {
            restoreSettings(pendingRestoreConfig);
            restoreResultSnapshotIfSameFile(pendingRestoreConfig);
            queueSaveCurrentSettings();
          }
          pendingRestoreConfig = null;
          hideRestoreDialog();
        },
      },
    ],
  });
}

function hideRestoreDialog() {
  hideDialog(appDialog);
}

function showAppDialog({ type = "message", title = "", message = "", actions = [] }) {
  appDialog.dataset.dialogType = type;
  dialogTitle.textContent = title;
  dialogMessage.textContent = message;
  dialogActions.innerHTML = "";
  actions.forEach((action) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = action.className;
    button.textContent = action.label;
    button.addEventListener("click", action.handler);
    dialogActions.appendChild(button);
  });
  showDialog(appDialog);
}

function showDialog(dialog) {
  dialog.hidden = false;
  dialog.classList.remove("hidden");
}

function hideDialog(dialog) {
  dialog.hidden = true;
  dialog.classList.add("hidden");
}

function setCalculating(isCalculating) {
  calculateButton.disabled = isCalculating || !hasRequiredColumns();
  calculateButton.textContent = isCalculating ? "계산 중..." : "계산하기";
  calculateButton.classList.toggle("loading", isCalculating);
}

function hasRequiredColumns() {
  return Boolean(productColumn.dataset.index && originalColumn.dataset.index && finalPriceColumn.dataset.index);
}

function updateCalculateButtonState() {
  calculateButton.disabled = !hasRequiredColumns();
}

function initializeCustomSelects() {
  document.querySelectorAll('[data-role="select"], .custom-select').forEach((wrapper) => {
    const select = findNativeSelectForWrapper(wrapper);
    if (select) mountCustomSelect(wrapper, select);
  });

  document.addEventListener("click", (event) => {
    if (!event.target.closest('[data-role="select"], .custom-select')) closeAllCustomSelects();
    if (!event.target.closest(".extra-custom-name-wrap")) {
      document.querySelectorAll(".extra-custom-select-menu").forEach((menu) => setElementHidden(menu, true));
    }
  });
}

function findNativeSelectForWrapper(wrapper) {
  const target = wrapper.dataset.target || wrapper.dataset.selectTarget;
  return wrapper.parentElement?.querySelector('[data-role="native-select"]')
    || (target ? document.querySelector(`[data-role="native-select"][data-target="${target}"]`) : null)
    || (target ? document.querySelector(`#${target}`) : null);
}

function setElementHidden(element, hidden) {
  if (!element) return;
  element.hidden = hidden;
  element.classList.toggle("hidden", hidden);
}

function closeAllCustomSelects(exceptWrapper = null) {
  document.querySelectorAll('[data-role="select"], .custom-select').forEach((wrapper) => {
    if (wrapper === exceptWrapper) return;
    wrapper.classList.remove("open");
    setElementHidden(wrapper.querySelector('[data-role="select-menu"], .custom-select-menu'), true);
  });
}

function mountCustomSelect(wrapper, select) {
  if (wrapper.dataset.mounted === "true") {
    syncCustomSelect(select);
    return;
  }
  wrapper.dataset.mounted = "true";
  const trigger = wrapper.querySelector('[data-role="select-trigger"], .custom-select-trigger');
  const menu = wrapper.querySelector('[data-role="select-menu"], .custom-select-menu');
  trigger.addEventListener("click", (event) => {
    event.stopPropagation();
    closeAllCustomSelects(wrapper);
    const nextOpen = !wrapper.classList.contains("open");
    wrapper.classList.toggle("open", nextOpen);
    setElementHidden(menu, !nextOpen);
  });
  select.addEventListener("change", () => syncCustomSelect(select));
  syncCustomSelect(select);
}

function syncCustomSelect(select) {
  const target = select.dataset.target || select.id;
  const wrapper = target
    ? document.querySelector(`[data-role="select"][data-target="${target}"], .custom-select[data-select-target="${target}"]`)
    : select.parentElement?.querySelector('[data-role="select"], .custom-select');
  if (!wrapper) return;
  const trigger = wrapper.querySelector('[data-role="select-trigger"], .custom-select-trigger');
  const menu = wrapper.querySelector('[data-role="select-menu"], .custom-select-menu');
  const selectedOption = select.options[select.selectedIndex];

  trigger.textContent = selectedOption?.textContent ?? "";
  menu.innerHTML = "";

  [...select.options].forEach((option) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = `custom-select-option${option.value === select.value ? " selected" : ""}`;
    button.textContent = option.textContent;
    button.addEventListener("click", () => {
      select.value = option.value;
      select.dispatchEvent(new Event("change", { bubbles: true }));
      wrapper.classList.remove("open");
      setElementHidden(menu, true);
    });
    menu.appendChild(button);
  });

  if (target === "sheet" || select.id === "sheetSelect") {
    requestAnimationFrame(() => {
      menu.classList.toggle("sheet-scrollable", menu.scrollWidth > menu.clientWidth);
    });
  }
}

async function copyText(text) {
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch {
    // file:// 환경 등에서는 아래 방식으로 한 번 더 시도합니다.
  }

  try {
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    const copied = document.execCommand("copy");
    textarea.remove();
    return copied;
  } catch {
    return false;
  }
}

function startProductColumnResize(event) {
  event.preventDefault();
  const startX = event.clientX;
  const visualStartWidth = productNameHeader.getBoundingClientRect().width;
  const startWidth = visualStartWidth;
  const maxWidth = getMaxProductColumnWidth();
  productColumnUserResized = true;

  function onPointerMove(moveEvent) {
    const targetWidth = startWidth + moveEvent.clientX - startX;
    productColumnDragWidth = clampProductColumnWidth(targetWidth, maxWidth, true);
    productColumnWidth = productColumnDragWidth;
    applyProductColumnWidth();
  }

  function onPointerUp() {
    window.removeEventListener("pointermove", onPointerMove);
    window.removeEventListener("pointerup", onPointerUp);
  }

  window.addEventListener("pointermove", onPointerMove);
  window.addEventListener("pointerup", onPointerUp);
}

function clampProductColumnWidth(width, maxWidth = getMaxProductColumnWidth(), isUserResize = false) {
  const minWidth = isUserResize ? 180 : getMinProductColumnWidth();
  return Math.max(minWidth, Math.min(Math.round(width), maxWidth));
}

function getMinProductColumnWidth() {
  const table = tableWrap?.querySelector("table");
  const availableWidth = tableWrap?.clientWidth || 0;
  if (!table || !availableWidth) return 180;
  const otherColumnsWidth = getVisibleResultHeaders()
    .filter((header) => header.dataset.column !== "product_name")
    .reduce((sum, header) => sum + getResultColumnPixelWidth(header), 0);
  return Math.max(180, Math.ceil(availableWidth - otherColumnsWidth));
}

function getMaxProductColumnWidth() {
  return Math.max(260, Math.ceil(measureMaxProductNameWidth()));
}

function scheduleResultColumnLayoutSync() {
  if (resultSection.hidden) {
    updateFloatingResultHeader();
    return;
  }
  requestAnimationFrame(() => applyProductColumnWidth());
}

function measureMaxProductNameWidth() {
  const measuringElement = document.createElement("span");
  measuringElement.className = "product-width-measurer";
  const sampleButton = resultBody.querySelector('[data-column="product_name"] .copy-value') ?? productNameHeader;
  const computedStyle = window.getComputedStyle(sampleButton);
  measuringElement.style.font = computedStyle.font;
  measuringElement.style.fontWeight = computedStyle.fontWeight;
  document.body.appendChild(measuringElement);

  const names = currentRows.length > 0
    ? currentRows.map((row) => displayProductName(row.product_name ?? ""))
    : [productNameHeader.textContent ?? "???"];
  const maxTextWidth = names.reduce((maxWidth, name) => {
    measuringElement.textContent = name;
    return Math.max(maxWidth, measuringElement.getBoundingClientRect().width);
  }, 0);
  measuringElement.remove();

  const editControlsWidth = isEditMode ? 58 : 0;
  return maxTextWidth + editControlsWidth + 44;
}

function applyProductColumnWidth() {
  const productCells = [
    productNameHeader,
    ...resultBody.querySelectorAll('tr[data-row-id] > td[data-column="product_name"]'),
  ].filter(Boolean);

  if (productColumnWidth === null) {
    appliedProductColumnWidth = null;
    appliedResultColumnWidths = [];
    productCells.forEach((cell) => {
      cell.style.removeProperty("width");
      cell.style.removeProperty("min-width");
      cell.style.removeProperty("max-width");
    });
    syncResultColumnWidths();
    updateFloatingResultHeader();
    return;
  }

  productColumnWidth = clampProductColumnWidth(productColumnWidth, getMaxProductColumnWidth(), productColumnUserResized);
  productCells.forEach((cell) => {
    cell.style.width = `${productColumnWidth}px`;
    cell.style.minWidth = `${productColumnWidth}px`;
    cell.style.maxWidth = `${productColumnWidth}px`;
  });

  syncResultColumnWidths();

  const visualProductWidth = appliedProductColumnWidth ?? productColumnWidth;
  productCells.forEach((cell) => {
    cell.style.width = `${visualProductWidth}px`;
    cell.style.minWidth = `${visualProductWidth}px`;
    cell.style.maxWidth = `${visualProductWidth}px`;
  });
  updateFloatingResultHeader();
}

function getVisibleResultHeaders() {
  return [...resultHeaderRow.querySelectorAll("th[data-column]")]
    .filter((header) => !header.classList.contains("hidden-result-column"));
}

function syncResultColumnWidths() {
  const table = tableWrap?.querySelector("table");
  if (!table) return;
  let colgroup = table.querySelector("colgroup[data-role='result-colgroup']");

  if (productColumnWidth === null) {
    colgroup?.remove();
    appliedProductColumnWidth = null;
    appliedResultColumnWidths = [];
    table.style.removeProperty("width");
    table.style.removeProperty("min-width");
    table.style.removeProperty("table-layout");
    return;
  }

  if (!colgroup) {
    colgroup = document.createElement("colgroup");
    colgroup.dataset.role = "result-colgroup";
    table.prepend(colgroup);
  }
  colgroup.innerHTML = "";

  const headers = getVisibleResultHeaders();
  const computedWidths = headers.map((header) => getResultColumnPixelWidth(header));
  const productIndex = headers.findIndex((header) => header.dataset.column === "product_name");
  const availableWidth = tableWrap?.clientWidth || 0;
  const totalBeforeFill = computedWidths.reduce((sum, width) => sum + width, 0);
  if (!productColumnUserResized && productIndex >= 0 && availableWidth > totalBeforeFill) {
    computedWidths[productIndex] += availableWidth - totalBeforeFill;
  }
  appliedResultColumnWidths = computedWidths;
  appliedProductColumnWidth = productIndex >= 0 ? computedWidths[productIndex] : productColumnWidth;

  let totalWidth = 0;
  computedWidths.forEach((width) => {
    const col = document.createElement("col");
    col.style.width = `${width}px`;
    totalWidth += width;
    colgroup.appendChild(col);
  });
  table.style.width = `${totalWidth}px`;
  table.style.minWidth = `${totalWidth}px`;
}

function getResultColumnPixelWidth(header) {
  const column = header.dataset.column;
  if (column === "product_name") return productColumnWidth ?? Math.ceil(header.getBoundingClientRect().width);
  if (column?.startsWith("extra_") && extraFields.length === 1) return 112;

  const cells = column ? [...resultBody.querySelectorAll(`td[data-column="${cssEscape(column)}"]`)] : [];
  const cellContentWidths = cells.map((cell) => {
    const content = cell.querySelector(".copy-value") ?? cell;
    return Math.ceil(content.scrollWidth || content.getBoundingClientRect().width || 0);
  });
  const headerWidth = Math.ceil(header.scrollWidth || header.getBoundingClientRect().width || 0);
  const contentWidth = Math.max(headerWidth, ...cellContentWidths, 72);
  return contentWidth + 8;
}

function ensureFloatingResultHeader() {
  if (floatingHeader) return floatingHeader;
  floatingHeader = document.createElement("div");
  floatingHeader.className = "floating-result-header";
  floatingHeader.hidden = true;
  floatingHeaderTable = document.createElement("table");
  floatingHeader.appendChild(floatingHeaderTable);
  document.body.appendChild(floatingHeader);
  return floatingHeader;
}

function updateFloatingResultHeader() {
  if (!tableWrap || resultSection.hidden || resultHeaderRow.children.length === 0) {
    hideFloatingResultHeader();
    hideFloatingCategory();
    return;
  }
  const table = tableWrap.querySelector("table");
  if (!table) {
    hideFloatingResultHeader();
    hideFloatingCategory();
    return;
  }

  const wrapRect = tableWrap.getBoundingClientRect();
  const headerRect = resultHeaderRow.getBoundingClientRect();
  const shouldShow = wrapRect.top < 0 && wrapRect.bottom > headerRect.height + 8;
  if (!shouldShow) {
    hideFloatingResultHeader();
    hideFloatingCategory();
    return;
  }

  ensureFloatingResultHeader();
  const clonedHead = document.createElement("thead");
  const clonedRow = resultHeaderRow.cloneNode(true);
  getVisibleResultHeaders().forEach((sourceCell) => {
    const clonedCell = clonedRow.querySelector(`th[data-column="${cssEscape(sourceCell.dataset.column)}"]`);
    if (!clonedCell) return;
    const width = sourceCell.getBoundingClientRect().width;
    clonedCell.style.width = `${width}px`;
    clonedCell.style.minWidth = `${width}px`;
    clonedCell.style.maxWidth = `${width}px`;
  });
  const floatingProductResizer = clonedRow.querySelector('[data-role="product-name-resizer"]');
  if (floatingProductResizer) {
    floatingProductResizer.removeAttribute("id");
    floatingProductResizer.addEventListener("pointerdown", startProductColumnResize);
  }
  clonedHead.appendChild(clonedRow);
  floatingHeaderTable.innerHTML = "";
  floatingHeaderTable.appendChild(clonedHead);
  floatingHeaderTable.style.width = `${table.scrollWidth}px`;
  floatingHeaderTable.style.minWidth = `${table.scrollWidth}px`;
  floatingHeaderTable.style.removeProperty("table-layout");
  floatingHeaderTable.style.transform = `translateX(${-tableWrap.scrollLeft}px)`;

  floatingHeader.hidden = false;
  floatingHeader.style.left = `${wrapRect.left}px`;
  floatingHeader.style.width = `${wrapRect.width}px`;
  updateFloatingCategory(wrapRect, headerRect.height);
}

function hideFloatingResultHeader() {
  if (!floatingHeader) return;
  floatingHeader.hidden = true;
  hideFloatingCategory();
}

function ensureFloatingCategory() {
  if (floatingCategory) return floatingCategory;
  floatingCategory = document.createElement("div");
  floatingCategory.className = "floating-category-label";
  floatingCategory.hidden = true;
  floatingCategory.addEventListener("click", () => {
    const categoryKey = floatingCategory.dataset.category;
    if (!categoryKey) return;
    const shouldCollapse = !collapsedCategories.has(categoryKey);
    const nextCategoryKey = shouldCollapse ? getNextCategoryKey(categoryKey) : "";
    if (shouldCollapse) collapsedCategories.add(categoryKey);
    else collapsedCategories.delete(categoryKey);
    renderFilteredResults();
    if (shouldCollapse && nextCategoryKey) {
      requestAnimationFrame(() => scrollCategoryIntoView(nextCategoryKey));
    }
  });
  document.body.appendChild(floatingCategory);
  return floatingCategory;
}

function updateFloatingCategory(wrapRect, headerHeight) {
  if (!categoryColumn.dataset.index) {
    hideFloatingCategory();
    return;
  }
  const categoryRows = [...resultBody.querySelectorAll(".category-group-row")];
  let activeRow = null;
  categoryRows.forEach((row) => {
    const rect = row.getBoundingClientRect();
    const activationLine = headerHeight + rect.height * 0.5;
    if (rect.top <= activationLine) activeRow = row;
  });
  if (!activeRow) {
    hideFloatingCategory();
    return;
  }
  const sourceCell = activeRow.querySelector("td");
  ensureFloatingCategory();
  floatingCategory.innerHTML = sourceCell?.innerHTML ?? "";
  floatingCategory.dataset.category = activeRow.dataset.category ?? "";
  floatingCategory.hidden = false;
  floatingCategory.style.left = `${wrapRect.left}px`;
  floatingCategory.style.top = `${headerHeight}px`;
  floatingCategory.style.width = `${wrapRect.width}px`;
}

function hideFloatingCategory() {
  if (!floatingCategory) return;
  floatingCategory.hidden = true;
}

function getNextCategoryKey(categoryKey) {
  const categoryRows = [...resultBody.querySelectorAll(".category-group-row")];
  const currentIndex = categoryRows.findIndex((row) => row.dataset.category === categoryKey);
  if (currentIndex < 0) return "";
  return categoryRows[currentIndex + 1]?.dataset.category ?? "";
}

function scrollCategoryIntoView(categoryKey) {
  const nextRow = resultBody.querySelector(`.category-group-row[data-category="${cssEscape(categoryKey)}"]`);
  if (!nextRow) return;
  const headerHeight = floatingHeader?.hidden ? resultHeaderRow.getBoundingClientRect().height : floatingHeader.getBoundingClientRect().height;
  const targetTop = nextRow.getBoundingClientRect().top + window.scrollY - headerHeight;
  window.scrollTo({ top: Math.max(0, targetTop), behavior: "auto" });
  updateFloatingResultHeader();
}

function cssEscape(value) {
  if (window.CSS?.escape) return CSS.escape(value);
  return String(value).replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}


















