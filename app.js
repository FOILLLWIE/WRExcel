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
const expandResultsButton = actionButton("toggle-expand-results");
const productSearch = qs('[data-role="product-search"]');
const productFilterButton = actionButton("toggle-product-filter");
const productFilterPanel = qs('[data-role="product-filter-panel"]');
const productFilterList = qs('[data-role="product-filter-list"]');
const addProductFilterButton = actionButton("add-product-filter");
const removeParenthesesText = optionInput("removeParenthesesText");
const removeBracketsText = optionInput("removeBracketsText");
const removeTrailingModelCode = optionInput("removeTrailingModelCode");
const removeAfterDelimiter = optionInput("removeAfterDelimiter");
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
let activeExtraFieldTarget = null;
let pendingRestoreConfig = null;
let currentFileKey = "";
let settingsSaveTimer = null;
const DB_NAME = "discountCalculatorDB";
const DB_VERSION = 1;
const SETTINGS_STORE_NAME = "fileSettings";
const SETTINGS_RETENTION_DAYS = 30;
const fallbackSettingsStore = new Map();
const extraFieldNameSuggestions = ["선택쿠폰", "중복쿠폰", "카드할인"];
const rewardFieldNameSuggestions = ["\uC2A4\uB9C8\uC77C\uCE74\uB4DC", "\uBA38\uB2C8\uCDA9\uC804", "\uAF2D\uBA64\uBC841", "\uAF2D\uBA64\uBC842"];
const worksheetMergeRangeCache = new WeakMap();
const columnFilters = {
  product: { input: productColumn, picker: productColorPicker, trigger: productColorTrigger, choices: productColorChoices, selected: "" },
  original: { input: originalColumn, picker: originalColorPicker, trigger: originalColorTrigger, choices: originalColorChoices, selected: "" },
  final: { input: finalPriceColumn, picker: finalColorPicker, trigger: finalColorTrigger, choices: finalColorChoices, selected: "" },
};

if (window.location.protocol === "file:") {
  setStatus("이 페이지는 파일로 직접 열려 있습니다. start.bat을 실행해 접속해주세요.", "error");
}
let productColumnWidth = null;

initializeCustomSelects();
cleanupStoredSettings();
window.addEventListener("beforeunload", () => {
  if (uploadedFile && currentFileKey) saveCurrentSettings();
});
document.addEventListener("visibilitychange", () => {
  if (document.visibilityState === "hidden" && uploadedFile && currentFileKey) saveCurrentSettings();
});

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  uploadedFile = file;
  currentFileKey = createFileKey(file);
  fileDrop.classList.add("uploaded");
  hideResults();
  extraFields = [];
  rewardFields = [];
  renderExtraFieldInputs();
  renderRewardFieldInputs();
  setStatus("엑셀 내용을 읽고 있습니다...", "success");
  try {
    workbookPreview = await postFile("/api/inspect");
    populateSheetSelect(workbookPreview.sheets);
    renderSheet(workbookPreview.sheets[0].name);
    setElementHidden(mappingSection, false);
    const savedConfig = await loadSettingsForUploadedFile(file);
    if (savedConfig) {
      pendingRestoreConfig = savedConfig;
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
  if (!productColumn.dataset.index || !originalColumn.dataset.index || !finalPriceColumn.dataset.index) {
    setStatus("제품명 열, 기존금액 열, 최종가 열을 모두 선택해주세요.", "error");
    return;
  }
  const extraFieldValidationMessage = validateExtraFields();
  if (extraFieldValidationMessage) {
    showMessageDialog(extraFieldValidationMessage);
    return;
  }
  try {
    setCalculating(true);
    const payload = await postFile("/api/calculate", {
      sheet_name: sheetSelect.value,
      product_col: productColumn.dataset.index,
      product_color: columnFilters.product.selected,
      original_col: originalColumn.dataset.index,
      original_color: columnFilters.original.selected,
      final_price_col: finalPriceColumn.dataset.index,
      final_price_color: columnFilters.final.selected,
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
});
removeParenthesesText.addEventListener("change", renderFilteredResults);
removeBracketsText.addEventListener("change", renderFilteredResults);
removeTrailingModelCode.addEventListener("change", renderFilteredResults);
removeAfterDelimiter.addEventListener("change", () => {
  setElementHidden(delimiterOptions, !removeAfterDelimiter.checked);
  renderFilteredResults();
});
removeLeadingText.addEventListener("change", () => {
  setElementHidden(leadingTextOptions, !removeLeadingText.checked);
  renderFilteredResults();
});
leadingTextValue.addEventListener("input", renderFilteredResults);
delimiterValue.addEventListener("input", renderFilteredResults);
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
titleCaseProductName.addEventListener("change", renderFilteredResults);
autoLoadColorFilters.addEventListener("change", () => {
  if (autoLoadColorFilters.checked) loadAllColumnColors();
  else Object.keys(columnFilters).forEach(resetColumnColors);
  queueSaveCurrentSettings();
});
productNameResizer.addEventListener("pointerdown", startProductColumnResize);
editResultsButton.addEventListener("click", () => {
  isEditMode = !isEditMode;
  editResultsButton.classList.toggle("active", isEditMode);
  editResultsButton.textContent = isEditMode ? "저장하기" : "편집";
  resultSection.classList.toggle("edit-mode", isEditMode);
});
expandResultsButton.addEventListener("click", () => {
  resultSection.classList.toggle("expanded-results");
  const expanded = resultSection.classList.contains("expanded-results");
  expandResultsButton.textContent = expanded ? "축소" : "확장";
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
    const payload = await postFile("/api/column-colors", {
      sheet_name: sheetSelect.value,
      column_index: filter.input.dataset.index,
      start_row: startRow.value,
      end_row: endRow.value,
    });
    const usableColors = payload.colors.filter((color) => color.value !== "FFFFFF");
    const availableColors = usableColors.map((color) => color.value);
    filter.selected = availableColors.includes(previousSelection) ? previousSelection : "";
    filter.choices.innerHTML = `<button class="color-choice selected" type="button" data-color="">전체</button>`;
    usableColors.forEach((color) => {
      const button = document.createElement("button");
      button.className = "color-choice";
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
  setElementHidden(filter.picker, true);
  setElementHidden(filter.choices, true);
  filter.choices.innerHTML = `<button class="color-choice selected" type="button" data-color="">전체</button>`;
  updateColorTrigger(key);
}

function closeAllColorMenus(exceptChoices = null) {
  Object.values(columnFilters).forEach((filter) => {
    if (filter.choices !== exceptChoices) setElementHidden(filter.choices, true);
  });
}

function bindColorChoiceEvents(key) {
  const filter = columnFilters[key];
  filter.choices.querySelectorAll(".color-choice").forEach((button) => {
    button.addEventListener("click", () => {
      filter.selected = button.dataset.color;
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
    ? `<span class="color-swatch" style="background:#${filter.selected}"></span>`
    : "+";
  filter.trigger.title = filter.selected ? `#${filter.selected}` : "";
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
  extraFields = payload.extra_fields ?? extraFields;
  rewardFields = payload.reward_fields ?? rewardFields;
  productSearch.value = "";
  productNameFilters = [];
  renderProductFilterInputs();
  sortField.value = "none";
  sortDirection.value = "desc";
  syncCustomSelect(sortField);
  syncCustomSelect(sortDirection);
  renderFilteredResults();
  setElementHidden(resultSection, false);
}

function renderFilteredResults() {
  resultBody.innerHTML = "";

  const keyword = productSearch.value.trim().toLowerCase();
  const rows = currentRows
    .filter((row) => row.product_name.toLowerCase().includes(keyword))
    .sort((a, b) => {
      if (sortField.value === "none") return 0;
      const diff = a[sortField.value] - b[sortField.value];
      return sortDirection.value === "asc" ? diff : -diff;
    });

  const groupedRows = categoryColumn.dataset.index ? groupRowsByCategory(rows) : [["", rows]];

  groupedRows.forEach(([category, categoryRows]) => {
    if (categoryColumn.dataset.index) {
      const groupRow = document.createElement("tr");
      groupRow.className = "category-group-row";
      groupRow.innerHTML = `<td colspan="${7 + extraFields.length + rewardFields.length}">${escapeHtml(category || "미분류")}</td>`;
      resultBody.appendChild(groupRow);
    }
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
    tr.innerHTML = `
      <td>${buildCopyCell(displayProductName(row.product_name ?? ""))}</td>
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
      button.addEventListener("click", async (event) => {
        event.stopPropagation();
        if (isEditMode) {
          startInlineEdit(event.currentTarget, row);
          return;
        }
        await copyText(event.currentTarget.dataset.copyValue);
      });
    });
    tr.addEventListener("click", () => {
      selectedRowId = String(row.row_id);
      renderFilteredResults();
    });
      resultBody.appendChild(tr);
    });
  });
  applyProductColumnWidth();
  renderExtraResultHeaders();
  renderRewardResultHeaders();
  applyRewardColumnVisibility();
  applyVisibleColumns();
  applyPriceHighlighting();
  setElementHidden(emptyResultMessage, rows.length > 0);
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

function buildCopyCell(value) {
  const displayValue = escapeHtml(value);
  const copyValue = escapeAttribute(value);
  return `<button class="copy-value" type="button" data-copy-value="${copyValue}">${displayValue}</button>`;
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
      total_rows: worksheet.rowCount,
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
  let first = null;
  let last = null;
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    if (isHiddenRow(worksheet, rowNumber)) continue;
    let count = 0;
    const row = worksheet.getRow(rowNumber);
    for (let colNumber = 1; colNumber <= worksheet.columnCount; colNumber += 1) {
      if (isHiddenColumn(worksheet, colNumber - 1)) continue;
      if (cleanExcelCellValue(row.getCell(colNumber).value) !== "") count += 1;
    }
    if (count >= 2) {
      if (first === null) first = rowNumber;
      last = rowNumber;
    }
  }
  return [first ?? 1, last ?? Math.max(1, worksheet.rowCount)];
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
  const start = Math.max(Number(fields.start_row || 1), 1);
  const end = Math.min(Number(fields.end_row || worksheet.rowCount), worksheet.rowCount);
  for (let rowNumber = start; rowNumber <= end; rowNumber += 1) {
    if (isHiddenRow(worksheet, rowNumber)) continue;
    const cell = getDisplayCell(worksheet, rowNumber, columnIndex);
    if (cleanExcelCellValue(cell.value) === "") continue;
    const color = normalizeBrowserFillColor(cell);
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
  const start = Math.max(Number(fields.start_row || 1), 1);
  const end = Math.min(Number(fields.end_row || worksheet.rowCount), worksheet.rowCount);
  for (let rowNumber = start; rowNumber <= end; rowNumber += 1) {
    if (isHiddenRow(worksheet, rowNumber)) continue;
    const row = worksheet.getRow(rowNumber);
    if (fields.product_color && normalizeBrowserFillColor(getDisplayCell(worksheet, rowNumber, productCol)) !== fields.product_color) continue;
    if (fields.original_color && normalizeBrowserFillColor(getDisplayCell(worksheet, rowNumber, originalCol)) !== fields.original_color) continue;
    if (fields.final_price_color && normalizeBrowserFillColor(getDisplayCell(worksheet, rowNumber, finalCol)) !== fields.final_price_color) continue;

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
    rows.push({
      row_id: rowNumber,
      product_name: String(productValue),
      category_name: categoryCol !== null ? String(cleanExcelCellValue(getDisplayCell(worksheet, rowNumber, categoryCol).value)) : "",
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
  if (rows.length === 0) throw new Error("선택한 범위에서 계산 가능한 숫자 행을 찾지 못했습니다.");
  return { rows, extra_fields: extraFieldList, reward_fields: rewardFieldList };
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

function groupRowsByCategory(rows) {
  const groups = new Map();
  rows.forEach((row) => {
    const key = row.category_name || "미분류";
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  });
  return [...groups.entries()];
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
  if (removeTrailingModelCode.checked) {
    baseValue = baseValue.replace(/\s+[A-Z0-9]+(?:-[A-Z0-9]+)+$/i, "");
  }
  if (removeAfterDelimiter.checked && delimiterValue.value) {
    const delimiterIndex = baseValue.indexOf(delimiterValue.value);
    if (delimiterIndex >= 0) {
      baseValue = baseValue.slice(0, delimiterIndex);
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
    if (firstCell && !row.classList.contains("category-group-row")) {
      firstCell.classList.toggle("hidden-result-column", hiddenColumns.has("product_name"));
    }
  });
  applyRewardColumnVisibility();
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
        queueSaveCurrentSettings();
      });
    });
    nameSelect.addEventListener("change", () => {
      updateExtraFieldFromCard(field, card);
      updateExtraNameVisibility(card, field.name_choice);
      activateDynamicFieldColumnTarget(card);
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
        queueSaveCurrentSettings();
      });
    });
    modeSelect.addEventListener("change", () => {
      updateExtraFieldFromCard(field, card);
      updateExtraFieldVisibility(card, field.mode);
      activateDynamicFieldColumnTarget(card);
      queueSaveCurrentSettings();
    });
    card.querySelector(".extra-field-remove").addEventListener("click", () => {
      extraFields = extraFields.filter((item) => item.id !== field.id);
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

function createFileKey(file) {
  return `${file.name}_${file.size}_${file.lastModified}`;
}

function getColumnLetterFromInput(input) {
  const parsed = normalizeColumnIndexValue(input.dataset.index ?? input.value);
  return parsed === null ? "" : columnLabel(parsed);
}

function buildCurrentSettingsData() {
  if (!uploadedFile || !currentFileKey) return null;
  return {
    fileKey: currentFileKey,
    fileName: uploadedFile.name,
    savedAt: Date.now(),
    sheetName: sheetSelect.value,
    mapping: {
      product: getColumnLetterFromInput(productColumn),
      original: getColumnLetterFromInput(originalColumn),
      finalPrice: getColumnLetterFromInput(finalPriceColumn),
      category: getColumnLetterFromInput(categoryColumn),
    },
    range: {
      startRow: Number(startRow.value || 1),
      endRow: Number(endRow.value || 1),
    },
    options: {
      autoLoadColorFilters: Boolean(autoLoadColorFilters.checked),
      showWonSuffix: Boolean(showWonSuffix.checked),
      showDiscountMinus: Boolean(showDiscountMinus.checked),
      highlightFinalPrices: Boolean(highlightFinalPrices.checked),
      productColor: columnFilters.product.selected,
      originalColor: columnFilters.original.selected,
      finalPriceColor: columnFilters.final.selected,
    },
    extraFields: serializeExtraFields(),
    rewardFields: serializeRewardFields(),
  };
}

function queueSaveCurrentSettings() {
  if (!uploadedFile || !currentFileKey) return;
  window.clearTimeout(settingsSaveTimer);
  settingsSaveTimer = window.setTimeout(() => {
    saveCurrentSettings();
  }, 250);
}

async function saveCurrentSettings() {
  const data = buildCurrentSettingsData();
  if (!data) return;
  try {
    await saveFileSettings(data.fileKey, data);
  } catch (error) {
    console.warn("Failed to save file settings.", error);
  }
}

function normalizeStoredConfig(config) {
  if (!config) return null;
  if (config.mapping || config.range || config.options) {
    return {
      sheet_name: config.sheetName ?? config.sheet_name ?? sheetSelect.value,
      product_col: config.mapping?.product ?? "",
      product_color: config.options?.productColor ?? "",
      original_col: config.mapping?.original ?? "",
      original_color: config.options?.originalColor ?? "",
      final_price_col: config.mapping?.finalPrice ?? "",
      final_price_color: config.options?.finalPriceColor ?? "",
      category_col: config.mapping?.category ?? "",
      start_row: config.range?.startRow ?? "",
      end_row: config.range?.endRow ?? "",
      auto_load_color_filters: config.options?.autoLoadColorFilters,
      show_won_suffix: config.options?.showWonSuffix,
      show_discount_minus: config.options?.showDiscountMinus,
      highlight_final_prices: config.options?.highlightFinalPrices,
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
  syncTypedColumns();
  columnFilters.product.selected = normalizedConfig.product_color ?? "";
  columnFilters.original.selected = normalizedConfig.original_color ?? "";
  columnFilters.final.selected = normalizedConfig.final_price_color ?? "";
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

async function saveFileSettings(fileKey, data) {
  if (!fileKey || !data) return;
  try {
    const saved = await withSettingsStore("readwrite", (store) => {
      store.put({ ...data, fileKey });
    });
    if (saved === null) fallbackSettingsStore.set(fileKey, { ...data, fileKey });
  } catch {
    fallbackSettingsStore.set(fileKey, { ...data, fileKey });
  }
}

async function loadSettingsForUploadedFile(file) {
  const exact = await loadFileSettings(createFileKey(file));
  if (exact) return exact;
  return loadLatestFileSettingsByName(file.name);
}

async function loadLatestFileSettingsByName(fileName) {
  if (!fileName) return null;
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
    return result ?? null;
  } catch {
    let latest = null;
    fallbackSettingsStore.forEach((value) => {
      if (value?.fileName === fileName && (!latest || (value.savedAt ?? 0) > (latest.savedAt ?? 0))) latest = value;
    });
    return latest;
  }
}

async function loadFileSettings(fileKey) {
  if (!fileKey) return null;
  try {
    const result = await withSettingsStore("readonly", (store, done) => {
      const request = store.get(fileKey);
      request.onsuccess = () => done(request.result ?? null);
      request.onerror = () => done(null);
    });
    return result ?? fallbackSettingsStore.get(fileKey) ?? null;
  } catch {
    return fallbackSettingsStore.get(fileKey) ?? null;
  }
}

async function deleteFileSettings(fileKey) {
  if (!fileKey) return;
  fallbackSettingsStore.delete(fileKey);
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
        queueSaveCurrentSettings();
      });
    });
    nameSelect.addEventListener("change", () => {
      updateRewardFieldFromCard(field, card);
      updateRewardNameVisibility(card, field.name_choice);
      activateDynamicFieldColumnTarget(card);
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
        queueSaveCurrentSettings();
      });
    });
    card.querySelector(".extra-field-remove").addEventListener("click", () => {
      rewardFields = rewardFields.filter((item) => item.id !== field.id);
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
    });
    panel.insertBefore(label, totalRewardLabel);
  });
}

function startInlineEdit(button, row) {
  const cell = button.closest("td");
  const column = cell?.dataset.column ?? "product_name";
  if (!cell || column === "total_discount_amount" || column === "discount_rate" || column === "total_reward_amount" || column === "effective_price") {
    return;
  }

  const input = document.createElement("input");
  input.className = "inline-edit-input";
  input.value = getEditableValue(row, column);
  button.replaceWith(input);
  input.focus();
  input.select();

  const commit = () => {
    applyEditedValue(row, column, input.value);
    recalculateDerivedValues(row);
    renderFilteredResults();
  };

  input.addEventListener("blur", commit, { once: true });
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") input.blur();
    if (event.key === "Escape") renderFilteredResults();
  });
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
    row.extra_values[column.replace("extra_", "")] = numericValue;
  } else if (column.startsWith("reward_")) {
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
    });
    removeButton.addEventListener("click", (event) => {
      event.stopPropagation();
      productNameFilters.splice(index, 1);
      renderProductFilterInputs();
      renderFilteredResults();
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
          hideRestoreDialog();
        },
      },
      {
        label: "예",
        className: "primary-button",
        handler: () => {
          if (pendingRestoreConfig) {
            restoreSettings(pendingRestoreConfig);
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
  const startWidth = productNameHeader.getBoundingClientRect().width;

  function onPointerMove(moveEvent) {
    productColumnWidth = Math.max(180, startWidth + moveEvent.clientX - startX);
    applyProductColumnWidth();
  }

  function onPointerUp() {
    window.removeEventListener("pointermove", onPointerMove);
    window.removeEventListener("pointerup", onPointerUp);
  }

  window.addEventListener("pointermove", onPointerMove);
  window.addEventListener("pointerup", onPointerUp);
}

function applyProductColumnWidth() {
  if (productColumnWidth === null) {
    productNameHeader.style.removeProperty("width");
  } else {
    productNameHeader.style.width = `${productColumnWidth}px`;
  }
  resultBody.querySelectorAll("tr").forEach((row) => {
    if (!row.children[0]) return;
    if (productColumnWidth === null) {
      row.children[0].style.removeProperty("width");
      row.children[0].style.removeProperty("max-width");
      return;
    }
    row.children[0].style.width = `${productColumnWidth}px`;
    row.children[0].style.maxWidth = `${productColumnWidth}px`;
  });
}
