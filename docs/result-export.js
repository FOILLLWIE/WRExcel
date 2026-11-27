const exportResultsButton = actionButton("export-results");

exportResultsButton.addEventListener("click", () => {
  showAppDialog({
    type: "csv-export",
    title: "CSV 다운로드",
    message: "전체 다운로드는 현재 표시 중인 모든 행을 저장합니다. 선택 행 다운로드는 선택한 행만 저장합니다. Ctrl(또는 ⌘)을 누른 채 값을 클릭하면 여러 행을 선택·해제할 수 있습니다.",
    actions: [
      { label: "전체 다운로드", className: "primary-button csv-export-all", handler: () => { hideConfirmDialog(); downloadResultCsv(false); } },
      { label: "금액 기준 다운로드", className: "ghost-button csv-export-amount", handler: showCsvAmountFilter },
      { label: "선택 행 다운로드", className: "ghost-button csv-export-selected", handler: () => { hideConfirmDialog(); downloadResultCsv(true); } },
      { label: "×", className: "ghost-button csv-export-close", handler: hideConfirmDialog },
    ],
  });
  const close = dialogActions.querySelector(".csv-export-close");
  close.setAttribute("aria-label", "닫기");
  close.title = "닫기";
});

function getCsvResultValue(source, column, tr = null) {
  if (column === "product_name") return tr?.querySelector('[data-column="product_name"] [data-copy-value]')?.dataset.copyValue ?? source.product_name;
  if (column === "total_discount_amount") {
    const value = Number(source.total_discount_amount) || 0;
    return showDiscountMinus.checked && value > 0 ? -value : Math.abs(value);
  }
  if (column === "discount_rate") return `${Math.round(source.discount_rate)}%`;
  let value = column.startsWith("extra_") ? source.extra_values?.[column.slice(6)]
    : column.startsWith("reward_") ? source.reward_values?.[column.slice(7)]
    : column === "effective_price" ? source.effective_price ?? source.discount_amount
    : source[column];
  value = cleanExcelCellValue(value);
  if (typeof value === "string" && /^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?\s*원?$/.test(value.trim())) {
    const numeric = Number(value.replace(/[,\s원]/g, ""));
    if (Number.isFinite(numeric)) value = numeric;
  }
  return value ?? "";
}

function showCsvAmountFilter() {
  commitInlineEdit(true);
  const rowsById = new Map(currentRows.map((row) => [String(row.row_id), row]));
  const rows = [...resultBody.querySelectorAll("tr[data-row-id]")].map((tr) => rowsById.get(tr.dataset.rowId)).filter(Boolean);
  const headers = getVisibleResultHeaders().filter((header) => {
    const column = header.dataset.column;
    return column !== "product_name" && column !== "discount_rate"
      && rows.some((row) => Number.isFinite(getCsvResultValue(row, column)));
  });
  if (!headers.length) {
    showMessageDialog("금액을 선택할 수 있는 표시 항목이나 상품이 없습니다.");
    return;
  }
  const amounts = new Set();
  let column = headers[0].dataset.column;
  showAppDialog({
    type: "csv-amount-export",
    title: "금액 기준 다운로드",
    message: "현재 표시 중인 상품에서 기준 열과 금액을 선택하세요. 체크한 금액 중 하나에 해당하는 행을 다운로드합니다.",
    actions: [
      { label: "취소", className: "ghost-button", handler: hideConfirmDialog },
      { label: "다운로드", className: "primary-button", handler: () => {
        if (!amounts.size) return;
        hideConfirmDialog();
        downloadResultCsv(false, { column, amounts });
      } },
    ],
  });
  const download = dialogActions.lastElementChild;
  download.disabled = true;
  const panel = document.createElement("span");
  panel.className = "csv-amount-filter";
  const label = document.createElement("label");
  label.textContent = "기준 열 ";
  const select = document.createElement("select");
  select.dataset.role = "csv-amount-column";
  select.id = "csvAmountColumn";
  select.className = "native-select-hidden";
  select.tabIndex = -1;
  select.setAttribute("aria-hidden", "true");
  headers.forEach((header) => {
    const option = document.createElement("option");
    option.value = header.dataset.column;
    option.textContent = header.textContent.trim();
    select.appendChild(option);
  });
  const wrapper = document.createElement("span");
  wrapper.className = "custom-select";
  wrapper.dataset.selectTarget = "csvAmountColumn";
  const trigger = document.createElement("button");
  trigger.type = "button";
  trigger.className = "custom-select-trigger";
  trigger.setAttribute("aria-label", "기준 열 선택");
  const menu = document.createElement("span");
  menu.className = "custom-select-menu";
  setElementHidden(menu, true);
  wrapper.append(trigger, menu);
  label.append(wrapper, select);
  const list = document.createElement("span");
  list.className = "csv-amount-list";
  list.setAttribute("role", "group");
  list.setAttribute("aria-label", "다운로드할 금액");
  const renderAmounts = () => {
    column = select.value;
    amounts.clear();
    download.disabled = true;
    list.replaceChildren();
    const counts = new Map();
    rows.forEach((row) => {
      const value = getCsvResultValue(row, column);
      if (typeof value === "number" && Number.isFinite(value)) counts.set(value, (counts.get(value) ?? 0) + 1);
    });
    [...counts].sort(([a], [b]) => b - a).forEach(([value, count]) => {
      const option = document.createElement("label");
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = String(value);
      checkbox.addEventListener("change", () => {
        if (checkbox.checked) amounts.add(value);
        else amounts.delete(value);
        download.disabled = amounts.size === 0;
      });
      const amountLabel = document.createElement("span");
      amountLabel.textContent = `${currencyFormatter.format(value)}원`;
      const countLabel = document.createElement("span");
      countLabel.className = "csv-amount-count";
      countLabel.textContent = `${count}개`;
      option.append(checkbox, amountLabel, countLabel);
      list.appendChild(option);
    });
  };
  select.addEventListener("change", renderAmounts);
  panel.append(label, list);
  dialogMessage.appendChild(panel);
  mountCustomSelect(wrapper, select);
  renderAmounts();
}

function downloadResultCsv(selectedOnly, amountFilter = null) {
  exportResultsButton.disabled = true;
  try {
    commitInlineEdit(true);
    const headers = getVisibleResultHeaders();
    if (!headers.length || !resultBody.querySelector("tr[data-row-id]")) {
      showMessageDialog("다운로드할 상품이 없습니다. 표시 항목과 검색·카테고리 상태를 확인해주세요.");
      return;
    }
    const visibleRows = [...resultBody.querySelectorAll("tr[data-row-id]")];
    const rowsById = new Map(currentRows.map((row) => [String(row.row_id), row]));
    const exportRows = new Set(visibleRows.filter((tr) => {
      if (selectedOnly && !isResultRowSelected(JSON.parse(tr.dataset.sourceRowIds ?? JSON.stringify([tr.dataset.rowId])))) return false;
      if (!amountFilter) return true;
      const source = rowsById.get(tr.dataset.rowId);
      return source && amountFilter.amounts.has(getCsvResultValue(source, amountFilter.column));
    }));
    if (!exportRows.size) {
      showMessageDialog(amountFilter ? "선택한 금액에 해당하는 상품이 없습니다." : "다운로드할 행을 먼저 선택해주세요. Ctrl을 누른 채 클릭하면 여러 행을 선택할 수 있습니다.");
      return;
    }
    const csvRows = [headers.map((header) => header.textContent.trim())];
    let pendingCategory = null;
    for (const tr of resultBody.querySelectorAll("tr")) {
      if (tr.classList.contains("category-group-row")) {
        if (selectedOnly || amountFilter) pendingCategory = tr.dataset.category;
        else csvRows.push([tr.dataset.category, ...Array(headers.length - 1).fill("")]);
        continue;
      }
      if (!exportRows.has(tr)) continue;
      if (pendingCategory !== null) {
        csvRows.push([pendingCategory, ...Array(headers.length - 1).fill("")]);
        pendingCategory = null;
      }
      const source = rowsById.get(tr.dataset.rowId);
      if (!source) continue;
      const values = headers.map((header) => {
        const column = header.dataset.column;
        return getCsvResultValue(source, column, tr);
      });
      csvRows.push(values);
    }
    const csv = "\uFEFF" + csvRows.map((row) => row.map((value) => {
      if (typeof value === "number") return `"${currencyFormatter.format(value)}"`;
      let text = String(value ?? "");
      // CSV text must not become a spreadsheet formula when the file is opened.
      if (/^[\s]*[=+@-]/.test(text)) text = "'" + text;
      return `"${text.replace(/"/g, '\"\"')}"`;
    }).join(",")).join("\r\n");
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    const link = document.createElement("a");
    link.href = url;
    const name = (uploadedFile?.name ?? "WRExcel").replace(/\.[^.]+$/, "").replace(/[<>:"/\\|?*\x00-\x1f]/g, "_");
    link.download = `${name}_${amountFilter ? "금액기준_" : selectedOnly ? "선택행_" : ""}계산결과.csv`;
    document.body.appendChild(link);
    try { link.click(); }
    finally { link.remove(); window.setTimeout(() => URL.revokeObjectURL(url), 1000); }
  } catch (error) {
    showMessageDialog(`CSV 파일을 만들지 못했습니다. ${error.message}`);
  } finally {
    exportResultsButton.disabled = false;
  }
}
