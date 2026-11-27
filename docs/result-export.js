const exportResultsButton = actionButton("export-results");

exportResultsButton.addEventListener("click", () => {
  showAppDialog({
    type: "csv-export",
    title: "CSV 다운로드",
    message: "전체 다운로드는 현재 표시 중인 모든 행을 저장합니다. 선택 행 다운로드는 선택한 행만 저장합니다. Ctrl(또는 ⌘)을 누른 채 값을 클릭하면 여러 행을 선택·해제할 수 있습니다.",
    actions: [
      { label: "취소", className: "ghost-button", handler: hideConfirmDialog },
      { label: "전체 다운로드", className: "ghost-button", handler: () => { hideConfirmDialog(); downloadResultCsv(false); } },
      { label: "선택 행 다운로드", className: "primary-button", handler: () => { hideConfirmDialog(); downloadResultCsv(true); } },
    ],
  });
});

function downloadResultCsv(selectedOnly) {
  exportResultsButton.disabled = true;
  try {
    commitInlineEdit(true);
    const headers = getVisibleResultHeaders();
    if (!headers.length || !resultBody.querySelector("tr[data-row-id]")) {
      showMessageDialog("다운로드할 상품이 없습니다. 표시 항목과 검색·카테고리 상태를 확인해주세요.");
      return;
    }
    const visibleRows = [...resultBody.querySelectorAll("tr[data-row-id]")];
    const exportRows = new Set(visibleRows.filter((tr) => !selectedOnly || isResultRowSelected(JSON.parse(tr.dataset.sourceRowIds ?? JSON.stringify([tr.dataset.rowId])))));
    if (!exportRows.size) {
      showMessageDialog("다운로드할 행을 먼저 선택해주세요. Ctrl을 누른 채 클릭하면 여러 행을 선택할 수 있습니다.");
      return;
    }
    const csvRows = [headers.map((header) => header.textContent.trim())];
    const rowsById = new Map(currentRows.map((row) => [String(row.row_id), row]));
    let pendingCategory = null;
    for (const tr of resultBody.querySelectorAll("tr")) {
      if (tr.classList.contains("category-group-row")) {
        if (selectedOnly) pendingCategory = tr.dataset.category;
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
        if (column === "product_name") return tr.querySelector('[data-column="product_name"] [data-copy-value]')?.dataset.copyValue ?? source.product_name;
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
    link.download = `${name}_${selectedOnly ? "선택행_" : ""}계산결과.csv`;
    document.body.appendChild(link);
    try { link.click(); }
    finally { link.remove(); window.setTimeout(() => URL.revokeObjectURL(url), 1000); }
  } catch (error) {
    showMessageDialog(`CSV 파일을 만들지 못했습니다. ${error.message}`);
  } finally {
    exportResultsButton.disabled = false;
  }
}
