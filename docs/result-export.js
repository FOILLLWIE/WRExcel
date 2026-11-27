const exportResultsButton = actionButton("export-results");

exportResultsButton.addEventListener("click", () => {
  exportResultsButton.disabled = true;
  try {
    commitInlineEdit(true);
    const headers = getVisibleResultHeaders();
    if (!headers.length || !resultBody.querySelector("tr[data-row-id]")) {
      showMessageDialog("다운로드할 상품이 없습니다. 표시 항목과 검색·카테고리 상태를 확인해주세요.");
      return;
    }
    const csvRows = [headers.map((header) => header.textContent.trim())];
    const rowsById = new Map(currentRows.map((row) => [String(row.row_id), row]));
    for (const tr of resultBody.querySelectorAll("tr")) {
      if (tr.classList.contains("category-group-row")) {
        csvRows.push([tr.dataset.category, ...Array(headers.length - 1).fill("")]);
        continue;
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
      if (typeof value === "number") return String(value);
      let text = String(value ?? "");
      // CSV text must not become a spreadsheet formula when the file is opened.
      if (/^[\s]*[=+@-]/.test(text)) text = "'" + text;
      return `"${text.replace(/"/g, '\"\"')}"`;
    }).join(",")).join("\r\n");
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    const link = document.createElement("a");
    link.href = url;
    const name = (uploadedFile?.name ?? "WRExcel").replace(/\.[^.]+$/, "").replace(/[<>:"/\\|?*\x00-\x1f]/g, "_");
    link.download = `${name}_계산결과.csv`;
    document.body.appendChild(link);
    try { link.click(); }
    finally { link.remove(); window.setTimeout(() => URL.revokeObjectURL(url), 1000); }
  } catch (error) {
    showMessageDialog(`CSV 파일을 만들지 못했습니다. ${error.message}`);
  } finally {
    exportResultsButton.disabled = false;
  }
});
