const exportResultsButton = actionButton("export-results");

exportResultsButton.addEventListener("click", async () => {
  exportResultsButton.disabled = true;
  try {
    commitInlineEdit(true);
    const headers = getVisibleResultHeaders();
    if (!headers.length || !resultBody.querySelector("tr[data-row-id]")) {
      showMessageDialog("다운로드할 상품이 없습니다. 표시 항목과 검색·카테고리 상태를 확인해주세요.");
      return;
    }
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("계산 결과");
    sheet.views = [{ state: "frozen", ySplit: 1 }];
    sheet.addRow(headers.map((header) => header.textContent.trim()));
    const rowsById = new Map(currentRows.map((row) => [String(row.row_id), row]));
    for (const tr of resultBody.querySelectorAll("tr")) {
      if (tr.classList.contains("category-group-row")) {
        const categoryRow = sheet.addRow([tr.dataset.category]);
        if (headers.length > 1) sheet.mergeCells(categoryRow.number, 1, categoryRow.number, headers.length);
        categoryRow.getCell(1).font = { bold: true };
        categoryRow.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF4EFE8" } };
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
        if (column === "discount_rate") return Math.round(source.discount_rate) / 100;
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
      const row = sheet.addRow(values);
      headers.forEach((header, index) => {
        const cell = row.getCell(index + 1);
        if (typeof cell.value !== "number") return;
        cell.numFmt = header.dataset.column === "discount_rate" ? "0%" : showWonSuffix.checked ? '#,##0.###"원"' : "#,##0.###";
      });
    }
    headers.forEach((header, index) => {
      sheet.getColumn(index + 1).width = header.dataset.column === "product_name" ? 55 : Math.max(16, Math.min(30, header.textContent.trim().length * 2 + 4));
      sheet.getRow(1).getCell(index + 1).font = { bold: true };
    });
    const buffer = await workbook.xlsx.writeBuffer();
    const url = URL.createObjectURL(new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const link = document.createElement("a");
    link.href = url;
    const name = (uploadedFile?.name ?? "WRExcel").replace(/\.[^.]+$/, "").replace(/[<>:"/\\|?*\x00-\x1f]/g, "_");
    link.download = `${name}_계산결과.xlsx`;
    document.body.appendChild(link);
    try { link.click(); }
    finally { link.remove(); window.setTimeout(() => URL.revokeObjectURL(url), 1000); }
  } catch (error) {
    showMessageDialog(`엑셀 파일을 만들지 못했습니다. ${error.message}`);
  } finally {
    exportResultsButton.disabled = false;
  }
});
