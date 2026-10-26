const ocrPanel = document.querySelector('[data-role="image-import"]');
const ocrElement = (role) => ocrPanel.querySelector(`[data-role="${role}"]`);
const ocrState = { file: null, url: "", rows: [], busy: false, run: 0, cancel: null };
let ocrEnginePromise = null;

ocrElement("image-file").addEventListener("change", (event) => {
  const file = event.target.files?.[0];
  if (ocrState.url) URL.revokeObjectURL(ocrState.url);
  ocrState.file = null;
  ocrState.rows = [];
  ocrElement("ocr-original").hidden = true;
  ocrElement("ocr-editor").hidden = true;
  ocrElement("ocr-start").disabled = true;
  ocrSetStatus("");
  if (!file) return;
  if (!/image\/(png|jpeg|webp)$/.test(file.type) || file.size > 20 * 1024 * 1024) {
    ocrSetStatus("20MB 이하의 PNG·JPG·WebP 이미지를 선택해주세요.", true);
    return;
  }
  ocrState.file = file;
  ocrState.url = URL.createObjectURL(file);
  ocrElement("ocr-original").src = ocrState.url;
  ocrElement("ocr-original").hidden = false;
  ocrElement("ocr-start").disabled = false;
});

ocrElement("ocr-start").addEventListener("click", ocrRecognize);
ocrElement("ocr-cancel").addEventListener("click", () => {
  const error = new Error("인식을 취소했습니다.");
  error.name = "AbortError";
  ocrState.cancel?.(error);
});
ocrElement("ocr-confirm").addEventListener("change", ocrUpdateApply);
ocrElement("ocr-add-row").addEventListener("click", () => {
  if (ocrState.rows.length >= 200) return;
  ocrState.rows.push(Array.from({ length: ocrState.rows[0].length }, () => ({ text: "", confidence: 100 })));
  ocrRenderTable();
});
ocrElement("ocr-add-column").addEventListener("click", () => {
  if (ocrState.rows[0].length >= 24) return;
  ocrState.rows.forEach((row) => row.push({ text: "", confidence: 100 }));
  ocrRenderTable();
});
ocrElement("ocr-apply").addEventListener("click", async () => {
  if (!ocrElement("ocr-confirm").checked || ocrState.busy) return;
  const rows = ocrState.rows.map((row) => row.map((cell) => cell.text.trim())).filter((row) => row.some(Boolean));
  if (!rows.length) {
    ocrSetStatus("불러올 내용이 없습니다. 표에 제품명과 금액을 입력해주세요.", true);
    return;
  }
  ocrSetBusy(true);
  ocrElement("ocr-cancel").hidden = true;
  ocrSetStatus("확인한 표를 준비하고 있습니다…");
  try {
    const workbook = new ExcelJS.Workbook();
    workbook.created = new Date(0);
    workbook.modified = new Date(0);
    workbook.addWorksheet("이미지 인식").addRows(rows);
    const buffer = await workbook.xlsx.writeBuffer();
    const file = new File([buffer], `${ocrState.file.name.replace(/\.[^.]+$/, "")}_이미지인식.xlsx`, {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      lastModified: ocrState.file.lastModified,
    });
    autoLoadColorFilters.checked = false;
    if (await loadWorkbookFile(file, { sourceLabel: "이미지 인식 표" })) {
      ocrPanel.querySelector("details").open = false;
      mappingSection.scrollIntoView({ behavior: "smooth", block: "start" });
      ocrSetStatus("표를 불러왔습니다. 계산할 열을 선택해주세요.");
    } else ocrSetStatus("표를 불러오지 못했습니다. 내용을 확인하고 다시 시도해주세요.", true);
  } catch {
    ocrSetStatus("표를 불러오지 못했습니다. 내용을 확인하고 다시 시도해주세요.", true);
  } finally {
    ocrSetBusy(false);
  }
});

function ocrSetStatus(message, error = false) {
  ocrElement("ocr-status").textContent = message;
  ocrElement("ocr-status").classList.toggle("error", error);
}

function ocrSetBusy(busy) {
  ocrState.busy = busy;
  ocrElement("image-file").disabled = busy;
  ocrElement("ocr-start").disabled = busy || !ocrState.file;
  ocrElement("ocr-start").textContent = busy ? "처리 중…" : "글자 인식";
  ocrElement("ocr-cancel").hidden = !busy;
  ocrElement("ocr-progress").hidden = !busy;
  ocrElement("ocr-editor").querySelectorAll("input, button").forEach((element) => { element.disabled = busy; });
  ocrElement("ocr-editor").querySelectorAll("[data-remove-row]").forEach((element) => { element.disabled = busy || ocrState.rows.length <= 1; });
  ocrElement("ocr-editor").querySelectorAll("[data-remove-column]").forEach((element) => { element.disabled = busy || ocrState.rows[0]?.length <= 1; });
  ocrElement("ocr-add-row").disabled = busy || ocrState.rows.length >= 200;
  ocrElement("ocr-add-column").disabled = busy || ocrState.rows[0]?.length >= 24;
  ocrUpdateApply();
}

function ocrUpdateApply() {
  ocrElement("ocr-apply").disabled = ocrState.busy || !ocrElement("ocr-confirm").checked || !ocrState.rows.length;
}

function ocrLoadEngine() {
  if (window.Tesseract) return Promise.resolve(window.Tesseract);
  if (!ocrEnginePromise) {
    ocrEnginePromise = new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = new URL("./vendor/ocr/tesseract.min.js", document.baseURI).href;
      script.onload = () => resolve(window.Tesseract);
      script.onerror = () => { script.remove(); reject(new Error("인식 모듈을 불러오지 못했습니다. 새로고침 후 다시 시도해주세요.")); };
      document.head.appendChild(script);
    }).catch((error) => { ocrEnginePromise = null; throw error; });
  }
  return ocrEnginePromise;
}

async function ocrRecognize() {
  if (!ocrState.file || ocrState.busy) return;
  const run = ++ocrState.run;
  ocrSetBusy(true);
  ocrElement("ocr-editor").hidden = true;
  ocrElement("ocr-confirm").checked = false;
  ocrElement("ocr-progress").value = 0;
  ocrSetStatus("인식기 준비 중… 최초 실행은 모델 다운로드로 시간이 걸릴 수 있습니다.");
  let worker;
  let bitmap;
  const canceled = new Promise((_, reject) => { ocrState.cancel = reject; });
  canceled.catch(() => {});
  const timer = window.setTimeout(() => ocrState.cancel?.(new Error("인식 시간이 초과되었습니다. 표 영역만 잘라서 다시 시도해주세요.")), 180000);
  try {
    bitmap = await createImageBitmap(ocrState.file);
    if (bitmap.width * bitmap.height > 20000000) throw new Error("이미지가 너무 큽니다. 표 영역만 잘라서 올려주세요.");
    const scale = Math.min(bitmap.width < 1600 ? 2 : 1, 3200 / Math.max(bitmap.width, bitmap.height));
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(bitmap.width * scale);
    canvas.height = Math.round(bitmap.height * scale);
    const context = canvas.getContext("2d", { willReadFrequently: true });
    context.fillStyle = "white";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
    const engine = await Promise.race([ocrLoadEngine(), canceled]);
    const pendingWorker = engine.createWorker(["kor", "eng"], 1, {
      workerPath: new URL("./vendor/ocr/worker.min.js", document.baseURI).href,
      corePath: "https://cdn.jsdelivr.net/npm/tesseract.js-core@7.0.0",
      errorHandler: () => {},
      logger: ({ status, progress }) => {
        if (!ocrState.busy || ocrState.run !== run) return;
        ocrElement("ocr-progress").value = Math.round((progress || 0) * 100);
        const stage = status === "recognizing text" ? "글자 인식 중" : status === "loading language traineddata" ? "한글·영문 모델 불러오는 중" : "인식기 준비 중";
        ocrSetStatus(`${stage}… ${Math.round((progress || 0) * 100)}%`);
      },
    }).then((created) => {
      if (!ocrState.busy || ocrState.run !== run) { created.terminate(); return created; }
      return created;
    });
    worker = await Promise.race([pendingWorker, canceled]);
    await worker.setParameters({ tessedit_pageseg_mode: "11", preserve_interword_spaces: "1", user_defined_dpi: "150" });
    const { data } = await Promise.race([worker.recognize(canvas, {}, { tsv: true }), canceled]);
    const words = ocrParseTsv(data.tsv || "");
    if (!words.length) throw new Error("읽을 수 있는 글자가 없습니다. 선명한 표 캡처로 다시 시도해주세요.");
    ocrState.rows = ocrWordsToTable(words, canvas.width, ocrFindGridLines(canvas, true), ocrFindGridLines(canvas, false));
    if (ocrState.rows.length > 200 || ocrState.rows[0].length > 24) throw new Error("한 번에 200행·24열까지 인식할 수 있습니다. 표를 나누어 올려주세요.");
    ocrRenderTable();
    ocrElement("ocr-editor").hidden = false;
    ocrSetStatus(`${ocrState.rows.length}행 · ${ocrState.rows[0].length}열을 인식했습니다. 원본과 비교한 뒤 확인해주세요.`);
  } catch (error) {
    ocrSetStatus(error instanceof Error && /[가-힣]/.test(error.message) ? error.message : "글자 인식에 실패했습니다. 인터넷 연결을 확인한 뒤 다시 시도해주세요.", error?.name !== "AbortError");
  } finally {
    window.clearTimeout(timer);
    ocrState.cancel = null;
    if (worker) await worker.terminate().catch(() => {});
    bitmap?.close();
    ocrSetBusy(false);
  }
}

function ocrParseTsv(tsv) {
  return tsv.trim().split(/\r?\n/).slice(1).flatMap((line) => {
    const fields = line.split("\t");
    const [x, y, width, height, confidence] = fields.slice(6, 11).map(Number);
    const text = fields.slice(11).join("\t").trim();
    return fields[0] === "5" && text && [x, y, width, height, confidence].every(Number.isFinite)
      ? [{ x, y, width, height, confidence, text }] : [];
  });
}

function ocrFindGridLines(canvas, vertical) {
  const { width, height } = canvas;
  const pixels = canvas.getContext("2d").getImageData(0, 0, width, height).data;
  const length = vertical ? width : height;
  const cross = vertical ? height : width;
  const lines = [];
  let start = -1;
  for (let position = 0; position <= length; position += 1) {
    let dark = 0;
    if (position < length) for (let offset = 0; offset < cross; offset += 2) {
      const index = (vertical ? offset * width + position : position * width + offset) * 4;
      const max = Math.max(pixels[index], pixels[index + 1], pixels[index + 2]);
      const min = Math.min(pixels[index], pixels[index + 1], pixels[index + 2]);
      if (max < 235 && max - min < 40) dark += 1;
    }
    const isLine = dark / Math.ceil(cross / 2) > 0.55;
    if (isLine && start < 0) start = position;
    if (!isLine && start >= 0) {
      if (position - start <= Math.max(6, length * 0.01)) lines.push((start + position - 1) / 2);
      start = -1;
    }
  }
  return lines;
}

function ocrWordsToTable(words, width, vertical = [], horizontal = []) {
  const heights = words.map((word) => word.height).sort((a, b) => a - b);
  const height = heights[Math.floor(heights.length / 2)];
  const lines = [];
  [...words].sort((a, b) => a.y + a.height / 2 - b.y - b.height / 2).forEach((word) => {
    const center = word.y + word.height / 2;
    const band = horizontal.filter((position) => position < center).length;
    let line = lines.find((item) => horizontal.length >= 2 ? item.band === band : Math.abs(item.center - center) <= height * 0.65);
    if (!line) { line = { band, center, words: [] }; lines.push(line); }
    line.words.push(word);
  });
  const left = Math.min(...words.map((word) => word.x));
  const right = Math.max(...words.map((word) => word.x + word.width));
  let dividers = vertical.filter((x) => x > left && x < right);
  if (!dividers.length) {
    const segments = lines.flatMap((line) => {
      const groups = [];
      [...line.words].sort((a, b) => a.x - b.x).forEach((word) => {
        const last = groups[groups.length - 1];
        if (last && word.x - last.end < height * 1.2) last.end = Math.max(last.end, word.x + word.width);
        else groups.push({ start: word.x, end: word.x + word.width });
      });
      return groups.length > 1 ? groups : [];
    }).sort((a, b) => a.start - b.start);
    const columns = [];
    segments.forEach((segment) => {
      const last = columns[columns.length - 1];
      if (last && segment.start <= last.end + height * 0.5) last.end = Math.max(last.end, segment.end);
      else columns.push({ ...segment });
    });
    dividers = columns.slice(1).map((column, index) => (columns[index].end + column.start) / 2);
  }
  return lines.map((line) => {
    const cells = Array.from({ length: dividers.length + 1 }, () => ({ text: "", confidence: 100 }));
    line.words.sort((a, b) => a.x - b.x).forEach((word) => {
      const column = dividers.filter((divider) => divider < word.x + word.width / 2).length;
      const cell = cells[column];
      cell.text += `${cell.text ? " " : ""}${word.text}`;
      cell.confidence = Math.min(cell.confidence, word.confidence);
    });
    return cells;
  });
}

function ocrRenderTable() {
  const table = ocrElement("ocr-table");
  table.innerHTML = "";
  ocrElement("ocr-confirm").checked = false;
  ocrUpdateApply();
  const header = table.createTHead().insertRow();
  header.appendChild(document.createElement("th"));
  ocrState.rows[0].forEach((_, index) => {
    const cell = document.createElement("th");
    const remove = document.createElement("button");
    remove.type = "button";
    remove.dataset.removeColumn = String(index);
    remove.textContent = `${columnLabel(index)} ×`;
    remove.setAttribute("aria-label", `${columnLabel(index)}열 삭제`);
    remove.disabled = ocrState.rows[0].length === 1;
    remove.addEventListener("click", () => { if (ocrState.rows[0].length <= 1) return; ocrState.rows.forEach((row) => row.splice(index, 1)); ocrRenderTable(); });
    cell.appendChild(remove);
    header.appendChild(cell);
  });
  const body = table.createTBody();
  ocrState.rows.forEach((row, rowIndex) => {
    const tr = body.insertRow();
    const heading = document.createElement("th");
    const remove = document.createElement("button");
    remove.type = "button";
    remove.dataset.removeRow = String(rowIndex);
    remove.textContent = `${rowIndex + 1} ×`;
    remove.setAttribute("aria-label", `${rowIndex + 1}행 삭제`);
    remove.disabled = ocrState.rows.length === 1;
    remove.addEventListener("click", () => { if (ocrState.rows.length <= 1) return; ocrState.rows.splice(rowIndex, 1); ocrRenderTable(); });
    heading.appendChild(remove);
    tr.appendChild(heading);
    row.forEach((cell, columnIndex) => {
      const td = tr.insertCell();
      const input = document.createElement("input");
      input.type = "text";
      input.value = cell.text;
      input.setAttribute("aria-label", `${columnLabel(columnIndex)}${rowIndex + 1}`);
      input.classList.toggle("ocr-uncertain", cell.confidence < 75);
      input.addEventListener("input", () => {
        cell.text = input.value;
        cell.confidence = 100;
        input.classList.remove("ocr-uncertain");
        ocrElement("ocr-confirm").checked = false;
        ocrUpdateApply();
      });
      td.appendChild(input);
    });
  });
  ocrElement("ocr-add-row").disabled = ocrState.busy || ocrState.rows.length >= 200;
  ocrElement("ocr-add-column").disabled = ocrState.busy || ocrState.rows[0].length >= 24;
}
