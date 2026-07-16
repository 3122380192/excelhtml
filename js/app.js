/**
 * HTML Excel — main application
 */
(function () {
  "use strict";

  const DEFAULT_ROWS = 60;
  const DEFAULT_COLS = 18;
  const MIN_ROWS = 25;
  const MIN_COLS = 12;
  const DEFAULT_COL_W = 110;
  const ROW_H = 26;
  const CORNER_W = 48;
  const MAX_UNDO = 80;

  // ── State ───────────────────────────────────────────────
  let data = [];
  let rows = DEFAULT_ROWS;
  let cols = DEFAULT_COLS;
  let colWidths = [];
  let fileName = "";
  let dirty = false;
  let freezeHeader = false;

  let activeR = 0,
    activeC = 0;
  let selR1 = 0,
    selC1 = 0,
    selR2 = 0,
    selC2 = 0;
  let selecting = false;
  let editing = false;
  let editEl = null;

  let clipText = "";
  let clipIsCut = false;
  let clipRange = null;

  let findHits = [];
  let findIdx = -1;

  const undoStack = [];
  const redoStack = [];

  // fill drag
  let filling = false;
  let fillStart = null;

  // col resize
  let resizingCol = null;
  let resizeStartX = 0;
  let resizeStartW = 0;

  // ── DOM ─────────────────────────────────────────────────
  const sheet = document.getElementById("sheet");
  const sheetWrap = document.getElementById("sheetWrap");
  const selOverlay = document.getElementById("selOverlay");
  const fillHandle = document.getElementById("fillHandle");
  const formulaInput = document.getElementById("formulaInput");
  const cellRef = document.getElementById("cellRef");
  const fileInput = document.getElementById("fileInput");
  const fileNameEl = document.getElementById("fileName");
  const toastEl = document.getElementById("toast");
  const hint = document.getElementById("hint");
  const findBar = document.getElementById("findBar");
  const findInput = document.getElementById("findInput");
  const replaceInput = document.getElementById("replaceInput");
  const findCount = document.getElementById("findCount");
  const statusReady = document.getElementById("statusReady");
  const statusStats = document.getElementById("statusStats");
  const helpModal = document.getElementById("helpModal");

  const cellMap = new Map();
  const colHeaderMap = new Map();
  const rowHeaderMap = new Map();

  // ── Utils ───────────────────────────────────────────────
  function colToLetter(n) {
    let s = "";
    n += 1;
    while (n > 0) {
      n--;
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26);
    }
    return s;
  }

  function cellAddr(r, c) {
    return colToLetter(c) + (r + 1);
  }

  function normalizeRange(r1, c1, r2, c2) {
    return {
      r1: Math.min(r1, r2),
      c1: Math.min(c1, c2),
      r2: Math.max(r1, r2),
      c2: Math.max(c1, c2),
    };
  }

  function getRange() {
    return normalizeRange(selR1, selC1, selR2, selC2);
  }

  function toast(msg, ms) {
    toastEl.textContent = msg;
    toastEl.classList.add("show");
    clearTimeout(toast._t);
    toast._t = setTimeout(() => toastEl.classList.remove("show"), ms || 2200);
  }

  function markDirty() {
    dirty = true;
    updateStatus();
  }

  function isNumeric(s) {
    return Formulas.parseNumber(s) !== null && String(s).trim() !== "";
  }

  function cloneData() {
    return data.map((row) => row.slice());
  }

  function pushUndo() {
    undoStack.push({
      data: cloneData(),
      rows,
      cols,
      colWidths: colWidths.slice(),
      activeR,
      activeC,
      selR1,
      selC1,
      selR2,
      selC2,
    });
    if (undoStack.length > MAX_UNDO) undoStack.shift();
    redoStack.length = 0;
    updateUndoButtons();
  }

  function restoreSnapshot(snap) {
    data = snap.data.map((row) => row.slice());
    rows = snap.rows;
    cols = snap.cols;
    colWidths = snap.colWidths.slice();
    activeR = snap.activeR;
    activeC = snap.activeC;
    selR1 = snap.selR1;
    selC1 = snap.selC1;
    selR2 = snap.selR2;
    selC2 = snap.selC2;
    renderSheet();
    paintSelection();
    updateFormulaBar();
    markDirty();
  }

  function undo() {
    if (!undoStack.length) {
      toast("Không còn thao tác để hoàn tác");
      return;
    }
    commitEdit();
    redoStack.push({
      data: cloneData(),
      rows,
      cols,
      colWidths: colWidths.slice(),
      activeR,
      activeC,
      selR1,
      selC1,
      selR2,
      selC2,
    });
    restoreSnapshot(undoStack.pop());
    dirty = true;
    updateStatus();
    updateUndoButtons();
    toast("Hoàn tác");
  }

  function redo() {
    if (!redoStack.length) {
      toast("Không còn thao tác để làm lại");
      return;
    }
    commitEdit();
    undoStack.push({
      data: cloneData(),
      rows,
      cols,
      colWidths: colWidths.slice(),
      activeR,
      activeC,
      selR1,
      selC1,
      selR2,
      selC2,
    });
    restoreSnapshot(redoStack.pop());
    dirty = true;
    updateStatus();
    updateUndoButtons();
    toast("Làm lại");
  }

  function updateUndoButtons() {
    const bu = document.getElementById("btnUndo");
    const br = document.getElementById("btnRedo");
    if (bu) bu.disabled = !undoStack.length;
    if (br) br.disabled = !redoStack.length;
  }

  function ensureSize(needR, needC) {
    let changed = false;
    if (needR >= rows) {
      const nr = needR + 15;
      for (let r = rows; r < nr; r++) data[r] = new Array(cols).fill("");
      rows = nr;
      changed = true;
    }
    if (needC >= cols) {
      const nc = needC + 5;
      for (let r = 0; r < rows; r++) {
        while (data[r].length < nc) data[r].push("");
      }
      while (colWidths.length < nc) colWidths.push(DEFAULT_COL_W);
      cols = nc;
      changed = true;
    }
    if (changed) renderSheet();
  }

  function getRaw(r, c) {
    if (r < 0 || c < 0 || r >= rows || c >= cols) return "";
    return data[r][c] ?? "";
  }

  function getCell(r, c) {
    return getRaw(r, c);
  }

  function getDisplay(r, c) {
    return Formulas.evaluate(getRaw(r, c), getRaw);
  }

  function setCell(r, c, val) {
    ensureSize(r, c);
    data[r][c] = val == null ? "" : String(val);
  }

  // ── Load / Save ─────────────────────────────────────────
  function loadGrid(grid, name) {
    if (!grid.length) grid = [[""]];
    const maxC = Math.max(...grid.map((r) => r.length), MIN_COLS);
    rows = Math.max(grid.length + 8, MIN_ROWS, DEFAULT_ROWS);
    cols = Math.max(maxC + 3, MIN_COLS, DEFAULT_COLS);
    data = [];
    for (let r = 0; r < rows; r++) {
      data[r] = [];
      for (let c = 0; c < cols; c++) {
        data[r][c] = grid[r] && grid[r][c] != null ? String(grid[r][c]) : "";
      }
    }
    colWidths = Array.from({ length: cols }, () => DEFAULT_COL_W);
    fileName = name || "untitled.csv";
    dirty = false;
    undoStack.length = 0;
    redoStack.length = 0;
    activeR = activeC = selR1 = selC1 = selR2 = selC2 = 0;
    hint.classList.remove("show");
    renderSheet();
    setSelection(0, 0, 0, 0);
    updateStatus();
    updateUndoButtons();
    toast(`Đã mở: ${fileName} (${grid.length} hàng × ${maxC} cột)`);
  }

  function newSheet() {
    if (dirty && !confirm("Dữ liệu chưa lưu sẽ mất. Tạo sheet mới?")) return;
    rows = DEFAULT_ROWS;
    cols = DEFAULT_COLS;
    data = Array.from({ length: rows }, () => Array(cols).fill(""));
    colWidths = Array.from({ length: cols }, () => DEFAULT_COL_W);
    fileName = "untitled.csv";
    dirty = false;
    undoStack.length = 0;
    redoStack.length = 0;
    hint.classList.remove("show");
    renderSheet();
    setSelection(0, 0, 0, 0);
    updateStatus();
    updateUndoButtons();
    toast("Sheet trống mới");
  }

  function openFile(file) {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        loadGrid(CSV.parseCSV(String(reader.result)), file.name);
      } catch (e) {
        toast("Lỗi đọc CSV: " + e.message);
        console.error(e);
      }
    };
    reader.onerror = () => toast("Không đọc được file");
    reader.readAsText(file, "UTF-8");
  }

  function usedBounds() {
    let maxR = 0,
      maxC = 0;
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        if (getRaw(r, c) !== "") {
          maxR = Math.max(maxR, r);
          maxC = Math.max(maxC, c);
        }
      }
    }
    return { maxR, maxC };
  }

  function saveCSV() {
    const { maxR, maxC } = usedBounds();
    const exportRows = [];
    for (let r = 0; r <= maxR; r++) {
      const row = [];
      for (let c = 0; c <= maxC; c++) row.push(getRaw(r, c));
      exportRows.push(row);
    }
    if (!exportRows.length) exportRows.push([""]);

    const csv = CSV.toCSV(exportRows);
    const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = fileName || "export.csv";
    a.click();
    URL.revokeObjectURL(a.href);
    dirty = false;
    updateStatus();
    toast("Đã lưu: " + (fileName || "export.csv"));
  }

  // ── Render ──────────────────────────────────────────────
  function applyColTemplate() {
    const parts = colWidths.map((w) => w + "px");
    sheet.style.gridTemplateColumns = `${CORNER_W}px ${parts.join(" ")}`;
    sheet.style.gridTemplateRows = `${ROW_H}px repeat(${rows}, ${ROW_H}px)`;
  }

  function renderSheet() {
    cellMap.clear();
    colHeaderMap.clear();
    rowHeaderMap.clear();
    sheet.innerHTML = "";
    sheet.classList.toggle("freeze-header", freezeHeader);

    while (colWidths.length < cols) colWidths.push(DEFAULT_COL_W);
    applyColTemplate();

    const corner = document.createElement("div");
    corner.className = "corner";
    corner.style.gridColumn = "1";
    corner.style.gridRow = "1";
    sheet.appendChild(corner);

    for (let c = 0; c < cols; c++) {
      const h = document.createElement("div");
      h.className = "col-header" + (freezeHeader ? " frozen" : "");
      h.textContent = colToLetter(c);
      h.dataset.col = String(c);
      h.style.gridColumn = String(c + 2);
      h.style.gridRow = "1";
      h.style.position = "sticky";

      const resizer = document.createElement("div");
      resizer.className = "col-resize";
      resizer.dataset.col = String(c);
      h.appendChild(resizer);

      h.addEventListener("mousedown", (e) => {
        if (e.target.classList.contains("col-resize")) return;
        e.preventDefault();
        commitEdit();
        setSelection(0, c, rows - 1, c);
        activeR = 0;
        activeC = c;
        updateActive();
      });
      sheet.appendChild(h);
      colHeaderMap.set(c, h);
    }

    for (let r = 0; r < rows; r++) {
      const rh = document.createElement("div");
      rh.className = "row-header";
      rh.textContent = String(r + 1);
      rh.dataset.row = String(r);
      rh.style.gridColumn = "1";
      rh.style.gridRow = String(r + 2);
      rh.addEventListener("mousedown", (e) => {
        e.preventDefault();
        commitEdit();
        setSelection(r, 0, r, cols - 1);
        activeR = r;
        activeC = 0;
        updateActive();
      });
      sheet.appendChild(rh);
      rowHeaderMap.set(r, rh);

      for (let c = 0; c < cols; c++) {
        const cell = document.createElement("div");
        cell.className = "cell";
        cell.dataset.r = String(r);
        cell.dataset.c = String(c);
        cell.style.gridColumn = String(c + 2);
        cell.style.gridRow = String(r + 2);
        paintCellContent(cell, r, c);
        cellMap.set(`${r},${c}`, cell);
        sheet.appendChild(cell);
      }
    }
    paintSelection();
  }

  function paintCellContent(el, r, c) {
    const raw = getRaw(r, c);
    const ev = Formulas.evaluate(raw, getRaw);
    el.textContent = ev.display;
    el.classList.remove("num", "formula-result", "formula-error");
    if (ev.kind === "error") el.classList.add("formula-error");
    else if (ev.kind === "formula") {
      el.classList.add("formula-result");
      if (typeof ev.value === "number") el.classList.add("num");
    } else if (ev.kind === "number" || isNumeric(ev.display)) {
      el.classList.add("num");
    }
    el.title = raw.startsWith("=") ? raw : "";
  }

  function refreshCell(r, c) {
    const el = cellMap.get(`${r},${c}`);
    if (!el || el.classList.contains("editing")) return;
    paintCellContent(el, r, c);
  }

  function refreshAllFormulas() {
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const raw = getRaw(r, c);
        if (raw.startsWith("=") || true) refreshCell(r, c);
      }
    }
  }

  function refreshRange(r1, c1, r2, c2) {
    // refresh range + any formulas that might depend (simple: refresh all formulas in used area)
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) refreshCell(r, c);
    }
    // recompute all formula cells (dependency naive)
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        if (getRaw(r, c).startsWith("=")) refreshCell(r, c);
      }
    }
  }

  // ── Selection ───────────────────────────────────────────
  function setSelection(r1, c1, r2, c2, setActive) {
    selR1 = r1;
    selC1 = c1;
    selR2 = r2;
    selC2 = c2;
    if (setActive !== false) {
      const rg = getRange();
      if (activeR < rg.r1 || activeR > rg.r2 || activeC < rg.c1 || activeC > rg.c2) {
        activeR = rg.r1;
        activeC = rg.c1;
      }
    }
    paintSelection();
    updateFormulaBar();
    updateStatus();
  }

  function paintSelection() {
    const rg = getRange();
    for (const el of cellMap.values()) el.classList.remove("selected", "active");
    for (const h of colHeaderMap.values()) h.classList.remove("active", "sel");
    for (const h of rowHeaderMap.values()) h.classList.remove("active", "sel");

    const multi = rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
    if (multi) {
      for (let r = rg.r1; r <= rg.r2; r++) {
        for (let c = rg.c1; c <= rg.c2; c++) {
          const el = cellMap.get(`${r},${c}`);
          if (el) el.classList.add("selected");
        }
      }
    }

    const act = cellMap.get(`${activeR},${activeC}`);
    if (act) act.classList.add("active");

    for (let c = rg.c1; c <= rg.c2; c++) {
      const h = colHeaderMap.get(c);
      if (h) h.classList.add("sel");
    }
    for (let r = rg.r1; r <= rg.r2; r++) {
      const h = rowHeaderMap.get(r);
      if (h) h.classList.add("sel");
    }
    const ah = colHeaderMap.get(activeC);
    if (ah) ah.classList.add("active");
    const arh = rowHeaderMap.get(activeR);
    if (arh) arh.classList.add("active");

    const tl = cellMap.get(`${rg.r1},${rg.c1}`);
    const br = cellMap.get(`${rg.r2},${rg.c2}`);
    if (tl && br) {
      selOverlay.style.left = tl.offsetLeft + "px";
      selOverlay.style.top = tl.offsetTop + "px";
      selOverlay.style.width = br.offsetLeft + br.offsetWidth - tl.offsetLeft + "px";
      selOverlay.style.height = br.offsetTop + br.offsetHeight - tl.offsetTop + "px";
      selOverlay.classList.toggle("visible", multi);

      fillHandle.style.left = br.offsetLeft + br.offsetWidth - 4 + "px";
      fillHandle.style.top = br.offsetTop + br.offsetHeight - 4 + "px";
      fillHandle.classList.add("visible");
    } else {
      selOverlay.classList.remove("visible");
      fillHandle.classList.remove("visible");
    }
  }

  function updateActive() {
    paintSelection();
    updateFormulaBar();
    scrollActiveIntoView();
  }

  function updateFormulaBar() {
    cellRef.textContent = cellAddr(activeR, activeC);
    if (!editing) formulaInput.value = getRaw(activeR, activeC);
  }

  function updateStatus() {
    const rg = getRange();
    const rr = rg.r2 - rg.r1 + 1;
    const cc = rg.c2 - rg.c1 + 1;
    fileNameEl.textContent = fileName ? (dirty ? "● " : "") + fileName : "Chưa mở file";
    statusReady.textContent = dirty ? "Đã chỉnh sửa" : "Sẵn sàng";

    // Excel-like selection stats
    let sum = 0,
      count = 0,
      countA = 0,
      min = Infinity,
      max = -Infinity;
    for (let r = rg.r1; r <= rg.r2; r++) {
      for (let c = rg.c1; c <= rg.c2; c++) {
        const ev = getDisplay(r, c);
        if (ev.kind === "empty" || ev.display === "") continue;
        countA++;
        const n = typeof ev.value === "number" ? ev.value : Formulas.parseNumber(ev.display);
        if (n !== null) {
          sum += n;
          count++;
          min = Math.min(min, n);
          max = Math.max(max, n);
        }
      }
    }
    const parts = [];
    parts.push(`Lưới <b>${rows}×${cols}</b>`);
    if (rr > 1 || cc > 1) parts.push(`Chọn <b>${rr}×${cc}</b>`);
    if (countA) parts.push(`Đếm <b>${countA}</b>`);
    if (count) {
      parts.push(`Tổng <b>${formatStat(sum)}</b>`);
      parts.push(`TB <b>${formatStat(sum / count)}</b>`);
      if (count > 1) {
        parts.push(`Min <b>${formatStat(min)}</b>`);
        parts.push(`Max <b>${formatStat(max)}</b>`);
      }
    }
    statusStats.innerHTML = parts.map((p) => `<span class="stat">${p}</span>`).join("");
  }

  function formatStat(n) {
    if (!isFinite(n)) return "—";
    if (Number.isInteger(n)) return String(n);
    return (Math.round(n * 1e6) / 1e6).toLocaleString("vi-VN");
  }

  function scrollActiveIntoView() {
    const el = cellMap.get(`${activeR},${activeC}`);
    if (el) el.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  // ── Editing ─────────────────────────────────────────────
  function startEdit(initialChar) {
    if (editing) return;
    editing = true;
    const el = cellMap.get(`${activeR},${activeC}`);
    if (!el) return;
    el.classList.add("editing");
    el.textContent = "";
    const input = document.createElement("input");
    input.className = "cell-edit";
    input.type = "text";
    input.spellcheck = false;
    input.value = initialChar != null ? initialChar : getRaw(activeR, activeC);
    el.appendChild(input);
    editEl = input;
    input.focus();
    if (initialChar == null) input.select();
    else input.setSelectionRange(input.value.length, input.value.length);
    formulaInput.value = input.value;

    input.addEventListener("input", () => {
      formulaInput.value = input.value;
    });
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        e.stopPropagation();
        commitEdit();
        moveActive(e.shiftKey ? -1 : 1, 0);
      } else if (e.key === "Tab") {
        e.preventDefault();
        e.stopPropagation();
        commitEdit();
        moveActive(0, e.shiftKey ? -1 : 1);
      } else if (e.key === "Escape") {
        e.preventDefault();
        cancelEdit();
      } else if (e.key === "ArrowUp" || e.key === "ArrowDown") {
        e.preventDefault();
        commitEdit();
        moveActive(e.key === "ArrowUp" ? -1 : 1, 0);
      } else if (e.key === "ArrowLeft" && input.selectionStart === 0 && input.selectionEnd === 0) {
        e.preventDefault();
        commitEdit();
        moveActive(0, -1);
      } else if (
        e.key === "ArrowRight" &&
        input.selectionStart === input.value.length &&
        input.selectionEnd === input.value.length
      ) {
        e.preventDefault();
        commitEdit();
        moveActive(0, 1);
      }
    });
    input.addEventListener("blur", () => {
      setTimeout(() => {
        if (editing && editEl === input) commitEdit();
      }, 0);
    });
  }

  function commitEdit() {
    if (!editing) return;
    const val = editEl ? editEl.value : formulaInput.value;
    const r = activeR,
      c = activeC;
    editing = false;
    editEl = null;
    const el = cellMap.get(`${r},${c}`);
    if (el) {
      el.classList.remove("editing");
      el.innerHTML = "";
    }
    if (getRaw(r, c) !== val) {
      pushUndo();
      setCell(r, c, val);
      markDirty();
    }
    refreshRange(r, c, r, c);
    updateFormulaBar();
    paintSelection();
    updateStatus();
  }

  function cancelEdit() {
    if (!editing) return;
    editing = false;
    editEl = null;
    const el = cellMap.get(`${activeR},${activeC}`);
    if (el) {
      el.classList.remove("editing");
      el.innerHTML = "";
    }
    refreshCell(activeR, activeC);
    updateFormulaBar();
    paintSelection();
  }

  function moveActive(dr, dc) {
    let nr = Math.max(0, Math.min(rows - 1, activeR + dr));
    let nc = Math.max(0, Math.min(cols - 1, activeC + dc));
    if (activeR + dr >= rows) {
      ensureSize(activeR + dr, activeC);
      nr = activeR + dr;
    }
    if (activeC + dc >= cols) {
      ensureSize(activeR, activeC + dc);
      nc = activeC + dc;
    }
    activeR = nr;
    activeC = nc;
    setSelection(nr, nc, nr, nc);
    scrollActiveIntoView();
  }

  // ── Clipboard ───────────────────────────────────────────
  async function copySelection(isCut) {
    commitEdit();
    const rg = getRange();
    // Copy display values for non-formula? Sheets copies formulas from source.
    // Copy raw values so formulas transfer; also good for paste into Sheets as text.
    const text = CSV.rangeToTSV(getRaw, rg.r1, rg.c1, rg.r2, rg.c2);
    clipText = text;
    clipIsCut = !!isCut;
    clipRange = isCut ? { ...rg } : null;

    try {
      await navigator.clipboard.writeText(text);
    } catch {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.style.position = "fixed";
      ta.style.left = "-9999px";
      document.body.appendChild(ta);
      ta.select();
      try {
        document.execCommand("copy");
      } catch {
        toast("Không copy được — hãy cho phép clipboard");
        document.body.removeChild(ta);
        return;
      }
      document.body.removeChild(ta);
    }
    const cells = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
    toast(isCut ? `Đã cắt ${cells} ô` : `Đã copy ${cells} ô`);
  }

  async function pasteSelection() {
    commitEdit();
    let text = "";
    try {
      text = await navigator.clipboard.readText();
    } catch {
      text = clipText;
    }
    if (!text && clipText) text = clipText;
    if (!text) {
      toast("Clipboard trống hoặc bị chặn");
      return;
    }

    pushUndo();
    const startR = Math.min(selR1, selR2);
    const startC = Math.min(selC1, selC2);

    if (clipIsCut && clipRange) {
      for (let r = clipRange.r1; r <= clipRange.r2; r++) {
        for (let c = clipRange.c1; c <= clipRange.c2; c++) setCell(r, c, "");
      }
      clipIsCut = false;
      clipRange = null;
    }

    const grid = CSV.parseClipboard(text);
    const endR = startR + grid.length - 1;
    const endC = startC + Math.max(...grid.map((r) => r.length)) - 1;
    ensureSize(endR, endC);
    for (let i = 0; i < grid.length; i++) {
      for (let j = 0; j < grid[i].length; j++) {
        setCell(startR + i, startC + j, grid[i][j]);
      }
    }
    markDirty();
    renderSheet();
    setSelection(startR, startC, endR, endC);
    activeR = startR;
    activeC = startC;
    paintSelection();
    toast("Đã dán");
  }

  function deleteSelection() {
    commitEdit();
    pushUndo();
    const rg = getRange();
    for (let r = rg.r1; r <= rg.r2; r++) {
      for (let c = rg.c1; c <= rg.c2; c++) setCell(r, c, "");
    }
    refreshRange(rg.r1, rg.c1, rg.r2, rg.c2);
    markDirty();
    updateFormulaBar();
    updateStatus();
  }

  function selectAll() {
    setSelection(0, 0, rows - 1, cols - 1);
    activeR = 0;
    activeC = 0;
    paintSelection();
  }

  // ── Fill handle ─────────────────────────────────────────
  function startFill(e) {
    e.preventDefault();
    e.stopPropagation();
    commitEdit();
    filling = true;
    fillStart = getRange();
  }

  function applyFill(toR, toC) {
    if (!fillStart) return;
    pushUndo();
    const src = fillStart;
    // Expand selection from original range toward target
    const r1 = Math.min(src.r1, toR);
    const r2 = Math.max(src.r2, toR);
    const c1 = Math.min(src.c1, toC);
    const c2 = Math.max(src.c2, toC);
    ensureSize(r2, c2);

    const srcH = src.r2 - src.r1 + 1;
    const srcW = src.c2 - src.c1 + 1;

    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        // skip source area
        if (r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2) continue;
        const sr2 = src.r1 + (((r - src.r1) % srcH + srcH) % srcH);
        const sc2 = src.c1 + (((c - src.c1) % srcW + srcW) % srcW);
        let raw = getRaw(sr2, sc2);
        // auto-increment numbers when single cell fill
        if (srcH === 1 && srcW === 1) {
          const n = Formulas.parseNumber(raw);
          if (n !== null && !String(raw).startsWith("=")) {
            const dist = (r - src.r1) + (c - src.c1);
            raw = String(n + dist);
          }
        }
        setCell(r, c, raw);
      }
    }
    markDirty();
    renderSheet();
    setSelection(r1, c1, r2, c2);
    toast("Đã fill");
  }

  // ── Sort ────────────────────────────────────────────────
  function sortByActiveCol(asc) {
    commitEdit();
    const col = activeC;
    const { maxR } = usedBounds();
    if (maxR < 1) {
      toast("Không đủ dữ liệu để sắp xếp");
      return;
    }
    pushUndo();
    // Keep row 0 as header if freeze or first row looks like header
    // Always treat first data row as header when sorting (CSV common case)
    const start = 1;
    const body = [];
    for (let r = start; r <= maxR; r++) {
      body.push(data[r].slice());
    }
    body.sort((a, b) => {
      const va = a[col] ?? "";
      const vb = b[col] ?? "";
      const na = Formulas.parseNumber(va);
      const nb = Formulas.parseNumber(vb);
      let cmp;
      if (na !== null && nb !== null) cmp = na - nb;
      else cmp = String(va).localeCompare(String(vb), "vi", { sensitivity: "base" });
      return asc ? cmp : -cmp;
    });
    for (let i = 0; i < body.length; i++) {
      data[start + i] = body[i];
      while (data[start + i].length < cols) data[start + i].push("");
    }
    markDirty();
    renderSheet();
    setSelection(activeR, col, activeR, col);
    toast(asc ? `Đã sắp xếp cột ${colToLetter(col)} A→Z` : `Đã sắp xếp cột ${colToLetter(col)} Z→A`);
  }

  // ── Find / Replace ──────────────────────────────────────
  function openFind(withReplace) {
    findBar.classList.add("open");
    document.getElementById("replaceGroup").style.display = withReplace ? "inline-flex" : "none";
    findInput.focus();
    findInput.select();
  }

  function closeFind() {
    findBar.classList.remove("open");
    findHits = [];
    findIdx = -1;
    findCount.textContent = "";
  }

  function runFind(dir) {
    const q = findInput.value;
    if (!q) {
      findHits = [];
      findCount.textContent = "";
      return;
    }
    const ql = q.toLowerCase();
    findHits = [];
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const v = getRaw(r, c);
        if (v && v.toLowerCase().includes(ql)) findHits.push({ r, c });
      }
    }
    if (!findHits.length) {
      findCount.textContent = "0 kết quả";
      toast("Không tìm thấy");
      return;
    }
    if (findIdx < 0) findIdx = 0;
    else findIdx = (findIdx + dir + findHits.length) % findHits.length;
    const hit = findHits[findIdx];
    activeR = hit.r;
    activeC = hit.c;
    setSelection(hit.r, hit.c, hit.r, hit.c);
    scrollActiveIntoView();
    findCount.textContent = `${findIdx + 1} / ${findHits.length}`;
  }

  function replaceOne() {
    const q = findInput.value;
    const rep = replaceInput.value;
    if (!q) return;
    const cur = getRaw(activeR, activeC);
    if (cur.toLowerCase().includes(q.toLowerCase())) {
      pushUndo();
      // case-insensitive replace first occurrence
      const idx = cur.toLowerCase().indexOf(q.toLowerCase());
      const next = cur.slice(0, idx) + rep + cur.slice(idx + q.length);
      setCell(activeR, activeC, next);
      markDirty();
      refreshRange(activeR, activeC, activeR, activeC);
      updateFormulaBar();
    }
    runFind(1);
  }

  function replaceAll() {
    const q = findInput.value;
    const rep = replaceInput.value;
    if (!q) return;
    pushUndo();
    let n = 0;
    const ql = q.toLowerCase();
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const v = getRaw(r, c);
        if (!v || !v.toLowerCase().includes(ql)) continue;
        // global case-insensitive
        const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
        setCell(r, c, v.replace(re, rep));
        n++;
      }
    }
    markDirty();
    renderSheet();
    toast(`Đã thay thế ${n} ô`);
    runFind(1);
  }

  // ── Rows / Cols ─────────────────────────────────────────
  function addRow() {
    commitEdit();
    pushUndo();
    const at = Math.max(selR1, selR2) + 1;
    data.splice(at, 0, Array(cols).fill(""));
    rows++;
    renderSheet();
    setSelection(at, activeC, at, activeC);
    markDirty();
    toast("Đã thêm 1 hàng");
  }

  function addCol() {
    commitEdit();
    pushUndo();
    const at = Math.max(selC1, selC2) + 1;
    for (let r = 0; r < rows; r++) data[r].splice(at, 0, "");
    colWidths.splice(at, 0, DEFAULT_COL_W);
    cols++;
    renderSheet();
    setSelection(activeR, at, activeR, at);
    markDirty();
    toast("Đã thêm 1 cột");
  }

  function delRow() {
    commitEdit();
    const rg = getRange();
    if (rows <= 1) return;
    const n = rg.r2 - rg.r1 + 1;
    if (!confirm(`Xóa ${n} hàng?`)) return;
    pushUndo();
    data.splice(rg.r1, n);
    rows -= n;
    while (rows < MIN_ROWS) {
      data.push(Array(cols).fill(""));
      rows++;
    }
    activeR = Math.min(rg.r1, rows - 1);
    renderSheet();
    setSelection(activeR, activeC, activeR, activeC);
    markDirty();
  }

  function delCol() {
    commitEdit();
    const rg = getRange();
    if (cols <= 1) return;
    const n = rg.c2 - rg.c1 + 1;
    if (!confirm(`Xóa ${n} cột?`)) return;
    pushUndo();
    for (let r = 0; r < rows; r++) data[r].splice(rg.c1, n);
    colWidths.splice(rg.c1, n);
    cols -= n;
    while (cols < MIN_COLS) {
      for (let r = 0; r < rows; r++) data[r].push("");
      colWidths.push(DEFAULT_COL_W);
      cols++;
    }
    activeC = Math.min(rg.c1, cols - 1);
    renderSheet();
    setSelection(activeR, activeC, activeR, activeC);
    markDirty();
  }

  function autoFitCol(c) {
    let max = 40;
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    ctx.font = '13px "Segoe UI", system-ui, sans-serif';
    max = Math.max(max, ctx.measureText(colToLetter(c)).width + 24);
    for (let r = 0; r < rows; r++) {
      const d = getDisplay(r, c).display;
      if (!d) continue;
      max = Math.max(max, ctx.measureText(d).width + 20);
    }
    colWidths[c] = Math.min(Math.ceil(max), 420);
    applyColTemplate();
    // update cell widths via grid — re-render headers width only
    paintSelection();
  }

  function toggleFreeze() {
    freezeHeader = !freezeHeader;
    sheet.classList.toggle("freeze-header", freezeHeader);
    document.getElementById("btnFreeze").classList.toggle("active-toggle", freezeHeader);
    toast(freezeHeader ? "Đã ghim hàng tiêu đề" : "Bỏ ghim hàng tiêu đề");
    // sort uses freeze as header hint
  }

  function toggleTheme() {
    const cur = document.documentElement.getAttribute("data-theme");
    const next = cur === "dark" ? "light" : "dark";
    if (next === "light") document.documentElement.removeAttribute("data-theme");
    else document.documentElement.setAttribute("data-theme", "dark");
    localStorage.setItem("htmlexxcel-theme", next === "dark" ? "dark" : "light");
    toast(next === "dark" ? "Giao diện tối" : "Giao diện sáng");
  }

  // ── Mouse ───────────────────────────────────────────────
  sheet.addEventListener("mousedown", (e) => {
    if (e.target.classList.contains("col-resize")) {
      e.preventDefault();
      e.stopPropagation();
      resizingCol = +e.target.dataset.col;
      resizeStartX = e.clientX;
      resizeStartW = colWidths[resizingCol];
      e.target.classList.add("dragging");
      return;
    }

    const cell = e.target.closest(".cell");
    if (!cell || cell.classList.contains("editing")) return;
    if (e.target.classList.contains("cell-edit")) return;

    e.preventDefault();
    commitEdit();
    const r = +cell.dataset.r;
    const c = +cell.dataset.c;

    if (e.shiftKey) {
      setSelection(activeR, activeC, r, c, false);
      selR2 = r;
      selC2 = c;
      paintSelection();
      updateStatus();
    } else {
      activeR = r;
      activeC = c;
      selR1 = selR2 = r;
      selC1 = selC2 = c;
      selecting = true;
      paintSelection();
      updateFormulaBar();
      updateStatus();
    }
  });

  sheet.addEventListener("dblclick", (e) => {
    if (e.target.classList.contains("col-resize")) {
      autoFitCol(+e.target.dataset.col);
      return;
    }
    const cell = e.target.closest(".cell");
    if (!cell) return;
    activeR = +cell.dataset.r;
    activeC = +cell.dataset.c;
    setSelection(activeR, activeC, activeR, activeC);
    startEdit();
  });

  sheet.addEventListener("mouseover", (e) => {
    if (selecting) {
      const cell = e.target.closest(".cell");
      if (!cell) return;
      selR2 = +cell.dataset.r;
      selC2 = +cell.dataset.c;
      paintSelection();
      updateStatus();
    }
    if (filling) {
      const cell = e.target.closest(".cell");
      if (!cell) return;
      // preview selection expand
      const toR = +cell.dataset.r;
      const toC = +cell.dataset.c;
      const src = fillStart;
      setSelection(
        Math.min(src.r1, toR),
        Math.min(src.c1, toC),
        Math.max(src.r2, toR),
        Math.max(src.c2, toC),
        false
      );
    }
  });

  window.addEventListener("mousemove", (e) => {
    if (resizingCol == null) return;
    const dx = e.clientX - resizeStartX;
    colWidths[resizingCol] = Math.max(40, resizeStartW + dx);
    applyColTemplate();
    paintSelection();
  });

  window.addEventListener("mouseup", (e) => {
    if (selecting) selecting = false;
    if (filling) {
      filling = false;
      const cell = e.target.closest && e.target.closest(".cell");
      if (cell) applyFill(+cell.dataset.r, +cell.dataset.c);
      else if (fillStart) {
        // keep selection
      }
      fillStart = null;
    }
    if (resizingCol != null) {
      document.querySelectorAll(".col-resize.dragging").forEach((el) => el.classList.remove("dragging"));
      resizingCol = null;
    }
  });

  fillHandle.addEventListener("mousedown", startFill);

  // ── Keyboard ────────────────────────────────────────────
  document.addEventListener("keydown", (e) => {
    const tag = (e.target && e.target.tagName) || "";
    const inInput = tag === "INPUT" || tag === "TEXTAREA" || e.target.isContentEditable;
    const mod = e.ctrlKey || e.metaKey;

    if (mod && e.key.toLowerCase() === "o") {
      e.preventDefault();
      fileInput.click();
      return;
    }
    if (mod && e.key.toLowerCase() === "s") {
      e.preventDefault();
      saveCSV();
      return;
    }
    if (mod && e.key.toLowerCase() === "f") {
      e.preventDefault();
      openFind(false);
      return;
    }
    if (mod && e.key.toLowerCase() === "h") {
      e.preventDefault();
      openFind(true);
      return;
    }
    if (mod && e.key.toLowerCase() === "z") {
      e.preventDefault();
      if (e.shiftKey) redo();
      else undo();
      return;
    }
    if (mod && e.key.toLowerCase() === "y") {
      e.preventDefault();
      redo();
      return;
    }

    if (e.key === "F1") {
      e.preventDefault();
      helpModal.classList.add("open");
      return;
    }

    if (inInput && e.target !== formulaInput) {
      if (e.target === findInput || e.target === replaceInput) {
        if (e.key === "Enter") {
          e.preventDefault();
          if (e.target === replaceInput && e.altKey) replaceAll();
          else if (e.target === replaceInput) replaceOne();
          else runFind(e.shiftKey ? -1 : 1);
        }
        if (e.key === "Escape") closeFind();
      }
      return;
    }

    if (e.target === formulaInput) {
      if (e.key === "Enter") {
        e.preventDefault();
        const val = formulaInput.value;
        if (getRaw(activeR, activeC) !== val) {
          pushUndo();
          setCell(activeR, activeC, val);
          markDirty();
        }
        refreshRange(activeR, activeC, activeR, activeC);
        moveActive(e.shiftKey ? -1 : 1, 0);
      } else if (e.key === "Escape") {
        formulaInput.value = getRaw(activeR, activeC);
        formulaInput.blur();
      } else if (e.key === "Tab") {
        e.preventDefault();
        const val = formulaInput.value;
        if (getRaw(activeR, activeC) !== val) {
          pushUndo();
          setCell(activeR, activeC, val);
          markDirty();
        }
        refreshRange(activeR, activeC, activeR, activeC);
        moveActive(0, e.shiftKey ? -1 : 1);
      }
      return;
    }

    if (editing) return;

    if (mod && e.key.toLowerCase() === "c") {
      e.preventDefault();
      copySelection(false);
      return;
    }
    if (mod && e.key.toLowerCase() === "x") {
      e.preventDefault();
      copySelection(true);
      return;
    }
    if (mod && e.key.toLowerCase() === "v") {
      e.preventDefault();
      pasteSelection();
      return;
    }
    if (mod && e.key.toLowerCase() === "a") {
      e.preventDefault();
      selectAll();
      return;
    }

    if (e.key === "ArrowUp") {
      e.preventDefault();
      if (e.shiftKey) {
        selR2 = Math.max(0, selR2 - 1);
        paintSelection();
        updateStatus();
      } else moveActive(-1, 0);
      return;
    }
    if (e.key === "ArrowDown") {
      e.preventDefault();
      if (e.shiftKey) {
        selR2 = Math.min(rows - 1, selR2 + 1);
        paintSelection();
        updateStatus();
      } else moveActive(1, 0);
      return;
    }
    if (e.key === "ArrowLeft") {
      e.preventDefault();
      if (e.shiftKey) {
        selC2 = Math.max(0, selC2 - 1);
        paintSelection();
        updateStatus();
      } else moveActive(0, -1);
      return;
    }
    if (e.key === "ArrowRight") {
      e.preventDefault();
      if (e.shiftKey) {
        selC2 = Math.min(cols - 1, selC2 + 1);
        paintSelection();
        updateStatus();
      } else moveActive(0, 1);
      return;
    }
    if (e.key === "Tab") {
      e.preventDefault();
      moveActive(0, e.shiftKey ? -1 : 1);
      return;
    }
    if (e.key === "Enter") {
      e.preventDefault();
      if (e.shiftKey) moveActive(-1, 0);
      else startEdit();
      return;
    }
    if (e.key === "F2") {
      e.preventDefault();
      startEdit();
      return;
    }
    if (e.key === "Delete" || e.key === "Backspace") {
      e.preventDefault();
      if (e.key === "Backspace") {
        deleteSelection();
        startEdit("");
      } else deleteSelection();
      return;
    }
    if (e.key === "Escape") {
      setSelection(activeR, activeC, activeR, activeC);
      closeMenus();
      return;
    }
    if (e.key === "Home") {
      e.preventDefault();
      if (mod) {
        activeR = 0;
        activeC = 0;
      } else activeC = 0;
      setSelection(activeR, activeC, activeR, activeC);
      scrollActiveIntoView();
      return;
    }
    if (e.key === "End") {
      e.preventDefault();
      let lc = cols - 1;
      while (lc > 0 && getRaw(activeR, lc) === "") lc--;
      activeC = lc;
      setSelection(activeR, activeC, activeR, activeC);
      scrollActiveIntoView();
      return;
    }
    if (e.key === "PageDown") {
      e.preventDefault();
      moveActive(20, 0);
      return;
    }
    if (e.key === "PageUp") {
      e.preventDefault();
      moveActive(-20, 0);
      return;
    }

    if (!mod && e.key.length === 1 && !e.altKey) {
      e.preventDefault();
      startEdit(e.key);
    }
  });

  // ── Drag & drop ─────────────────────────────────────────
  ["dragenter", "dragover"].forEach((ev) => {
    document.body.addEventListener(ev, (e) => {
      e.preventDefault();
      document.body.classList.add("drag-over");
    });
  });
  document.body.addEventListener("dragleave", () => document.body.classList.remove("drag-over"));
  document.body.addEventListener("drop", (e) => {
    e.preventDefault();
    document.body.classList.remove("drag-over");
    const f = e.dataTransfer.files[0];
    if (f) openFile(f);
  });

  // ── Menus ───────────────────────────────────────────────
  function closeMenus() {
    document.querySelectorAll(".menu-drop.open").forEach((el) => el.classList.remove("open"));
  }

  document.querySelectorAll("[data-menu]").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const id = btn.getAttribute("data-menu");
      const menu = document.getElementById(id);
      const was = menu.classList.contains("open");
      closeMenus();
      if (!was) menu.classList.add("open");
    });
  });
  document.addEventListener("click", () => closeMenus());

  // ── Buttons ─────────────────────────────────────────────
  document.getElementById("btnOpen").onclick = () => fileInput.click();
  document.getElementById("btnOpenHint").onclick = () => fileInput.click();
  document.getElementById("btnSave").onclick = () => saveCSV();
  document.getElementById("btnNew").onclick = () => newSheet();
  document.getElementById("btnCopy").onclick = () => copySelection(false);
  document.getElementById("btnCut").onclick = () => copySelection(true);
  document.getElementById("btnPaste").onclick = () => pasteSelection();
  document.getElementById("btnFind").onclick = () => openFind(false);
  document.getElementById("btnReplace").onclick = () => openFind(true);
  document.getElementById("btnAddRow").onclick = () => addRow();
  document.getElementById("btnAddCol").onclick = () => addCol();
  document.getElementById("btnDelRow").onclick = () => delRow();
  document.getElementById("btnDelCol").onclick = () => delCol();
  document.getElementById("btnUndo").onclick = () => undo();
  document.getElementById("btnRedo").onclick = () => redo();
  document.getElementById("btnSortAsc").onclick = () => sortByActiveCol(true);
  document.getElementById("btnSortDesc").onclick = () => sortByActiveCol(false);
  document.getElementById("btnFreeze").onclick = () => toggleFreeze();
  document.getElementById("btnTheme").onclick = () => toggleTheme();
  document.getElementById("btnHelp").onclick = () => helpModal.classList.add("open");
  document.getElementById("helpClose").onclick = () => helpModal.classList.remove("open");
  helpModal.addEventListener("click", (e) => {
    if (e.target === helpModal) helpModal.classList.remove("open");
  });

  document.getElementById("findClose").onclick = () => closeFind();
  document.getElementById("findNext").onclick = () => runFind(1);
  document.getElementById("findPrev").onclick = () => runFind(-1);
  document.getElementById("btnReplaceOne").onclick = () => replaceOne();
  document.getElementById("btnReplaceAll").onclick = () => replaceAll();
  findInput.addEventListener("input", () => {
    findIdx = -1;
    runFind(1);
  });

  fileInput.addEventListener("change", () => {
    const f = fileInput.files[0];
    if (f) openFile(f);
    fileInput.value = "";
  });

  window.addEventListener("resize", () => paintSelection());
  window.addEventListener("beforeunload", (e) => {
    if (dirty) {
      e.preventDefault();
      e.returnValue = "";
    }
  });

  // Sample data button
  document.getElementById("btnSample").onclick = () => {
    if (dirty && !confirm("Ghi đè dữ liệu hiện tại bằng mẫu?")) return;
    loadGrid(
      [
        ["Sản phẩm", "Số lượng", "Đơn giá", "Thành tiền"],
        ["Laptop", "3", "15000000", "=B2*C2"],
        ["Chuột", "10", "150000", "=B3*C3"],
        ["Bàn phím", "5", "450000", "=B4*C4"],
        ["Màn hình", "2", "3200000", "=B5*C5"],
        ["", "", "Tổng:", "=SUM(D2:D5)"],
        ["", "", "Trung bình:", "=AVERAGE(D2:D5)"],
        ["", "", "Max:", "=MAX(D2:D5)"],
      ],
      "mau-cong-thuc.csv"
    );
  };

  // ── Init ────────────────────────────────────────────────
  (function init() {
    const theme = localStorage.getItem("htmlexxcel-theme");
    if (theme === "dark") document.documentElement.setAttribute("data-theme", "dark");

    data = Array.from({ length: DEFAULT_ROWS }, () => Array(DEFAULT_COLS).fill(""));
    rows = DEFAULT_ROWS;
    cols = DEFAULT_COLS;
    colWidths = Array.from({ length: cols }, () => DEFAULT_COL_W);
    renderSheet();
    setSelection(0, 0, 0, 0);
    updateStatus();
    updateUndoButtons();
  })();
})();
