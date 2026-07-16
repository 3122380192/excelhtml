/**
 * HTML Excel — Excel-like application
 */
(function () {
  "use strict";

  const DEFAULT_ROWS = 80;
  const DEFAULT_COLS = 20;
  const MIN_ROWS = 30;
  const MIN_COLS = 14;
  const DEFAULT_COL_W = 100;
  const ROW_H = 22;
  const CORNER_W = 46;
  const MAX_UNDO = 60;

  const emptyStyle = () => ({
    bold: false,
    italic: false,
    underline: false,
    align: "",
    color: "",
    fill: "",
    font: "",
    size: 0,
    border: "",
    wrap: false,
    numFmt: "general",
    decimals: 2,
    note: "",
  });

  // ── State ───────────────────────────────────────────────
  let data = [];
  let styles = [];
  let rows = DEFAULT_ROWS;
  let cols = DEFAULT_COLS;
  let colWidths = [];
  let fileName = "";
  let fileFormat = "xlsx";
  let dirty = false;
  let freezeHeader = false;
  let showGrid = true;
  let showFormulas = false;
  let zoom = 1;

  let workbookSheets = [];
  let activeSheetIndex = 0;

  let activeR = 0, activeC = 0;
  let selR1 = 0, selC1 = 0, selR2 = 0, selC2 = 0;
  let selecting = false;
  let editing = false;
  let editEl = null;

  let clipText = "";
  let clipIsCut = false;
  let clipRange = null;
  let clipStyles = null;
  let formatPainter = null;

  let findHits = [];
  let findIdx = -1;

  const undoStack = [];
  const redoStack = [];

  let filling = false;
  let fillStart = null;
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
  const sheetTabs = document.getElementById("sheetTabs");
  const ctxMenu = document.getElementById("ctxMenu");
  const backstage = document.getElementById("backstage");
  const backstageMain = document.getElementById("backstageMain");
  const zoomLabel = document.getElementById("zoomLabel");

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
  function cellAddr(r, c) { return colToLetter(c) + (r + 1); }
  function getRange() {
    return {
      r1: Math.min(selR1, selR2),
      c1: Math.min(selC1, selC2),
      r2: Math.max(selR1, selR2),
      c2: Math.max(selC1, selC2),
    };
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
  function clone2D(arr) {
    return arr.map((row) => row.map((c) => (typeof c === "object" && c ? { ...c } : c)));
  }
  function cloneStyles(st) {
    return st.map((row) => row.map((s) => (s ? { ...s } : emptyStyle())));
  }

  function pushUndo() {
    undoStack.push({
      data: clone2D(data),
      styles: cloneStyles(styles),
      rows, cols,
      colWidths: colWidths.slice(),
      activeR, activeC, selR1, selC1, selR2, selC2,
    });
    if (undoStack.length > MAX_UNDO) undoStack.shift();
    redoStack.length = 0;
  }

  function restoreSnap(snap) {
    data = clone2D(snap.data);
    styles = cloneStyles(snap.styles);
    rows = snap.rows;
    cols = snap.cols;
    colWidths = snap.colWidths.slice();
    activeR = snap.activeR; activeC = snap.activeC;
    selR1 = snap.selR1; selC1 = snap.selC1; selR2 = snap.selR2; selC2 = snap.selC2;
    renderSheet();
    paintSelection();
    updateFormulaBar();
    syncRibbonFromActive();
    markDirty();
  }

  function undo() {
    if (!undoStack.length) return toast("Không còn hoàn tác");
    commitEdit();
    redoStack.push({
      data: clone2D(data), styles: cloneStyles(styles),
      rows, cols, colWidths: colWidths.slice(),
      activeR, activeC, selR1, selC1, selR2, selC2,
    });
    restoreSnap(undoStack.pop());
    toast("Hoàn tác");
  }
  function redo() {
    if (!redoStack.length) return toast("Không còn làm lại");
    commitEdit();
    undoStack.push({
      data: clone2D(data), styles: cloneStyles(styles),
      rows, cols, colWidths: colWidths.slice(),
      activeR, activeC, selR1, selC1, selR2, selC2,
    });
    restoreSnap(redoStack.pop());
    toast("Làm lại");
  }

  function ensureSize(needR, needC) {
    let changed = false;
    if (needR >= rows) {
      const nr = needR + 20;
      for (let r = rows; r < nr; r++) {
        data[r] = new Array(cols).fill("");
        styles[r] = Array.from({ length: cols }, emptyStyle);
      }
      rows = nr;
      changed = true;
    }
    if (needC >= cols) {
      const nc = needC + 5;
      for (let r = 0; r < rows; r++) {
        while (data[r].length < nc) data[r].push("");
        while (styles[r].length < nc) styles[r].push(emptyStyle());
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
  function getStyle(r, c) {
    if (r < 0 || c < 0 || r >= rows || c >= cols) return emptyStyle();
    return styles[r][c] || emptyStyle();
  }
  function setCell(r, c, val) {
    ensureSize(r, c);
    data[r][c] = val == null ? "" : String(val);
  }
  function setStyle(r, c, patch) {
    ensureSize(r, c);
    styles[r][c] = { ...getStyle(r, c), ...patch };
  }

  // ── Format display value ────────────────────────────────
  function formatDisplay(raw, st, r, c) {
    if (showFormulas && String(raw).startsWith("=")) return raw;
    const ev = Formulas.evaluate(raw, getRaw);
    let display = ev.display;
    let kind = ev.kind;
    const n = typeof ev.value === "number" ? ev.value : Formulas.parseNumber(ev.display);
    const fmt = (st && st.numFmt) || "general";
    const dec = st && st.decimals != null ? st.decimals : 2;

    if (kind !== "error" && n !== null && fmt !== "text" && fmt !== "general") {
      if (fmt === "number") {
        display = n.toLocaleString("vi-VN", { minimumFractionDigits: dec, maximumFractionDigits: dec });
      } else if (fmt === "currency") {
        display = n.toLocaleString("vi-VN", { minimumFractionDigits: dec, maximumFractionDigits: dec }) + " ₫";
      } else if (fmt === "percent") {
        display = (n * (String(raw).includes("%") ? 1 : 100)).toLocaleString("vi-VN", {
          minimumFractionDigits: dec,
          maximumFractionDigits: dec,
        }) + "%";
        // if value already 0.5 meaning 50%, Excel stores 0.5 — we treat number as ratio if < 2 else already percent-ish
        if (!String(raw).includes("%") && Math.abs(n) <= 1) {
          display = (n * 100).toLocaleString("vi-VN", { minimumFractionDigits: dec, maximumFractionDigits: dec }) + "%";
        } else if (!String(raw).includes("%")) {
          display = n.toLocaleString("vi-VN", { minimumFractionDigits: dec, maximumFractionDigits: dec }) + "%";
        }
      } else if (fmt === "date") {
        // excel serial or timestamp-ish
        if (n > 20000 && n < 60000) {
          const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
          display = d.toLocaleDateString("vi-VN");
        }
      }
    }
    return { display, kind, value: ev.value, isNum: n !== null && kind !== "error" && kind !== "text" };
  }

  // ── Workbook ────────────────────────────────────────────
  function gridFromData() {
    const g = [];
    for (let r = 0; r < rows; r++) {
      g[r] = [];
      for (let c = 0; c < cols; c++) g[r][c] = data[r][c] ?? "";
    }
    return g;
  }
  function stylesFromData() {
    return cloneStyles(styles);
  }

  function padGrid(grid, stGrid) {
    if (!grid || !grid.length) grid = [[""]];
    const maxC = Math.max(...grid.map((r) => (r ? r.length : 0)), MIN_COLS);
    const nr = Math.max(grid.length + 10, MIN_ROWS, DEFAULT_ROWS);
    const nc = Math.max(maxC + 3, MIN_COLS, DEFAULT_COLS);
    const out = [], st = [];
    for (let r = 0; r < nr; r++) {
      out[r] = [];
      st[r] = [];
      for (let c = 0; c < nc; c++) {
        out[r][c] = grid[r] && grid[r][c] != null ? String(grid[r][c]) : "";
        st[r][c] = (stGrid && stGrid[r] && stGrid[r][c]) ? { ...emptyStyle(), ...stGrid[r][c] } : emptyStyle();
      }
    }
    return { data: out, styles: st, rows: nr, cols: nc, colWidths: Array.from({ length: nc }, () => DEFAULT_COL_W) };
  }

  function snapshotActiveToWorkbook() {
    if (!workbookSheets.length) {
      workbookSheets = [{
        name: "Sheet1",
        data: gridFromData(),
        styles: stylesFromData(),
        rows, cols,
        colWidths: colWidths.slice(),
      }];
      activeSheetIndex = 0;
      return;
    }
    const sh = workbookSheets[activeSheetIndex];
    if (!sh) return;
    sh.data = gridFromData();
    sh.styles = stylesFromData();
    sh.rows = rows;
    sh.cols = cols;
    sh.colWidths = colWidths.slice();
  }

  function applySheetState(sh) {
    data = sh.data.map((row) => row.slice());
    styles = cloneStyles(sh.styles || padGrid(sh.data).styles);
    rows = sh.rows;
    cols = sh.cols;
    colWidths = (sh.colWidths && sh.colWidths.slice()) || Array.from({ length: cols }, () => DEFAULT_COL_W);
    while (styles.length < rows) styles.push(Array.from({ length: cols }, emptyStyle));
    for (let r = 0; r < rows; r++) {
      while (styles[r].length < cols) styles[r].push(emptyStyle());
    }
    activeR = activeC = selR1 = selC1 = selR2 = selC2 = 0;
    editing = false;
    editEl = null;
    renderSheet();
    setSelection(0, 0, 0, 0);
    updateStatus();
    syncRibbonFromActive();
  }

  function renderSheetTabs() {
    if (!sheetTabs) return;
    sheetTabs.innerHTML = "";
    workbookSheets.forEach((sh, i) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "sheet-tab" + (i === activeSheetIndex ? " active" : "");
      btn.textContent = sh.name || "Sheet" + (i + 1);
      btn.addEventListener("click", () => switchSheet(i));
      btn.addEventListener("dblclick", (e) => { e.stopPropagation(); renameSheet(i); });
      btn.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        if (confirm("Xóa sheet \"" + sh.name + "\"?")) deleteSheet(i);
      });
      sheetTabs.appendChild(btn);
    });
    const add = document.createElement("button");
    add.type = "button";
    add.className = "sheet-tab-add";
    add.textContent = "+";
    add.title = "Thêm sheet";
    add.addEventListener("click", addWorkbookSheet);
    sheetTabs.appendChild(add);
  }

  function switchSheet(index) {
    if (index === activeSheetIndex || index < 0 || index >= workbookSheets.length) return;
    commitEdit();
    snapshotActiveToWorkbook();
    activeSheetIndex = index;
    undoStack.length = 0;
    redoStack.length = 0;
    applySheetState(workbookSheets[index]);
    renderSheetTabs();
  }

  function addWorkbookSheet() {
    commitEdit();
    snapshotActiveToWorkbook();
    const padded = padGrid([[""]]);
    workbookSheets.push({
      name: "Sheet" + (workbookSheets.length + 1),
      data: padded.data,
      styles: padded.styles,
      rows: padded.rows,
      cols: padded.cols,
      colWidths: padded.colWidths,
    });
    activeSheetIndex = workbookSheets.length - 1;
    if (["csv", "tsv", "txt"].includes(fileFormat)) {
      fileFormat = "xlsx";
      fileName = (IO.baseName(fileName || "Workbook") || "Workbook") + ".xlsx";
    }
    undoStack.length = 0;
    applySheetState(workbookSheets[activeSheetIndex]);
    renderSheetTabs();
    markDirty();
    toast("Đã thêm sheet");
  }

  function deleteSheet(index) {
    if (workbookSheets.length <= 1) return toast("Phải còn ít nhất 1 sheet");
    commitEdit();
    snapshotActiveToWorkbook();
    workbookSheets.splice(index, 1);
    if (activeSheetIndex >= workbookSheets.length) activeSheetIndex = workbookSheets.length - 1;
    else if (index < activeSheetIndex) activeSheetIndex--;
    applySheetState(workbookSheets[activeSheetIndex]);
    renderSheetTabs();
    markDirty();
  }

  function renameSheet(index) {
    const sh = workbookSheets[index];
    if (!sh) return;
    const name = prompt("Tên sheet:", sh.name);
    if (name == null) return;
    const cleaned = name.trim().slice(0, 31).replace(/[\\/?*[\]]/g, "_");
    if (!cleaned) return;
    sh.name = cleaned;
    renderSheetTabs();
    markDirty();
  }

  function exportSheetsPayload() {
    snapshotActiveToWorkbook();
    return workbookSheets.map((sh) => ({ name: sh.name, grid: sh.data }));
  }

  // ── Load / Save ─────────────────────────────────────────
  function loadWorkbook(wb, quiet) {
    const sheets = (wb.sheets && wb.sheets.length ? wb.sheets : [{ name: "Sheet1", grid: [[""]] }]).map((s, i) => {
      const padded = padGrid(s.grid);
      return {
        name: s.name || "Sheet" + (i + 1),
        data: padded.data,
        styles: padded.styles,
        rows: padded.rows,
        cols: padded.cols,
        colWidths: padded.colWidths,
      };
    });
    workbookSheets = sheets;
    activeSheetIndex = 0;
    fileName = wb.fileName || "Workbook.xlsx";
    fileFormat = (wb.format || IO.extOf(fileName) || "xlsx").toLowerCase();
    if (["xlsm", "xlsb", "fods", "html", "htm"].includes(fileFormat)) fileFormat = "xlsx";
    if (fileFormat === "tab") fileFormat = "tsv";
    dirty = false;
    undoStack.length = 0;
    redoStack.length = 0;
    hint.classList.remove("show");
    closeBackstage();
    applySheetState(workbookSheets[0]);
    renderSheetTabs();
    if (!quiet) {
      const g = wb.sheets[0].grid || [[""]];
      const maxC = Math.max(...g.map((r) => (r ? r.length : 0)), 1);
      const multi = sheets.length > 1 ? ` · ${sheets.length} sheet` : "";
      toast(`Đã mở ${fileName} (${g.length}×${maxC}${multi})`);
    }
  }

  function newWorkbook(force) {
    if (!force && dirty && !confirm("Dữ liệu chưa lưu sẽ mất. Tạo mới?")) return;
    const padded = padGrid([[""]]);
    workbookSheets = [{
      name: "Sheet1",
      data: padded.data,
      styles: padded.styles,
      rows: padded.rows,
      cols: padded.cols,
      colWidths: padded.colWidths,
    }];
    activeSheetIndex = 0;
    fileName = "Workbook.xlsx";
    fileFormat = "xlsx";
    dirty = false;
    undoStack.length = 0;
    redoStack.length = 0;
    hint.classList.remove("show");
    closeBackstage();
    applySheetState(workbookSheets[0]);
    renderSheetTabs();
    toast("Sổ làm việc mới");
  }

  function openFile(file) {
    toast("Đang mở " + file.name + "…");
    IO.readFile(file)
      .then((wb) => loadWorkbook(wb))
      .catch((e) => {
        toast("Lỗi: " + (e.message || e));
        console.error(e);
      });
  }

  function resolveSaveFormat(fmt) {
    if (!fmt || fmt === "auto") {
      const ext = IO.extOf(fileName);
      if (ext && ["csv", "tsv", "txt", "xlsx", "xls", "ods"].includes(ext)) return ext;
      return fileFormat || "xlsx";
    }
    return fmt;
  }

  function saveAs(fmt) {
    commitEdit();
    const format = resolveSaveFormat(fmt);
    if (["csv", "tsv", "txt"].includes(format) && workbookSheets.length > 1) {
      if (!confirm("CSV/TSV chỉ lưu sheet đang mở. Tiếp tục?")) return;
    }
    try {
      const saved = IO.saveWorkbook(exportSheetsPayload(), fileName || "Workbook", format);
      fileName = saved;
      fileFormat = format === "txt" ? "csv" : format;
      dirty = false;
      updateStatus();
      toast("Đã lưu: " + saved);
      closeBackstage();
    } catch (e) {
      toast("Lỗi lưu: " + (e.message || e));
    }
  }

  // ── Render ──────────────────────────────────────────────
  function applyColTemplate() {
    sheet.style.gridTemplateColumns = `${CORNER_W}px ${colWidths.map((w) => w + "px").join(" ")}`;
    sheet.style.gridTemplateRows = `${ROW_H}px repeat(${rows}, ${ROW_H}px)`;
    sheet.style.transform = zoom !== 1 ? `scale(${zoom})` : "";
  }

  function renderSheet() {
    cellMap.clear();
    colHeaderMap.clear();
    rowHeaderMap.clear();
    sheet.innerHTML = "";
    sheet.classList.toggle("freeze-header", freezeHeader);
    sheet.classList.toggle("no-grid", !showGrid);
    while (colWidths.length < cols) colWidths.push(DEFAULT_COL_W);
    applyColTemplate();

    const corner = document.createElement("div");
    corner.className = "corner";
    corner.style.gridColumn = "1";
    corner.style.gridRow = "1";
    corner.title = "Chọn tất cả";
    corner.addEventListener("mousedown", (e) => {
      e.preventDefault();
      selectAll();
    });
    sheet.appendChild(corner);

    for (let c = 0; c < cols; c++) {
      const h = document.createElement("div");
      h.className = "col-header";
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
        activeR = 0; activeC = c;
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
        activeR = r; activeC = 0;
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
        cellMap.set(r + "," + c, cell);
        sheet.appendChild(cell);
      }
    }
    paintSelection();
  }

  function paintCellContent(el, r, c) {
    const raw = getRaw(r, c);
    const st = getStyle(r, c);
    const fd = formatDisplay(raw, st, r, c);
    el.textContent = fd.display;
    el.className = "cell";
    if (fd.kind === "error") el.classList.add("formula-error");
    else if (fd.kind === "formula") {
      el.classList.add("formula-result");
      if (fd.isNum) el.classList.add("num");
    } else if (fd.isNum || (fd.kind === "number")) el.classList.add("num");

    if (st.bold) el.classList.add("bold");
    if (st.italic) el.classList.add("italic");
    if (st.underline) el.classList.add("underline");
    if (st.wrap) el.classList.add("wrap");
    if (st.align === "left") el.classList.add("align-left");
    if (st.align === "center") el.classList.add("align-center");
    if (st.align === "right") el.classList.add("align-right");
    if (st.border === "all") el.classList.add("b-all");
    if (st.border === "bottom") el.classList.add("b-b");

    if (st.fill) el.style.background = st.fill;
    else el.style.background = "";
    if (st.color) el.style.color = st.color;
    else el.style.color = "";
    if (st.font) el.style.fontFamily = st.font;
    else el.style.fontFamily = "";
    if (st.size) el.style.fontSize = st.size + "px";
    else el.style.fontSize = "";

    el.title = st.note ? st.note : (raw.startsWith("=") ? raw : "");
  }

  function refreshCell(r, c) {
    const el = cellMap.get(r + "," + c);
    if (!el || el.classList.contains("editing")) return;
    paintCellContent(el, r, c);
  }

  function refreshRange(r1, c1, r2, c2) {
    for (let r = r1; r <= r2; r++)
      for (let c = c1; c <= c2; c++) refreshCell(r, c);
    for (let r = 0; r < rows; r++)
      for (let c = 0; c < cols; c++)
        if (getRaw(r, c).startsWith("=")) refreshCell(r, c);
  }

  function refreshAll() {
    for (let r = 0; r < rows; r++)
      for (let c = 0; c < cols; c++) refreshCell(r, c);
  }

  // ── Selection ───────────────────────────────────────────
  function setSelection(r1, c1, r2, c2, setActive) {
    selR1 = r1; selC1 = c1; selR2 = r2; selC2 = c2;
    if (setActive !== false) {
      const rg = getRange();
      if (activeR < rg.r1 || activeR > rg.r2 || activeC < rg.c1 || activeC > rg.c2) {
        activeR = rg.r1; activeC = rg.c1;
      }
    }
    paintSelection();
    updateFormulaBar();
    updateStatus();
    syncRibbonFromActive();
  }

  function paintSelection() {
    const rg = getRange();
    for (const el of cellMap.values()) el.classList.remove("selected", "active");
    for (const h of colHeaderMap.values()) h.classList.remove("active", "sel");
    for (const h of rowHeaderMap.values()) h.classList.remove("active", "sel");

    const multi = rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
    if (multi) {
      for (let r = rg.r1; r <= rg.r2; r++)
        for (let c = rg.c1; c <= rg.c2; c++) {
          const el = cellMap.get(r + "," + c);
          if (el) el.classList.add("selected");
        }
    }
    const act = cellMap.get(activeR + "," + activeC);
    if (act) act.classList.add("active");

    for (let c = rg.c1; c <= rg.c2; c++) {
      const h = colHeaderMap.get(c);
      if (h) h.classList.add("sel");
    }
    for (let r = rg.r1; r <= rg.r2; r++) {
      const h = rowHeaderMap.get(r);
      if (h) h.classList.add("sel");
    }
    colHeaderMap.get(activeC)?.classList.add("active");
    rowHeaderMap.get(activeR)?.classList.add("active");

    const tl = cellMap.get(rg.r1 + "," + rg.c1);
    const br = cellMap.get(rg.r2 + "," + rg.c2);
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
    syncRibbonFromActive();
  }

  function updateFormulaBar() {
    cellRef.textContent = cellAddr(activeR, activeC);
    if (!editing) formulaInput.value = getRaw(activeR, activeC);
  }

  function updateStatus() {
    const rg = getRange();
    const rr = rg.r2 - rg.r1 + 1, cc = rg.c2 - rg.c1 + 1;
    fileNameEl.textContent = (dirty ? "● " : "") + (fileName || "Sổ làm việc chưa đặt tên");
    statusReady.textContent = dirty ? "Đã chỉnh sửa" : "Sẵn sàng";
    document.title = (dirty ? "* " : "") + (fileName || "HTML Excel") + " — HTML Excel";

    let sum = 0, count = 0, countA = 0, min = Infinity, max = -Infinity;
    for (let r = rg.r1; r <= rg.r2; r++) {
      for (let c = rg.c1; c <= rg.c2; c++) {
        const fd = formatDisplay(getRaw(r, c), getStyle(r, c), r, c);
        if (!fd.display) continue;
        countA++;
        const n = typeof fd.value === "number" ? fd.value : Formulas.parseNumber(fd.display);
        if (n !== null) {
          sum += n; count++;
          min = Math.min(min, n); max = Math.max(max, n);
        }
      }
    }
    const parts = [];
    if (rr > 1 || cc > 1) parts.push(`Trung bình: <b>${count ? fmtStat(sum / count) : "—"}</b>`);
    if (count) {
      parts.push(`Đếm: <b>${count}</b>`);
      parts.push(`Tổng: <b>${fmtStat(sum)}</b>`);
      if (count > 1) {
        parts.push(`Min: <b>${fmtStat(min)}</b>`);
        parts.push(`Max: <b>${fmtStat(max)}</b>`);
      }
    } else if (countA) parts.push(`Đếm: <b>${countA}</b>`);
    statusStats.innerHTML = parts.map((p) => `<span class="stat">${p}</span>`).join("");
    zoomLabel.textContent = Math.round(zoom * 100) + "%";
  }

  function fmtStat(n) {
    if (!isFinite(n)) return "—";
    if (Number.isInteger(n)) return n.toLocaleString("vi-VN");
    return (Math.round(n * 1e6) / 1e6).toLocaleString("vi-VN");
  }

  function scrollActiveIntoView() {
    cellMap.get(activeR + "," + activeC)?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  function selectAll() {
    setSelection(0, 0, rows - 1, cols - 1);
    activeR = 0; activeC = 0;
    paintSelection();
  }

  // ── Style apply ─────────────────────────────────────────
  function applyStyleToSelection(patch) {
    commitEdit();
    pushUndo();
    const rg = getRange();
    for (let r = rg.r1; r <= rg.r2; r++)
      for (let c = rg.c1; c <= rg.c2; c++) setStyle(r, c, patch);
    refreshRange(rg.r1, rg.c1, rg.r2, rg.c2);
    markDirty();
    syncRibbonFromActive();
  }

  function toggleStyle(key) {
    const st = getStyle(activeR, activeC);
    applyStyleToSelection({ [key]: !st[key] });
  }

  function syncRibbonFromActive() {
    const st = getStyle(activeR, activeC);
    const on = (id, v) => document.getElementById(id)?.classList.toggle("on", !!v);
    on("btnBold", st.bold);
    on("btnItalic", st.italic);
    on("btnUnderline", st.underline);
    on("btnAlignLeft", st.align === "left");
    on("btnAlignCenter", st.align === "center");
    on("btnAlignRight", st.align === "right");
    on("btnWrap", st.wrap);
    on("btnFreeze", freezeHeader);
    on("btnGridlines", showGrid);
    on("btnShowFormulas", showFormulas);
    on("btnFormatPainter", !!formatPainter);
    const ff = document.getElementById("fontFamily");
    const fs = document.getElementById("fontSize");
    const nf = document.getElementById("numFmt");
    if (ff) ff.value = st.font || "Segoe UI";
    if (fs) fs.value = String(st.size || 12);
    if (nf) nf.value = st.numFmt || "general";
    const tc = document.getElementById("textColor");
    const fc = document.getElementById("fillColor");
    if (tc && st.color) tc.value = st.color;
    if (fc && st.fill) fc.value = st.fill;
    const sc = document.getElementById("swatchColor");
    const sf = document.getElementById("swatchFill");
    if (sc) sc.style.background = st.color || "#000";
    if (sf) sf.style.background = st.fill || "#ffff00";
  }

  // ── Editing ─────────────────────────────────────────────
  function startEdit(initialChar) {
    if (editing) return;
    editing = true;
    const el = cellMap.get(activeR + "," + activeC);
    if (!el) return;
    el.classList.add("editing");
    el.textContent = "";
    const input = document.createElement("input");
    input.className = "cell-edit";
    input.spellcheck = false;
    input.value = initialChar != null ? initialChar : getRaw(activeR, activeC);
    el.appendChild(input);
    editEl = input;
    input.focus();
    if (initialChar == null) input.select();
    else input.setSelectionRange(input.value.length, input.value.length);
    formulaInput.value = input.value;
    input.addEventListener("input", () => { formulaInput.value = input.value; });
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault(); e.stopPropagation();
        commitEdit(); moveActive(e.shiftKey ? -1 : 1, 0);
      } else if (e.key === "Tab") {
        e.preventDefault(); e.stopPropagation();
        commitEdit(); moveActive(0, e.shiftKey ? -1 : 1);
      } else if (e.key === "Escape") {
        e.preventDefault(); cancelEdit();
      } else if (e.key === "ArrowUp" || e.key === "ArrowDown") {
        e.preventDefault();
        commitEdit();
        moveActive(e.key === "ArrowUp" ? -1 : 1, 0);
      }
    });
    input.addEventListener("blur", () => {
      setTimeout(() => { if (editing && editEl === input) commitEdit(); }, 0);
    });
  }

  function commitEdit() {
    if (!editing) return;
    const val = editEl ? editEl.value : formulaInput.value;
    const r = activeR, c = activeC;
    editing = false;
    editEl = null;
    const el = cellMap.get(r + "," + c);
    if (el) { el.classList.remove("editing"); el.innerHTML = ""; }
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
    const el = cellMap.get(activeR + "," + activeC);
    if (el) { el.classList.remove("editing"); el.innerHTML = ""; }
    refreshCell(activeR, activeC);
    updateFormulaBar();
    paintSelection();
  }

  function moveActive(dr, dc) {
    let nr = Math.max(0, Math.min(rows - 1, activeR + dr));
    let nc = Math.max(0, Math.min(cols - 1, activeC + dc));
    if (activeR + dr >= rows) { ensureSize(activeR + dr, activeC); nr = activeR + dr; }
    if (activeC + dc >= cols) { ensureSize(activeR, activeC + dc); nc = activeC + dc; }
    activeR = nr; activeC = nc;
    setSelection(nr, nc, nr, nc);
    scrollActiveIntoView();
  }

  // ── Clipboard ───────────────────────────────────────────
  async function copySelection(isCut) {
    commitEdit();
    const rg = getRange();
    const text = CSV.rangeToTSV(getRaw, rg.r1, rg.c1, rg.r2, rg.c2);
    clipText = text;
    clipIsCut = !!isCut;
    clipRange = isCut ? { ...rg } : null;
    clipStyles = [];
    for (let r = rg.r1; r <= rg.r2; r++) {
      const row = [];
      for (let c = rg.c1; c <= rg.c2; c++) row.push({ ...getStyle(r, c) });
      clipStyles.push(row);
    }
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      const ta = document.createElement("textarea");
      ta.value = text;
      ta.style.cssText = "position:fixed;left:-9999px";
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      document.body.removeChild(ta);
    }
    toast(isCut ? "Đã cắt" : "Đã sao chép");
  }

  async function pasteSelection() {
    commitEdit();
    let text = "";
    try { text = await navigator.clipboard.readText(); } catch { text = clipText; }
    if (!text && clipText) text = clipText;
    if (!text) return toast("Clipboard trống");

    pushUndo();
    const startR = Math.min(selR1, selR2);
    const startC = Math.min(selC1, selC2);

    if (clipIsCut && clipRange) {
      for (let r = clipRange.r1; r <= clipRange.r2; r++)
        for (let c = clipRange.c1; c <= clipRange.c2; c++) {
          setCell(r, c, "");
          setStyle(r, c, emptyStyle());
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
        if (clipStyles && clipStyles[i] && clipStyles[i][j]) {
          setStyle(startR + i, startC + j, clipStyles[i][j]);
        }
      }
    }
    markDirty();
    renderSheet();
    setSelection(startR, startC, endR, endC);
    activeR = startR; activeC = startC;
    toast("Đã dán");
  }

  function deleteSelection(clearStyles) {
    commitEdit();
    pushUndo();
    const rg = getRange();
    for (let r = rg.r1; r <= rg.r2; r++)
      for (let c = rg.c1; c <= rg.c2; c++) {
        setCell(r, c, "");
        if (clearStyles) setStyle(r, c, emptyStyle());
      }
    refreshRange(rg.r1, rg.c1, rg.r2, rg.c2);
    markDirty();
    updateFormulaBar();
  }

  // ── AutoSum ─────────────────────────────────────────────
  function autoSum() {
    commitEdit();
    // Place SUM below/right of selection of numbers
    const rg = getRange();
    let targetR = rg.r2 + 1;
    let targetC = rg.c1;
    // If single empty cell selected, look up for numbers
    if (rg.r1 === rg.r2 && rg.c1 === rg.c2 && !getRaw(rg.r1, rg.c1)) {
      let r = rg.r1 - 1;
      while (r >= 0 && Formulas.parseNumber(formatDisplay(getRaw(r, rg.c1), getStyle(r, rg.c1)).display) !== null) r--;
      r++;
      if (r < rg.r1) {
        const formula = `=SUM(${cellAddr(r, rg.c1)}:${cellAddr(rg.r1 - 1, rg.c1)})`;
        pushUndo();
        setCell(rg.r1, rg.c1, formula);
        refreshRange(rg.r1, rg.c1, rg.r1, rg.c1);
        markDirty();
        updateFormulaBar();
        toast(formula);
        return;
      }
    }
    const formula = `=SUM(${cellAddr(rg.r1, rg.c1)}:${cellAddr(rg.r2, rg.c2)})`;
    ensureSize(targetR, targetC);
    // if target has data, put in active cell
    if (getRaw(targetR, targetC)) {
      targetR = activeR;
      targetC = activeC;
    }
    pushUndo();
    setCell(targetR, targetC, formula);
    activeR = targetR; activeC = targetC;
    setSelection(targetR, targetC, targetR, targetC);
    refreshRange(targetR, targetC, targetR, targetC);
    markDirty();
    toast(formula);
  }

  function insertFx(name) {
    commitEdit();
    const rg = getRange();
    let formula;
    const range = `${cellAddr(rg.r1, rg.c1)}:${cellAddr(rg.r2, rg.c2)}`;
    const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
    if (name === "IF") formula = '=IF(A1>0,"OK","")';
    else if (name === "ROUND") formula = single ? `=ROUND(${cellAddr(activeR, activeC)},2)` : `=ROUND(${cellAddr(rg.r1, rg.c1)},2)`;
    else if (name === "POWER") formula = `=POWER(${cellAddr(activeR, activeC)},2)`;
    else if (name === "MOD") formula = `=MOD(${cellAddr(activeR, activeC)},2)`;
    else if (name === "ABS" || name === "SQRT") formula = `=${name}(${cellAddr(activeR, activeC)})`;
    else if (name === "PMT") formula = "=PMT(0.01,12,-10000000)";
    else if (name === "FV") formula = "=FV(0.01,12,-1000000)";
    else if (name === "PV") formula = "=PV(0.01,12,-1000000)";
    else if (name === "NPV") formula = "=NPV(0.1,100,200,300)";
    else if (single) formula = `=${name}()`;
    else formula = `=${name}(${range})`;
    pushUndo();
    setCell(activeR, activeC, formula);
    refreshRange(activeR, activeC, activeR, activeC);
    markDirty();
    updateFormulaBar();
    startEdit();
  }

  function selectionNumbers() {
    const rg = getRange();
    const vals = [];
    const cells = [];
    for (let r = rg.r1; r <= rg.r2; r++) {
      for (let c = rg.c1; c <= rg.c2; c++) {
        const fd = formatDisplay(getRaw(r, c), getStyle(r, c), r, c);
        const n = typeof fd.value === "number" ? fd.value : Formulas.parseNumber(getRaw(r, c));
        if (n !== null) {
          vals.push(n);
          cells.push({ r, c, n, raw: getRaw(r, c) });
        }
      }
    }
    return { rg, vals, cells };
  }

  function applyCalcTool(tool) {
    commitEdit();
    const { rg, vals, cells } = selectionNumbers();

    if (tool === "sum" || tool === "avg" || tool === "product" || tool === "count" || tool === "minmax") {
      if (!vals.length && tool !== "count") return toast("Chọn vùng có số");
      if (tool === "sum") {
        const f = `=SUM(${cellAddr(rg.r1, rg.c1)}:${cellAddr(rg.r2, rg.c2)})`;
        pushUndo();
        const tr = rg.r2 + 1;
        ensureSize(tr, rg.c1);
        setCell(tr, rg.c1, f);
        setStyle(tr, rg.c1, { bold: true, fill: "#e7f3ec" });
        renderSheet();
        setSelection(tr, rg.c1, tr, rg.c1);
        markDirty();
        return toast("Tổng: " + f);
      }
      if (tool === "avg") {
        const f = `=AVERAGE(${cellAddr(rg.r1, rg.c1)}:${cellAddr(rg.r2, rg.c2)})`;
        pushUndo();
        const tr = rg.r2 + 1;
        ensureSize(tr, rg.c1);
        setCell(tr, rg.c1, f);
        setStyle(tr, rg.c1, { bold: true, fill: "#e7f3ec" });
        renderSheet();
        markDirty();
        return toast("TB: " + f);
      }
      if (tool === "product") {
        const f = `=PRODUCT(${cellAddr(rg.r1, rg.c1)}:${cellAddr(rg.r2, rg.c2)})`;
        pushUndo();
        const tr = rg.r2 + 1;
        ensureSize(tr, rg.c1);
        setCell(tr, rg.c1, f);
        setStyle(tr, rg.c1, { bold: true, fill: "#e7f3ec" });
        renderSheet();
        markDirty();
        return toast("Tích: " + f);
      }
      if (tool === "count") {
        const f = `=COUNT(${cellAddr(rg.r1, rg.c1)}:${cellAddr(rg.r2, rg.c2)})`;
        pushUndo();
        const tr = rg.r2 + 1;
        ensureSize(tr, rg.c1);
        setCell(tr, rg.c1, f);
        setStyle(tr, rg.c1, { bold: true, fill: "#e7f3ec" });
        renderSheet();
        markDirty();
        return toast("Đếm: " + f);
      }
      if (tool === "minmax") {
        pushUndo();
        const tr = rg.r2 + 1;
        ensureSize(tr + 1, rg.c1);
        setCell(tr, rg.c1, `=MIN(${cellAddr(rg.r1, rg.c1)}:${cellAddr(rg.r2, rg.c2)})`);
        setCell(tr + 1, rg.c1, `=MAX(${cellAddr(rg.r1, rg.c1)}:${cellAddr(rg.r2, rg.c2)})`);
        setStyle(tr, rg.c1, { bold: true, fill: "#fff3cd" });
        setStyle(tr + 1, rg.c1, { bold: true, fill: "#fff3cd" });
        renderSheet();
        markDirty();
        return toast("Đã chèn MIN / MAX");
      }
    }

    if (tool === "pct") {
      if (vals.length < 1) return toast("Chọn dãy số");
      const total = vals.reduce((a, b) => a + b, 0);
      if (!total) return toast("Tổng = 0");
      pushUndo();
      // write % in column to the right
      const outC = rg.c2 + 1;
      ensureSize(rg.r2, outC);
      setCell(Math.max(0, rg.r1 - 1), outC, "% tổng");
      setStyle(Math.max(0, rg.r1 - 1), outC, { bold: true });
      for (const cell of cells) {
        setCell(cell.r, outC, String(cell.n / total));
        setStyle(cell.r, outC, { numFmt: "percent", decimals: 2 });
      }
      renderSheet();
      markDirty();
      return toast("Đã tính % của tổng (cột bên phải)");
    }

    if (tool === "pctChange") {
      if (cells.length < 2) return toast("Cần ≥ 2 số theo cột/hàng");
      pushUndo();
      const outC = rg.c2 + 1;
      ensureSize(rg.r2, outC);
      setCell(cells[0].r, outC, "—");
      for (let i = 1; i < cells.length; i++) {
        const prev = cells[i - 1].n;
        const cur = cells[i].n;
        const ch = prev === 0 ? 0 : (cur - prev) / prev;
        setCell(cells[i].r, outC, String(ch));
        setStyle(cells[i].r, outC, { numFmt: "percent", decimals: 2 });
      }
      renderSheet();
      markDirty();
      return toast("Đã tính % thay đổi");
    }

    if (tool === "running") {
      if (!cells.length) return toast("Chọn dãy số");
      pushUndo();
      const outC = rg.c2 + 1;
      ensureSize(rg.r2, outC);
      let run = 0;
      for (const cell of cells) {
        run += cell.n;
        setCell(cell.r, outC, String(run));
        setStyle(cell.r, outC, { numFmt: "number", fill: "#e7f3ec" });
      }
      renderSheet();
      markDirty();
      return toast("Đã cộng dồn → cột phải");
    }

    if (tool === "diff") {
      if (cells.length < 2) return toast("Cần ≥ 2 số");
      pushUndo();
      const outC = rg.c2 + 1;
      ensureSize(rg.r2, outC);
      setCell(cells[0].r, outC, "—");
      for (let i = 1; i < cells.length; i++) {
        setCell(cells[i].r, outC, String(cells[i].n - cells[i - 1].n));
        setStyle(cells[i].r, outC, { numFmt: "number" });
      }
      renderSheet();
      markDirty();
      return toast("Đã tính hiệu số");
    }

    if (tool === "vatAdd" || tool === "vatRem" || tool === "vatCustom") {
      let rate = 0.1;
      if (tool === "vatCustom") {
        const p = prompt("Thuế VAT (%)", "10");
        if (p == null) return;
        rate = Number(p) / 100;
        if (!isFinite(rate)) return toast("VAT không hợp lệ");
      }
      if (!cells.length) return toast("Chọn ô số");
      pushUndo();
      for (const cell of cells) {
        const v = tool === "vatRem" ? cell.n / (1 + rate) : cell.n * (1 + rate);
        setCell(cell.r, cell.c, String(Math.round(v * 100) / 100));
        setStyle(cell.r, cell.c, { numFmt: "currency" });
      }
      renderSheet();
      markDirty();
      return toast(tool === "vatRem" ? "Đã bỏ VAT" : "Đã cộng VAT " + rate * 100 + "%");
    }

    if (tool === "margin") {
      // cost in col1, price in col2 of selection (2 cols) OR prompt
      if (rg.c2 - rg.c1 < 1 && cells.length < 2) {
        const cost = Number(prompt("Giá vốn:", getRaw(activeR, activeC) || "0"));
        const price = Number(prompt("Giá bán:", "0"));
        if (!isFinite(cost) || !isFinite(price) || price === 0) return toast("Số không hợp lệ");
        const margin = (price - cost) / price;
        pushUndo();
        setCell(activeR, activeC, String(margin));
        setStyle(activeR, activeC, { numFmt: "percent", decimals: 2, bold: true });
        refreshCell(activeR, activeC);
        markDirty();
        return toast("Biên lợi nhuận: " + (margin * 100).toFixed(2) + "%");
      }
      // two columns: cost | price → margin in next col
      pushUndo();
      const outC = rg.c2 + 1;
      ensureSize(rg.r2, outC);
      for (let r = rg.r1; r <= rg.r2; r++) {
        const cost = Formulas.parseNumber(getRaw(r, rg.c1));
        const price = Formulas.parseNumber(getRaw(r, rg.c1 + 1));
        if (cost === null || price === null || !price) continue;
        setCell(r, outC, String((price - cost) / price));
        setStyle(r, outC, { numFmt: "percent", decimals: 2 });
      }
      renderSheet();
      markDirty();
      return toast("Biên LN → cột phải (vốn | giá bán)");
    }

    if (tool === "interest") {
      const p = Number(prompt("Số tiền gốc (P):", "10000000"));
      const r = Number(prompt("Lãi suất năm %:", "8"));
      const t = Number(prompt("Số năm (t):", "1"));
      if (![p, r, t].every(isFinite)) return toast("Số không hợp lệ");
      const interest = p * (r / 100) * t;
      const total = p + interest;
      pushUndo();
      ensureSize(activeR + 3, activeC + 1);
      setCell(activeR, activeC, "Gốc"); setCell(activeR, activeC + 1, String(p));
      setCell(activeR + 1, activeC, "Lãi đơn"); setCell(activeR + 1, activeC + 1, String(interest));
      setCell(activeR + 2, activeC, "Tổng"); setCell(activeR + 2, activeC + 1, String(total));
      setStyle(activeR + 1, activeC + 1, { numFmt: "currency", fill: "#e7f3ec" });
      setStyle(activeR + 2, activeC + 1, { numFmt: "currency", bold: true });
      renderSheet();
      markDirty();
      return toast("Lãi đơn: " + interest.toLocaleString("vi-VN"));
    }

    if (tool === "compound") {
      const p = Number(prompt("Số tiền gốc (P):", "10000000"));
      const r = Number(prompt("Lãi suất năm %:", "8"));
      const t = Number(prompt("Số năm (t):", "5"));
      const n = Number(prompt("Kỳ ghép/năm (n):", "12"));
      if (![p, r, t, n].every(isFinite) || !n) return toast("Số không hợp lệ");
      const amount = p * Math.pow(1 + r / 100 / n, n * t);
      pushUndo();
      ensureSize(activeR + 2, activeC + 1);
      setCell(activeR, activeC, "Gốc"); setCell(activeR, activeC + 1, String(p));
      setCell(activeR + 1, activeC, "Lãi kép (FV)"); setCell(activeR + 1, activeC + 1, String(Math.round(amount)));
      setStyle(activeR + 1, activeC + 1, { numFmt: "currency", bold: true, fill: "#e7f3ec" });
      renderSheet();
      markDirty();
      return toast("FV: " + Math.round(amount).toLocaleString("vi-VN"));
    }

    if (tool === "loan") {
      const pv = Number(prompt("Khoản vay (VNĐ):", "500000000"));
      const annual = Number(prompt("Lãi suất năm %:", "10"));
      const years = Number(prompt("Số năm:", "5"));
      if (![pv, annual, years].every(isFinite) || !years) return toast("Số không hợp lệ");
      const rate = annual / 100 / 12;
      const nper = years * 12;
      const pmt = rate === 0 ? pv / nper : (pv * rate * Math.pow(1 + rate, nper)) / (Math.pow(1 + rate, nper) - 1);
      pushUndo();
      ensureSize(activeR + 4, activeC + 1);
      const rowsW = [
        ["Khoản vay", pv],
        ["Lãi suất/năm", annual + "%"],
        ["Thời hạn", years + " năm"],
        ["Trả hàng tháng", Math.round(pmt)],
        ["Tổng trả", Math.round(pmt * nper)],
      ];
      rowsW.forEach((row, i) => {
        setCell(activeR + i, activeC, row[0]);
        setCell(activeR + i, activeC + 1, String(row[1]));
      });
      setStyle(activeR + 3, activeC + 1, { numFmt: "currency", bold: true, fill: "#e7f3ec" });
      setStyle(activeR + 4, activeC + 1, { numFmt: "currency", bold: true });
      renderSheet();
      markDirty();
      return toast("Trả/tháng ≈ " + Math.round(pmt).toLocaleString("vi-VN") + " ₫");
    }

    if (tool === "discount") {
      if (!cells.length) {
        const price = Number(prompt("Giá gốc:", "100000"));
        const pct = Number(prompt("Giảm %:", "20"));
        if (![price, pct].every(isFinite)) return toast("Số không hợp lệ");
        const sale = price * (1 - pct / 100);
        pushUndo();
        setCell(activeR, activeC, String(sale));
        setStyle(activeR, activeC, { numFmt: "currency", bold: true });
        refreshCell(activeR, activeC);
        markDirty();
        return toast("Giá sau giảm: " + sale.toLocaleString("vi-VN"));
      }
      const pct = Number(prompt("Giảm giá % cho vùng chọn:", "10"));
      if (!isFinite(pct)) return toast("Không hợp lệ");
      pushUndo();
      for (const cell of cells) {
        setCell(cell.r, cell.c, String(Math.round(cell.n * (1 - pct / 100) * 100) / 100));
        setStyle(cell.r, cell.c, { numFmt: "currency" });
      }
      renderSheet();
      markDirty();
      return toast("Đã giảm " + pct + "%");
    }

    if (tool === "evaluate") {
      const expr = prompt("Nhập biểu thức (vd: 15*20+100 hoặc =A1*2):", getRaw(activeR, activeC) || "");
      if (expr == null || expr === "") return;
      pushUndo();
      const formula = expr.startsWith("=") ? expr : "=" + expr;
      setCell(activeR, activeC, formula);
      refreshRange(activeR, activeC, activeR, activeC);
      markDirty();
      updateFormulaBar();
      return toast("Đã tính: " + formula);
    }
  }

  // Calculator
  let calc = null;
  function ensureCalc() {
    if (calc) return calc;
    calc = CalcUI.createCalculator({
      onInsert: (val) => {
        commitEdit();
        pushUndo();
        setCell(activeR, activeC, val);
        refreshRange(activeR, activeC, activeR, activeC);
        markDirty();
        updateFormulaBar();
        hint.classList.remove("show");
      },
      onToast: toast,
    });
    return calc;
  }
  function openCalculator() {
    ensureCalc().open();
    toast("Máy tính — phím số hoạt động khi panel mở");
  }

  // ── Fill ────────────────────────────────────────────────
  function applyFill(toR, toC) {
    if (!fillStart) return;
    pushUndo();
    const src = fillStart;
    const r1 = Math.min(src.r1, toR), r2 = Math.max(src.r2, toR);
    const c1 = Math.min(src.c1, toC), c2 = Math.max(src.c2, toC);
    ensureSize(r2, c2);
    const srcH = src.r2 - src.r1 + 1, srcW = src.c2 - src.c1 + 1;
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        if (r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2) continue;
        const sr = src.r1 + (((r - src.r1) % srcH + srcH) % srcH);
        const sc = src.c1 + (((c - src.c1) % srcW + srcW) % srcW);
        let raw = getRaw(sr, sc);
        if (srcH === 1 && srcW === 1) {
          const n = Formulas.parseNumber(raw);
          if (n !== null && !String(raw).startsWith("=")) {
            raw = String(n + (r - src.r1) + (c - src.c1));
          }
        }
        setCell(r, c, raw);
        setStyle(r, c, getStyle(sr, sc));
      }
    }
    markDirty();
    renderSheet();
    setSelection(r1, c1, r2, c2);
    toast("Đã điền");
  }

  // ── Sort / Find ─────────────────────────────────────────
  function sortByActiveCol(asc) {
    commitEdit();
    const col = activeC;
    let maxR = 0;
    for (let r = 0; r < rows; r++)
      for (let c = 0; c < cols; c++) if (getRaw(r, c)) maxR = Math.max(maxR, r);
    if (maxR < 1) return toast("Không đủ dữ liệu");
    pushUndo();
    const start = 1;
    const body = [];
    for (let r = start; r <= maxR; r++) {
      body.push({ data: data[r].slice(), styles: styles[r].map((s) => ({ ...s })) });
    }
    body.sort((a, b) => {
      const va = a.data[col] ?? "", vb = b.data[col] ?? "";
      const na = Formulas.parseNumber(va), nb = Formulas.parseNumber(vb);
      let cmp = na !== null && nb !== null ? na - nb : String(va).localeCompare(String(vb), "vi", { sensitivity: "base" });
      return asc ? cmp : -cmp;
    });
    for (let i = 0; i < body.length; i++) {
      data[start + i] = body[i].data;
      styles[start + i] = body[i].styles;
      while (data[start + i].length < cols) data[start + i].push("");
      while (styles[start + i].length < cols) styles[start + i].push(emptyStyle());
    }
    markDirty();
    renderSheet();
    setSelection(activeR, col, activeR, col);
    toast(asc ? "Đã sắp xếp A→Z" : "Đã sắp xếp Z→A");
  }

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
    if (!q) return;
    const ql = q.toLowerCase();
    findHits = [];
    for (let r = 0; r < rows; r++)
      for (let c = 0; c < cols; c++) {
        const v = getRaw(r, c);
        if (v && v.toLowerCase().includes(ql)) findHits.push({ r, c });
      }
    if (!findHits.length) { findCount.textContent = "0"; return toast("Không tìm thấy"); }
    if (findIdx < 0) findIdx = 0;
    else findIdx = (findIdx + dir + findHits.length) % findHits.length;
    const hit = findHits[findIdx];
    activeR = hit.r; activeC = hit.c;
    setSelection(hit.r, hit.c, hit.r, hit.c);
    scrollActiveIntoView();
    findCount.textContent = `${findIdx + 1} / ${findHits.length}`;
  }
  function replaceOne() {
    const q = findInput.value, rep = replaceInput.value;
    if (!q) return;
    const cur = getRaw(activeR, activeC);
    if (cur.toLowerCase().includes(q.toLowerCase())) {
      pushUndo();
      const idx = cur.toLowerCase().indexOf(q.toLowerCase());
      setCell(activeR, activeC, cur.slice(0, idx) + rep + cur.slice(idx + q.length));
      markDirty();
      refreshRange(activeR, activeC, activeR, activeC);
    }
    runFind(1);
  }
  function replaceAll() {
    const q = findInput.value, rep = replaceInput.value;
    if (!q) return;
    pushUndo();
    let n = 0;
    const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
    for (let r = 0; r < rows; r++)
      for (let c = 0; c < cols; c++) {
        const v = getRaw(r, c);
        if (v && v.toLowerCase().includes(q.toLowerCase())) {
          setCell(r, c, v.replace(re, rep));
          n++;
        }
      }
    markDirty();
    renderSheet();
    toast(`Đã thay ${n} ô`);
  }

  // ── Rows / Cols ─────────────────────────────────────────
  function addRow() {
    commitEdit(); pushUndo();
    const at = Math.max(selR1, selR2) + 1;
    data.splice(at, 0, Array(cols).fill(""));
    styles.splice(at, 0, Array.from({ length: cols }, emptyStyle));
    rows++;
    renderSheet();
    setSelection(at, activeC, at, activeC);
    markDirty();
  }
  function addCol() {
    commitEdit(); pushUndo();
    const at = Math.max(selC1, selC2) + 1;
    for (let r = 0; r < rows; r++) {
      data[r].splice(at, 0, "");
      styles[r].splice(at, 0, emptyStyle());
    }
    colWidths.splice(at, 0, DEFAULT_COL_W);
    cols++;
    renderSheet();
    setSelection(activeR, at, activeR, at);
    markDirty();
  }
  function delRow() {
    commitEdit();
    const rg = getRange();
    if (rows <= 1) return;
    if (!confirm(`Xóa ${rg.r2 - rg.r1 + 1} hàng?`)) return;
    pushUndo();
    const n = rg.r2 - rg.r1 + 1;
    data.splice(rg.r1, n);
    styles.splice(rg.r1, n);
    rows -= n;
    while (rows < MIN_ROWS) {
      data.push(Array(cols).fill(""));
      styles.push(Array.from({ length: cols }, emptyStyle));
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
    if (!confirm(`Xóa ${rg.c2 - rg.c1 + 1} cột?`)) return;
    pushUndo();
    const n = rg.c2 - rg.c1 + 1;
    for (let r = 0; r < rows; r++) {
      data[r].splice(rg.c1, n);
      styles[r].splice(rg.c1, n);
    }
    colWidths.splice(rg.c1, n);
    cols -= n;
    while (cols < MIN_COLS) {
      for (let r = 0; r < rows; r++) {
        data[r].push("");
        styles[r].push(emptyStyle());
      }
      colWidths.push(DEFAULT_COL_W);
      cols++;
    }
    activeC = Math.min(rg.c1, cols - 1);
    renderSheet();
    setSelection(activeR, activeC, activeR, activeC);
    markDirty();
  }

  function autoFitCol(c) {
    const canvas = document.createElement("canvas");
    const ctx = canvas.getContext("2d");
    ctx.font = '13px "Segoe UI"';
    let max = ctx.measureText(colToLetter(c)).width + 24;
    for (let r = 0; r < rows; r++) {
      const d = formatDisplay(getRaw(r, c), getStyle(r, c)).display;
      if (d) max = Math.max(max, ctx.measureText(d).width + 16);
    }
    colWidths[c] = Math.min(Math.ceil(max), 420);
    applyColTemplate();
    paintSelection();
  }

  function setZoom(z) {
    zoom = Math.max(0.5, Math.min(2, z));
    applyColTemplate();
    paintSelection();
    updateStatus();
  }

  // ── Ribbon / Backstage ──────────────────────────────────
  function setRibbon(name) {
    if (name === "file") {
      openBackstage("new");
      return;
    }
    if (name === "calculator") {
      openCalculator();
      // keep previous ribbon body; highlight calc tools tab visually via calc ribbon
      document.querySelectorAll(".ribbon-tab").forEach((t) => {
        t.classList.toggle("active", t.dataset.ribbon === "calc");
      });
      document.querySelectorAll(".ribbon-body").forEach((b) => {
        b.classList.toggle("active", b.id === "ribbon-calc");
      });
      return;
    }
    document.querySelectorAll(".ribbon-tab").forEach((t) => {
      t.classList.toggle("active", t.dataset.ribbon === name);
    });
    document.querySelectorAll(".ribbon-body").forEach((b) => {
      b.classList.toggle("active", b.id === "ribbon-" + name);
    });
  }

  function openBackstage(panel) {
    backstage.classList.add("open");
    renderBackstage(panel || "new");
    document.querySelectorAll(".backstage-nav button[data-bs]").forEach((b) => {
      b.classList.toggle("active", b.dataset.bs === (panel || "new"));
    });
  }
  function closeBackstage() {
    backstage.classList.remove("open");
    document.querySelectorAll(".ribbon-tab").forEach((t) => {
      t.classList.toggle("active", t.dataset.ribbon === "home");
    });
    document.querySelectorAll(".ribbon-body").forEach((b) => {
      b.classList.toggle("active", b.id === "ribbon-home");
    });
  }

  function renderBackstage(panel) {
    const main = backstageMain;
    if (panel === "new") {
      main.innerHTML = `<h2>Mới</h2>
        <div class="backstage-grid">
          <div class="bs-card" data-act="blank"><div class="ico">📄</div><div class="lbl">Sổ trống</div><div class="sub">Workbook.xlsx</div></div>
          <div class="bs-card" data-act="sample"><div class="ico">✨</div><div class="lbl">Mẫu bán hàng</div><div class="sub">Công thức + 2 sheet</div></div>
        </div>`;
    } else if (panel === "open") {
      main.innerHTML = `<h2>Mở</h2>
        <div class="backstage-grid">
          <div class="bs-card" data-act="browse"><div class="ico">📂</div><div class="lbl">Duyệt máy tính</div><div class="sub">XLSX, XLS, ODS, CSV…</div></div>
        </div>
        <p style="margin-top:20px;color:var(--muted)">Hoặc kéo thả file vào cửa sổ làm việc.</p>`;
    } else if (panel === "save") {
      main.innerHTML = `<h2>Lưu</h2>
        <div class="backstage-grid">
          <div class="bs-card" data-act="save-auto"><div class="ico">💾</div><div class="lbl">Lưu</div><div class="sub">Định dạng hiện tại</div></div>
          <div class="bs-card" data-act="save-xlsx"><div class="ico">📗</div><div class="lbl">Excel (.xlsx)</div></div>
          <div class="bs-card" data-act="save-xls"><div class="ico">📘</div><div class="lbl">Excel 97 (.xls)</div></div>
          <div class="bs-card" data-act="save-ods"><div class="ico">📙</div><div class="lbl">ODS</div></div>
          <div class="bs-card" data-act="save-csv"><div class="ico">📄</div><div class="lbl">CSV</div></div>
        </div>`;
    } else if (panel === "export") {
      main.innerHTML = `<h2>Xuất</h2>
        <div class="backstage-grid">
          <div class="bs-card" data-act="save-xlsx"><div class="ico">📗</div><div class="lbl">Xuất Excel</div></div>
          <div class="bs-card" data-act="save-csv"><div class="ico">📄</div><div class="lbl">Xuất CSV</div></div>
          <div class="bs-card" data-act="save-tsv"><div class="ico">📄</div><div class="lbl">Xuất TSV</div></div>
        </div>`;
    } else {
      main.innerHTML = `<h2>Thông tin</h2>
        <p><b>Tệp:</b> ${fileName || "(chưa đặt tên)"}</p>
        <p><b>Định dạng:</b> ${(fileFormat || "").toUpperCase()}</p>
        <p><b>Số sheet:</b> ${workbookSheets.length}</p>
        <p><b>Lưới hiện tại:</b> ${rows} × ${cols}</p>
        <p style="margin-top:16px;color:var(--muted)">HTML Excel — bảng tính trong trình duyệt, giao diện giống Microsoft Excel.</p>`;
    }
    main.querySelectorAll(".bs-card").forEach((card) => {
      card.addEventListener("click", () => {
        const a = card.dataset.act;
        if (a === "blank") newWorkbook(true);
        else if (a === "sample") loadSample();
        else if (a === "browse") { closeBackstage(); fileInput.click(); }
        else if (a === "save-auto") saveAs("auto");
        else if (a === "save-xlsx") saveAs("xlsx");
        else if (a === "save-xls") saveAs("xls");
        else if (a === "save-ods") saveAs("ods");
        else if (a === "save-csv") saveAs("csv");
        else if (a === "save-tsv") saveAs("tsv");
      });
    });
  }

  function loadSample() {
    loadWorkbook({
      fileName: "Mau_BanHang.xlsx",
      format: "xlsx",
      sheets: [
        {
          name: "Bán hàng",
          grid: [
            ["Sản phẩm", "Số lượng", "Đơn giá", "Thành tiền", "Ghi chú"],
            ["Laptop Dell", "3", "18500000", "=B2*C2", ""],
            ["Chuột Logitech", "12", "250000", "=B3*C3", ""],
            ["Bàn phím cơ", "8", "890000", "=B4*C4", "Hot"],
            ["Màn hình 27\"", "4", "4500000", "=B5*C5", ""],
            ["USB 64GB", "20", "120000", "=B6*C6", ""],
            ["", "", "TỔNG:", "=SUM(D2:D6)", ""],
            ["", "", "Trung bình:", "=AVERAGE(D2:D6)", ""],
            ["", "", "Max:", "=MAX(D2:D6)", ""],
            ["", "", "Min:", "=MIN(D2:D6)", ""],
          ],
        },
        {
          name: "Hướng dẫn",
          grid: [
            ["Tính năng", "Cách dùng"],
            ["Ribbon Trang chủ", "Định dạng chữ, màu, căn lề, số"],
            ["Công thức", "Gõ =SUM(A1:A5) hoặc nút ∑ AutoSum"],
            ["Nhiều sheet", "Tab dưới lưới · nút + thêm sheet"],
            ["Mở / Lưu", "Ribbon Tệp → Mở / Lưu (XLSX, CSV…)"],
            ["Copy/Paste", "Ctrl+C / Ctrl+V giống Google Sheets"],
            ["Chuột phải", "Menu: cắt, chèn hàng, AutoSum…"],
          ],
        },
      ],
    });
    // style header row
    for (let c = 0; c < 5; c++) {
      setStyle(0, c, { bold: true, fill: "#217346", color: "#ffffff", align: "center" });
    }
    setStyle(6, 2, { bold: true });
    setStyle(6, 3, { bold: true, numFmt: "currency" });
    for (let r = 1; r <= 5; r++) {
      setStyle(r, 2, { numFmt: "currency" });
      setStyle(r, 3, { numFmt: "currency" });
    }
    refreshAll();
    snapshotActiveToWorkbook();
  }

  // ── Context menu ────────────────────────────────────────
  function showCtx(x, y) {
    ctxMenu.classList.add("open");
    const w = ctxMenu.offsetWidth || 200;
    const h = ctxMenu.offsetHeight || 280;
    ctxMenu.style.left = Math.min(x, window.innerWidth - w - 8) + "px";
    ctxMenu.style.top = Math.min(y, window.innerHeight - h - 8) + "px";
  }
  function hideCtx() { ctxMenu.classList.remove("open"); }

  // ── Mouse ───────────────────────────────────────────────
  sheet.addEventListener("mousedown", (e) => {
    hideCtx();
    if (e.target.classList.contains("col-resize")) {
      e.preventDefault();
      resizingCol = +e.target.dataset.col;
      resizeStartX = e.clientX;
      resizeStartW = colWidths[resizingCol];
      e.target.classList.add("dragging");
      return;
    }
    const cell = e.target.closest(".cell");
    if (!cell || cell.classList.contains("editing")) return;
    if (e.button === 2) return;

    e.preventDefault();
    commitEdit();
    const r = +cell.dataset.r, c = +cell.dataset.c;

    if (formatPainter) {
      pushUndo();
      setStyle(r, c, { ...formatPainter });
      refreshCell(r, c);
      markDirty();
      if (!e.ctrlKey) {
        formatPainter = null;
        syncRibbonFromActive();
      }
      return;
    }

    if (e.shiftKey) {
      setSelection(activeR, activeC, r, c, false);
      selR2 = r; selC2 = c;
      paintSelection();
      updateStatus();
    } else {
      activeR = r; activeC = c;
      selR1 = selR2 = r; selC1 = selC2 = c;
      selecting = true;
      paintSelection();
      updateFormulaBar();
      updateStatus();
      syncRibbonFromActive();
    }
  });

  sheet.addEventListener("dblclick", (e) => {
    if (e.target.classList.contains("col-resize")) {
      autoFitCol(+e.target.dataset.col);
      return;
    }
    const cell = e.target.closest(".cell");
    if (!cell) return;
    activeR = +cell.dataset.r; activeC = +cell.dataset.c;
    setSelection(activeR, activeC, activeR, activeC);
    startEdit();
  });

  sheet.addEventListener("contextmenu", (e) => {
    e.preventDefault();
    const cell = e.target.closest(".cell");
    if (cell) {
      const r = +cell.dataset.r, c = +cell.dataset.c;
      const rg = getRange();
      if (r < rg.r1 || r > rg.r2 || c < rg.c1 || c > rg.c2) {
        activeR = r; activeC = c;
        setSelection(r, c, r, c);
      }
    }
    showCtx(e.clientX, e.clientY);
  });

  sheet.addEventListener("mouseover", (e) => {
    if (selecting) {
      const cell = e.target.closest(".cell");
      if (!cell) return;
      selR2 = +cell.dataset.r; selC2 = +cell.dataset.c;
      paintSelection();
      updateStatus();
    }
    if (filling) {
      const cell = e.target.closest(".cell");
      if (!cell) return;
      const toR = +cell.dataset.r, toC = +cell.dataset.c;
      const src = fillStart;
      setSelection(Math.min(src.r1, toR), Math.min(src.c1, toC), Math.max(src.r2, toR), Math.max(src.c2, toC), false);
    }
  });

  window.addEventListener("mousemove", (e) => {
    if (resizingCol == null) return;
    colWidths[resizingCol] = Math.max(36, resizeStartW + (e.clientX - resizeStartX));
    applyColTemplate();
    paintSelection();
  });

  window.addEventListener("mouseup", (e) => {
    selecting = false;
    if (filling) {
      filling = false;
      const cell = e.target.closest && e.target.closest(".cell");
      if (cell) applyFill(+cell.dataset.r, +cell.dataset.c);
      fillStart = null;
    }
    if (resizingCol != null) {
      document.querySelectorAll(".col-resize.dragging").forEach((el) => el.classList.remove("dragging"));
      resizingCol = null;
    }
  });

  fillHandle.addEventListener("mousedown", (e) => {
    e.preventDefault();
    e.stopPropagation();
    commitEdit();
    filling = true;
    fillStart = getRange();
  });

  document.addEventListener("click", (e) => {
    if (!ctxMenu.contains(e.target)) hideCtx();
  });

  // ── Keyboard ────────────────────────────────────────────
  document.addEventListener("keydown", (e) => {
    const tag = (e.target && e.target.tagName) || "";
    const inInput = tag === "INPUT" || tag === "TEXTAREA" || e.target.isContentEditable;
    const mod = e.ctrlKey || e.metaKey;

    if (mod && e.key.toLowerCase() === "o") { e.preventDefault(); fileInput.click(); return; }
    if (mod && e.key.toLowerCase() === "s") { e.preventDefault(); saveAs("auto"); return; }
    if (mod && e.key.toLowerCase() === "n") { e.preventDefault(); newWorkbook(); return; }
    if (mod && e.key.toLowerCase() === "f") { e.preventDefault(); openFind(false); return; }
    if (mod && e.key.toLowerCase() === "h") { e.preventDefault(); openFind(true); return; }
    if (mod && e.key.toLowerCase() === "z") { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
    if (mod && e.key.toLowerCase() === "y") { e.preventDefault(); redo(); return; }
    if (mod && e.key.toLowerCase() === "b" && !inInput) { e.preventDefault(); toggleStyle("bold"); return; }
    if (mod && e.key.toLowerCase() === "i" && !inInput) { e.preventDefault(); toggleStyle("italic"); return; }
    if (mod && e.key.toLowerCase() === "u" && !inInput) { e.preventDefault(); toggleStyle("underline"); return; }
    if (e.altKey && e.key === "=" && !inInput) { e.preventDefault(); autoSum(); return; }
    if (e.key === "F1") { e.preventDefault(); helpModal.classList.add("open"); return; }
    if (e.key === "F4" && !inInput) { e.preventDefault(); /* reserved */ return; }

    if (inInput && e.target !== formulaInput) {
      if (e.target === findInput || e.target === replaceInput) {
        if (e.key === "Enter") {
          e.preventDefault();
          if (e.target === replaceInput) replaceOne();
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
          pushUndo(); setCell(activeR, activeC, val); markDirty();
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
          pushUndo(); setCell(activeR, activeC, val); markDirty();
        }
        refreshRange(activeR, activeC, activeR, activeC);
        moveActive(0, e.shiftKey ? -1 : 1);
      }
      return;
    }

    if (editing) return;

    if (mod && e.key.toLowerCase() === "c") { e.preventDefault(); copySelection(false); return; }
    if (mod && e.key.toLowerCase() === "x") { e.preventDefault(); copySelection(true); return; }
    if (mod && e.key.toLowerCase() === "v") { e.preventDefault(); pasteSelection(); return; }
    if (mod && e.key.toLowerCase() === "a") { e.preventDefault(); selectAll(); return; }

    const arrow = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
    if (arrow) {
      e.preventDefault();
      if (e.shiftKey) {
        selR2 = Math.max(0, Math.min(rows - 1, selR2 + arrow[0]));
        selC2 = Math.max(0, Math.min(cols - 1, selC2 + arrow[1]));
        paintSelection();
        updateStatus();
      } else moveActive(arrow[0], arrow[1]);
      return;
    }
    if (e.key === "Tab") { e.preventDefault(); moveActive(0, e.shiftKey ? -1 : 1); return; }
    if (e.key === "Enter") {
      e.preventDefault();
      if (e.shiftKey) moveActive(-1, 0);
      else startEdit();
      return;
    }
    if (e.key === "F2") { e.preventDefault(); startEdit(); return; }
    if (e.key === "Delete") { e.preventDefault(); deleteSelection(false); return; }
    if (e.key === "Backspace") { e.preventDefault(); deleteSelection(false); startEdit(""); return; }
    if (e.key === "Escape") {
      setSelection(activeR, activeC, activeR, activeC);
      hideCtx();
      closeBackstage();
      return;
    }
    if (e.key === "Home") {
      e.preventDefault();
      if (mod) { activeR = 0; activeC = 0; } else activeC = 0;
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
    if (e.key === "PageDown") { e.preventDefault(); moveActive(20, 0); return; }
    if (e.key === "PageUp") { e.preventDefault(); moveActive(-20, 0); return; }

    if (!mod && e.key.length === 1 && !e.altKey) {
      e.preventDefault();
      startEdit(e.key);
    }
  });

  // ── Drag drop ───────────────────────────────────────────
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

  // ── Wire UI ─────────────────────────────────────────────
  document.querySelectorAll(".ribbon-tab").forEach((tab) => {
    tab.addEventListener("click", () => setRibbon(tab.dataset.ribbon));
  });
  document.getElementById("bsBack").onclick = () => closeBackstage();
  document.querySelectorAll(".backstage-nav button[data-bs]").forEach((b) => {
    b.addEventListener("click", () => openBackstage(b.dataset.bs));
  });

  document.getElementById("btnCut").onclick = () => copySelection(true);
  document.getElementById("btnCopy").onclick = () => copySelection(false);
  document.getElementById("btnPaste").onclick = () => pasteSelection();
  document.getElementById("btnFormatPainter").onclick = () => {
    formatPainter = formatPainter ? null : { ...getStyle(activeR, activeC) };
    syncRibbonFromActive();
    toast(formatPainter ? "Chọn ô để dán định dạng" : "Tắt sao chép định dạng");
  };

  document.getElementById("btnBold").onclick = () => toggleStyle("bold");
  document.getElementById("btnItalic").onclick = () => toggleStyle("italic");
  document.getElementById("btnUnderline").onclick = () => toggleStyle("underline");
  document.getElementById("btnAlignLeft").onclick = () => applyStyleToSelection({ align: "left" });
  document.getElementById("btnAlignCenter").onclick = () => applyStyleToSelection({ align: "center" });
  document.getElementById("btnAlignRight").onclick = () => applyStyleToSelection({ align: "right" });
  document.getElementById("btnWrap").onclick = () => toggleStyle("wrap");
  document.getElementById("btnMerge").onclick = () => toast("Gộp ô: căn giữa vùng chọn (demo)");
  document.getElementById("btnBorderAll").onclick = () => applyStyleToSelection({ border: "all" });
  document.getElementById("btnBorderBottom").onclick = () => applyStyleToSelection({ border: "bottom" });
  document.getElementById("btnBorderNone").onclick = () => applyStyleToSelection({ border: "" });
  document.getElementById("btnClearFmt").onclick = () => applyStyleToSelection(emptyStyle());
  document.getElementById("btnClearAll").onclick = () => deleteSelection(true);

  document.getElementById("fontFamily").onchange = (e) => applyStyleToSelection({ font: e.target.value });
  document.getElementById("fontSize").onchange = (e) => applyStyleToSelection({ size: +e.target.value });
  document.getElementById("numFmt").onchange = (e) => applyStyleToSelection({ numFmt: e.target.value });
  document.getElementById("textColor").oninput = (e) => {
    document.getElementById("swatchColor").style.background = e.target.value;
    applyStyleToSelection({ color: e.target.value });
  };
  document.getElementById("fillColor").oninput = (e) => {
    document.getElementById("swatchFill").style.background = e.target.value;
    applyStyleToSelection({ fill: e.target.value });
  };
  document.getElementById("btnPct").onclick = () => applyStyleToSelection({ numFmt: "percent" });
  document.getElementById("btnComma").onclick = () => applyStyleToSelection({ numFmt: "number", decimals: 0 });
  document.getElementById("btnDecMore").onclick = () => {
    const d = (getStyle(activeR, activeC).decimals || 2) + 1;
    applyStyleToSelection({ decimals: Math.min(d, 8), numFmt: getStyle(activeR, activeC).numFmt === "general" ? "number" : getStyle(activeR, activeC).numFmt });
  };
  document.getElementById("btnDecLess").onclick = () => {
    const d = Math.max(0, (getStyle(activeR, activeC).decimals || 2) - 1);
    applyStyleToSelection({ decimals: d });
  };

  document.getElementById("btnAutoSum").onclick = () => autoSum();
  document.getElementById("btnSortAsc").onclick = () => sortByActiveCol(true);
  document.getElementById("btnSortDesc").onclick = () => sortByActiveCol(false);
  document.getElementById("btnSortAsc2").onclick = () => sortByActiveCol(true);
  document.getElementById("btnSortDesc2").onclick = () => sortByActiveCol(false);
  document.getElementById("btnFind").onclick = () => openFind(false);
  document.getElementById("btnFind2").onclick = () => openFind(false);
  document.getElementById("btnReplace").onclick = () => openFind(true);

  document.getElementById("btnAddRow").onclick = () => addRow();
  document.getElementById("btnAddCol").onclick = () => addCol();
  document.getElementById("btnAddSheet").onclick = () => addWorkbookSheet();
  document.getElementById("btnDelRow").onclick = () => delRow();
  document.getElementById("btnDelCol").onclick = () => delCol();
  document.getElementById("btnLink").onclick = () => {
    const url = prompt("URL:", "https://");
    if (url) {
      pushUndo();
      setCell(activeR, activeC, url);
      setStyle(activeR, activeC, { color: "#0563c1", underline: true });
      refreshCell(activeR, activeC);
      markDirty();
    }
  };
  document.getElementById("btnComment").onclick = () => {
    const note = prompt("Ghi chú ô:", getStyle(activeR, activeC).note || "");
    if (note != null) {
      applyStyleToSelection({ note });
      toast(note ? "Đã gắn ghi chú" : "Đã xóa ghi chú");
    }
  };

  document.querySelectorAll("[data-fx]").forEach((el) => {
    el.addEventListener("click", () => insertFx(el.dataset.fx));
  });
  document.querySelectorAll("[data-tool]").forEach((el) => {
    el.addEventListener("click", () => applyCalcTool(el.dataset.tool));
  });
  document.getElementById("btnShowFormulas").onclick = () => {
    showFormulas = !showFormulas;
    refreshAll();
    syncRibbonFromActive();
    toast(showFormulas ? "Hiện công thức" : "Hiện kết quả");
  };
  document.getElementById("btnOpenCalc")?.addEventListener("click", () => openCalculator());
  document.getElementById("btnOpenCalc2")?.addEventListener("click", () => openCalculator());
  // shortcut Ctrl+Shift+C → calculator
  document.addEventListener("keydown", (e) => {
    if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key.toLowerCase() === "c") {
      e.preventDefault();
      openCalculator();
    }
  });

  document.getElementById("btnTextToCol").onclick = () => {
    commitEdit();
    pushUndo();
    const rg = getRange();
    for (let r = rg.r1; r <= rg.r2; r++) {
      const v = getRaw(r, rg.c1);
      if (!v || !v.includes(",")) continue;
      const parts = v.split(",");
      ensureSize(r, rg.c1 + parts.length);
      parts.forEach((p, i) => setCell(r, rg.c1 + i, p.trim()));
    }
    renderSheet();
    markDirty();
    toast("Đã tách cột theo dấu phẩy");
  };

  document.getElementById("btnFreeze").onclick = () => {
    freezeHeader = !freezeHeader;
    sheet.classList.toggle("freeze-header", freezeHeader);
    syncRibbonFromActive();
    toast(freezeHeader ? "Đã ghim hàng tiêu đề" : "Bỏ ghim");
  };
  document.getElementById("btnGridlines").onclick = () => {
    showGrid = !showGrid;
    sheet.classList.toggle("no-grid", !showGrid);
    syncRibbonFromActive();
  };
  document.getElementById("btnZoomIn").onclick = () => setZoom(zoom + 0.1);
  document.getElementById("btnZoomOut").onclick = () => setZoom(zoom - 0.1);
  document.getElementById("btnZoom100").onclick = () => setZoom(1);
  document.getElementById("zoomIn").onclick = () => setZoom(zoom + 0.1);
  document.getElementById("zoomOut").onclick = () => setZoom(zoom - 0.1);

  document.getElementById("btnTheme").onclick = () => {
    const dark = document.documentElement.getAttribute("data-theme") === "dark";
    if (dark) document.documentElement.removeAttribute("data-theme");
    else document.documentElement.setAttribute("data-theme", "dark");
    localStorage.setItem("htmlexxcel-theme", dark ? "light" : "dark");
  };
  document.getElementById("btnHelp").onclick = () => helpModal.classList.add("open");
  document.getElementById("helpClose").onclick = () => helpModal.classList.remove("open");
  helpModal.addEventListener("click", (e) => { if (e.target === helpModal) helpModal.classList.remove("open"); });

  document.getElementById("fxOk").onclick = () => {
    const val = formulaInput.value;
    if (getRaw(activeR, activeC) !== val) {
      pushUndo(); setCell(activeR, activeC, val); markDirty();
    }
    refreshRange(activeR, activeC, activeR, activeC);
    updateFormulaBar();
  };
  document.getElementById("fxCancel").onclick = () => {
    formulaInput.value = getRaw(activeR, activeC);
    cancelEdit();
  };

  document.getElementById("findClose").onclick = () => closeFind();
  document.getElementById("findNext").onclick = () => runFind(1);
  document.getElementById("findPrev").onclick = () => runFind(-1);
  document.getElementById("btnReplaceOne").onclick = () => replaceOne();
  document.getElementById("btnReplaceAll").onclick = () => replaceAll();

  document.getElementById("btnOpenHint").onclick = () => fileInput.click();
  document.getElementById("btnSample").onclick = () => loadSample();
  document.getElementById("btnBlank").onclick = () => { newWorkbook(true); };

  fileInput.addEventListener("change", () => {
    const f = fileInput.files[0];
    if (f) openFile(f);
    fileInput.value = "";
  });
  if (typeof IO !== "undefined") fileInput.setAttribute("accept", IO.ACCEPT);

  ctxMenu.querySelectorAll("button[data-act]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const a = btn.dataset.act;
      hideCtx();
      if (a === "cut") copySelection(true);
      else if (a === "copy") copySelection(false);
      else if (a === "paste") pasteSelection();
      else if (a === "insertRow") addRow();
      else if (a === "insertCol") addCol();
      else if (a === "deleteRow") delRow();
      else if (a === "deleteCol") delCol();
      else if (a === "clear") deleteSelection(false);
      else if (a === "clearFmt") applyStyleToSelection(emptyStyle());
      else if (a === "autoSum") autoSum();
    });
  });

  window.addEventListener("resize", () => paintSelection());
  window.addEventListener("beforeunload", (e) => {
    if (dirty) { e.preventDefault(); e.returnValue = ""; }
  });

  // ── Init ────────────────────────────────────────────────
  (function init() {
    if (localStorage.getItem("htmlexxcel-theme") === "dark") {
      document.documentElement.setAttribute("data-theme", "dark");
    }
    const padded = padGrid([[""]]);
    workbookSheets = [{
      name: "Sheet1",
      data: padded.data,
      styles: padded.styles,
      rows: padded.rows,
      cols: padded.cols,
      colWidths: padded.colWidths,
    }];
    activeSheetIndex = 0;
    fileName = "";
    fileFormat = "xlsx";
    applySheetState(workbookSheets[0]);
    renderSheetTabs();
    hint.classList.add("show");
  })();
})();
