/**
 * Multi-format spreadsheet I/O
 * Supports: CSV, TSV, TXT, XLSX, XLS, XLSM, ODS, HTML tables
 * Requires: CSV (csv.js) + XLSX (SheetJS vendor)
 */
(function (global) {
  "use strict";

  const TEXT_EXTS = new Set(["csv", "tsv", "txt", "tab"]);
  const EXCEL_EXTS = new Set(["xlsx", "xls", "xlsm", "xlsb", "ods", "fods", "xltx", "xltm"]);
  const HTML_EXTS = new Set(["html", "htm"]);

  const ACCEPT =
    ".csv,.tsv,.txt,.tab,.xlsx,.xls,.xlsm,.xlsb,.ods,.fods,.html,.htm," +
    "text/csv,text/tab-separated-values,text/plain," +
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet," +
    "application/vnd.ms-excel," +
    "application/vnd.oasis.opendocument.spreadsheet";

  function extOf(name) {
    const m = String(name || "").toLowerCase().match(/\.([a-z0-9]+)$/);
    return m ? m[1] : "";
  }

  function baseName(name) {
    return String(name || "workbook").replace(/\.[^.]+$/, "") || "workbook";
  }

  function ensureXLSX() {
    if (typeof global.XLSX === "undefined") {
      throw new Error("Thiếu thư viện SheetJS (xlsx). Kiểm tra js/vendor/xlsx.full.min.js");
    }
    return global.XLSX;
  }

  /** Worksheet → 2D string grid (giữ công thức nếu có) */
  function sheetToGrid(ws) {
    const XLSX = ensureXLSX();
    if (!ws || !ws["!ref"]) return [[""]];
    const range = XLSX.utils.decode_range(ws["!ref"]);
    // Guard huge sheets
    const maxR = Math.min(range.e.r, range.s.r + 5000);
    const maxC = Math.min(range.e.c, range.s.c + 200);
    const grid = [];
    for (let R = range.s.r; R <= maxR; R++) {
      const row = [];
      for (let C = range.s.c; C <= maxC; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = ws[addr];
        if (!cell) {
          row.push("");
          continue;
        }
        if (cell.f) {
          row.push("=" + cell.f);
        } else if (cell.v != null && cell.v !== "") {
          if (cell.t === "d" && cell.v instanceof Date) {
            row.push(cell.v.toLocaleDateString());
          } else if (cell.t === "b") {
            row.push(cell.v ? "TRUE" : "FALSE");
          } else {
            row.push(String(cell.v));
          }
        } else if (cell.w != null) {
          row.push(String(cell.w));
        } else {
          row.push("");
        }
      }
      grid.push(row);
    }
    // trim trailing empty rows
    while (grid.length && grid[grid.length - 1].every((c) => c === "")) grid.pop();
    return grid.length ? grid : [[""]];
  }

  function gridToSheet(grid) {
    const XLSX = ensureXLSX();
    // Build sheet manually so formulas stay as formulas
    const ws = {};
    let maxR = 0,
      maxC = 0;
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r] || [];
      for (let c = 0; c < row.length; c++) {
        const val = row[c];
        if (val == null || val === "") continue;
        maxR = Math.max(maxR, r);
        maxC = Math.max(maxC, c);
        const addr = XLSX.utils.encode_cell({ r, c });
        const s = String(val);
        if (s.startsWith("=")) {
          ws[addr] = { t: "n", f: s.slice(1), v: 0 };
        } else {
          const n = Number(String(s).replace(/,/g, ""));
          if (s.trim() !== "" && !isNaN(n) && isFinite(n) && /^-?[\d.,]+$/.test(s.trim())) {
            ws[addr] = { t: "n", v: n };
          } else if (s.toUpperCase() === "TRUE" || s.toUpperCase() === "FALSE") {
            ws[addr] = { t: "b", v: s.toUpperCase() === "TRUE" };
          } else {
            ws[addr] = { t: "s", v: s };
          }
        }
      }
    }
    ws["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: Math.max(maxR, 0), c: Math.max(maxC, 0) },
    });
    return ws;
  }

  function parseWorkbookArrayBuffer(buf, fileName) {
    const XLSX = ensureXLSX();
    const wb = XLSX.read(buf, { type: "array", cellDates: true, cellNF: false, cellText: false });
    const sheets = [];
    for (const name of wb.SheetNames) {
      sheets.push({
        name: name || "Sheet" + (sheets.length + 1),
        grid: sheetToGrid(wb.Sheets[name]),
      });
    }
    if (!sheets.length) sheets.push({ name: "Sheet1", grid: [[""]] });
    return { sheets, fileName: fileName || "workbook.xlsx", format: extOf(fileName) || "xlsx" };
  }

  function parseTextWorkbook(text, fileName) {
    const ext = extOf(fileName);
    let grid;
    if (ext === "tsv" || ext === "tab") {
      let t = String(text);
      if (t.charCodeAt(0) === 0xfeff) t = t.slice(1);
      t = t.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
      if (t.endsWith("\n")) t = t.slice(0, -1);
      grid = t === "" ? [[""]] : t.split("\n").map((line) => line.split("\t"));
      while (grid.length && grid[grid.length - 1].every((c) => c === "")) grid.pop();
      if (!grid.length) grid = [[""]];
    } else {
      grid = global.CSV.parseCSV(text);
    }
    return {
      sheets: [{ name: "Sheet1", grid }],
      fileName: fileName || "data.csv",
      format: ext === "tab" ? "tsv" : ext || "csv",
    };
  }

  function parseHtmlTable(html, fileName) {
    const XLSX = ensureXLSX();
    try {
      const wb = XLSX.read(html, { type: "string", raw: true });
      const sheets = wb.SheetNames.map((name) => ({
        name,
        grid: sheetToGrid(wb.Sheets[name]),
      }));
      if (sheets.length) {
        return { sheets, fileName: fileName || "table.html", format: "html" };
      }
    } catch (_) {
      /* fall through */
    }
    // fallback: DOM parse first table
    const doc = new DOMParser().parseFromString(html, "text/html");
    const table = doc.querySelector("table");
    if (!table) throw new Error("Không tìm thấy bảng trong HTML");
    const grid = [];
    table.querySelectorAll("tr").forEach((tr) => {
      const row = [];
      tr.querySelectorAll("th,td").forEach((td) => row.push((td.textContent || "").trim()));
      if (row.length) grid.push(row);
    });
    return {
      sheets: [{ name: "Sheet1", grid: grid.length ? grid : [[""]] }],
      fileName: fileName || "table.html",
      format: "html",
    };
  }

  /**
   * Read a File / Blob into workbook structure
   * @returns {Promise<{sheets: Array<{name:string, grid:string[][]}>, fileName:string, format:string}>}
   */
  function readFile(file) {
    const name = file.name || "file";
    const ext = extOf(name);

    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error("Không đọc được file"));

      if (EXCEL_EXTS.has(ext) || (!ext && file.type && file.type.includes("sheet"))) {
        reader.onload = () => {
          try {
            resolve(parseWorkbookArrayBuffer(new Uint8Array(reader.result), name));
          } catch (e) {
            reject(e);
          }
        };
        reader.readAsArrayBuffer(file);
        return;
      }

      if (HTML_EXTS.has(ext) || (file.type && file.type.includes("html"))) {
        reader.onload = () => {
          try {
            resolve(parseHtmlTable(String(reader.result), name));
          } catch (e) {
            reject(e);
          }
        };
        reader.readAsText(file, "UTF-8");
        return;
      }

      // default: text CSV/TSV/TXT — also try excel if binary sniff fails
      reader.onload = () => {
        try {
          const text = String(reader.result);
          // ZIP/XLS magic: if user picked wrong ext
          if (text.charCodeAt(0) === 0x50 && text.charCodeAt(1) === 0x4b) {
            // re-read as array buffer
            const r2 = new FileReader();
            r2.onload = () => {
              try {
                resolve(parseWorkbookArrayBuffer(new Uint8Array(r2.result), name));
              } catch (e) {
                reject(e);
              }
            };
            r2.onerror = () => reject(new Error("Không đọc được file Excel"));
            r2.readAsArrayBuffer(file);
            return;
          }
          resolve(parseTextWorkbook(text, name));
        } catch (e) {
          reject(e);
        }
      };
      reader.readAsText(file, "UTF-8");
    });
  }

  function downloadBlob(blob, fileName) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = fileName;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }

  function usedGrid(grid) {
    let maxR = 0,
      maxC = 0;
    for (let r = 0; r < grid.length; r++) {
      for (let c = 0; c < (grid[r] || []).length; c++) {
        if (grid[r][c] != null && grid[r][c] !== "") {
          maxR = Math.max(maxR, r);
          maxC = Math.max(maxC, c);
        }
      }
    }
    const out = [];
    for (let r = 0; r <= maxR; r++) {
      const row = [];
      for (let c = 0; c <= maxC; c++) row.push((grid[r] && grid[r][c]) || "");
      out.push(row);
    }
    return out.length ? out : [[""]];
  }

  /**
   * @param {Array<{name:string, grid:string[][]}>} sheets
   * @param {string} fileName
   * @param {'csv'|'tsv'|'xlsx'|'xls'|'ods'} format
   */
  function saveWorkbook(sheets, fileName, format) {
    format = (format || "csv").toLowerCase();
    const base = baseName(fileName);

    if (format === "csv" || format === "tsv" || format === "txt") {
      const grid = usedGrid((sheets[0] && sheets[0].grid) || [[""]]);
      const delim = format === "tsv" || format === "tab" ? "\t" : ",";
      const text = global.CSV.toCSV(grid, delim);
      const ext = format === "txt" ? "txt" : format === "tsv" ? "tsv" : "csv";
      const mime =
        ext === "csv" ? "text/csv;charset=utf-8" : "text/tab-separated-values;charset=utf-8";
      downloadBlob(new Blob(["\uFEFF" + text], { type: mime }), base + "." + ext);
      return base + "." + ext;
    }

    const XLSX = ensureXLSX();
    const wb = XLSX.utils.book_new();
    const list = sheets && sheets.length ? sheets : [{ name: "Sheet1", grid: [[""]] }];
    const usedNames = new Set();
    list.forEach((sh, i) => {
      let name = (sh.name || "Sheet" + (i + 1)).slice(0, 31) || "Sheet" + (i + 1);
      // Excel sheet name rules
      name = name.replace(/[\\/?*[\]]/g, "_");
      let unique = name;
      let n = 1;
      while (usedNames.has(unique.toLowerCase())) {
        unique = (name.slice(0, 28) + "_" + n).slice(0, 31);
        n++;
      }
      usedNames.add(unique.toLowerCase());
      const ws = gridToSheet(usedGrid(sh.grid || [[""]]));
      XLSX.utils.book_append_sheet(wb, ws, unique);
    });

    let bookType = "xlsx";
    let ext = "xlsx";
    if (format === "xls") {
      bookType = "xls";
      ext = "xls";
    } else if (format === "ods") {
      bookType = "ods";
      ext = "ods";
    } else if (format === "xlsm") {
      bookType = "xlsx";
      ext = "xlsx";
    }

    const out = XLSX.write(wb, { bookType, type: "array" });
    const mime =
      ext === "xls"
        ? "application/vnd.ms-excel"
        : ext === "ods"
          ? "application/vnd.oasis.opendocument.spreadsheet"
          : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    downloadBlob(new Blob([out], { type: mime }), base + "." + ext);
    return base + "." + ext;
  }

  function formatLabel(fmt) {
    const m = {
      csv: "CSV",
      tsv: "TSV",
      txt: "TXT",
      xlsx: "Excel (.xlsx)",
      xls: "Excel 97-2003 (.xls)",
      xlsm: "Excel macro (.xlsm)",
      ods: "OpenDocument (.ods)",
      html: "HTML",
    };
    return m[fmt] || (fmt || "").toUpperCase();
  }

  global.IO = {
    ACCEPT,
    TEXT_EXTS,
    EXCEL_EXTS,
    HTML_EXTS,
    extOf,
    baseName,
    readFile,
    saveWorkbook,
    formatLabel,
    usedGrid,
    sheetToGrid,
  };
})(typeof window !== "undefined" ? window : globalThis);
