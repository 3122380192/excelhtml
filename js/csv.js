/**
 * CSV / TSV parse & serialize (Excel / Google Sheets compatible)
 */
(function (global) {
  "use strict";

  function detectDelimiter(text) {
    const firstLine = text.split(/\r?\n/).find((l) => l.trim()) || "";
    const tabs = (firstLine.match(/\t/g) || []).length;
    const semis = (firstLine.match(/;/g) || []).length;
    const commas = (firstLine.match(/,/g) || []).length;
    if (tabs > commas && tabs >= semis) return "\t";
    if (semis > commas && semis > tabs) return ";";
    return ",";
  }

  function parseCSV(text) {
    if (text == null) return [[""]];
    text = String(text);
    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);

    const delim = detectDelimiter(text);
    const rowsOut = [];
    let row = [];
    let field = "";
    let i = 0;
    let inQuotes = false;

    while (i < text.length) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') {
          if (text[i + 1] === '"') {
            field += '"';
            i += 2;
            continue;
          }
          inQuotes = false;
          i++;
          continue;
        }
        field += ch;
        i++;
        continue;
      }
      if (ch === '"') {
        inQuotes = true;
        i++;
        continue;
      }
      if (ch === delim) {
        row.push(field);
        field = "";
        i++;
        continue;
      }
      if (ch === "\r") {
        i++;
        continue;
      }
      if (ch === "\n") {
        row.push(field);
        rowsOut.push(row);
        row = [];
        field = "";
        i++;
        continue;
      }
      field += ch;
      i++;
    }
    if (field.length || row.length) {
      row.push(field);
      rowsOut.push(row);
    }
    while (rowsOut.length && rowsOut[rowsOut.length - 1].every((c) => c === "")) {
      rowsOut.pop();
    }
    return rowsOut.length ? rowsOut : [[""]];
  }

  function toCSV(grid, delim) {
    delim = delim || ",";
    return grid
      .map((row) =>
        row
          .map((cell) => {
            const s = cell == null ? "" : String(cell);
            if (s.includes('"') || s.includes(delim) || /[\n\r]/.test(s)) {
              return '"' + s.replace(/"/g, '""') + '"';
            }
            return s;
          })
          .join(delim)
      )
      .join("\r\n");
  }

  /** Range → TSV for clipboard (Excel / Sheets) */
  function rangeToTSV(getCell, r1, c1, r2, c2) {
    const lines = [];
    for (let r = r1; r <= r2; r++) {
      const parts = [];
      for (let c = c1; c <= c2; c++) {
        parts.push(getCell(r, c) ?? "");
      }
      lines.push(parts.join("\t"));
    }
    return lines.join("\n");
  }

  /**
   * Parse clipboard TSV/CSV text into 2D array
   */
  function parseClipboard(text) {
    if (text == null) return [[""]];
    let t = String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    if (t.endsWith("\n")) t = t.slice(0, -1);
    if (t === "") return [[""]];

    // Prefer TSV when tabs present
    if (t.includes("\t")) {
      return t.split("\n").map((line) =>
        line.split("\t").map((v) => {
          if (v.length >= 2 && v.startsWith('"') && v.endsWith('"')) {
            return v.slice(1, -1).replace(/""/g, '"');
          }
          return v;
        })
      );
    }
    // Single line with commas → maybe CSV row
    if (t.includes(",") && !t.includes("\n")) {
      return parseCSV(t);
    }
    // Multi-line plain text → one cell per line (or CSV if looks like it)
    if (t.includes(",") && t.includes("\n")) {
      return parseCSV(t);
    }
    return t.split("\n").map((line) => [line]);
  }

  global.CSV = {
    parseCSV,
    toCSV,
    rangeToTSV,
    parseClipboard,
    detectDelimiter,
  };
})(typeof window !== "undefined" ? window : globalThis);
