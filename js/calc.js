/**
 * Floating calculator for HTML Excel
 */
(function (global) {
  "use strict";

  function createCalculator(opts) {
    opts = opts || {};
    const onInsert = opts.onInsert || function () {};
    const onToast = opts.onToast || function () {};

    let expr = "";
    let display = "0";
    let memory = 0;
    let lastResult = null;
    let history = [];
    let scientific = false;

    const panel = document.getElementById("calcPanel");
    const dispEl = document.getElementById("calcDisplay");
    const exprEl = document.getElementById("calcExpr");
    const histEl = document.getElementById("calcHistory");
    const modeBtn = document.getElementById("calcMode");
    const sciRow = document.getElementById("calcSci");

    if (!panel) return { open() {}, close() {}, toggle() {} };

    function render() {
      if (dispEl) dispEl.textContent = display;
      if (exprEl) exprEl.textContent = expr || "\u00a0";
      if (histEl) {
        histEl.innerHTML = history
          .slice(-8)
          .reverse()
          .map((h) => `<div class="calc-hist-item" data-v="${escapeAttr(String(h.result))}"><span>${escapeHtml(h.expr)}</span><b>${escapeHtml(String(h.result))}</b></div>`)
          .join("");
        histEl.querySelectorAll(".calc-hist-item").forEach((el) => {
          el.addEventListener("click", () => {
            display = el.dataset.v || "0";
            expr = "";
            lastResult = Number(display);
            render();
          });
        });
      }
      if (sciRow) sciRow.style.display = scientific ? "grid" : "none";
      if (modeBtn) modeBtn.textContent = scientific ? "Cơ bản" : "Khoa học";
    }

    function escapeHtml(s) {
      return s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
    }
    function escapeAttr(s) {
      return s.replace(/"/g, "&quot;");
    }

    function safeEval(str) {
      // Allow digits, ops, parentheses, decimal, %, ^, functions
      let s = String(str).trim();
      if (!s) return 0;
      s = s.replace(/×/g, "*").replace(/÷/g, "/").replace(/−/g, "-");
      s = s.replace(/%/g, "/100");
      s = s.replace(/\^/g, "**");
      s = s.replace(/\bpi\b/gi, "Math.PI");
      s = s.replace(/\bsqrt\s*\(/gi, "Math.sqrt(");
      s = s.replace(/\bsin\s*\(/gi, "Math.sin(");
      s = s.replace(/\bcos\s*\(/gi, "Math.cos(");
      s = s.replace(/\btan\s*\(/gi, "Math.tan(");
      s = s.replace(/\blog\s*\(/gi, "Math.log10(");
      s = s.replace(/\bln\s*\(/gi, "Math.log(");
      s = s.replace(/\babs\s*\(/gi, "Math.abs(");
      s = s.replace(/\bfact\s*\(/gi, "__fact(");
      if (/[^0-9+\-*/().%\s,*a-zA-Z_]/.test(s.replace(/Math\./g, "").replace(/__fact/g, "f"))) {
        throw new Error("Biểu thức không hợp lệ");
      }
      const __fact = (n) => {
        n = Math.floor(n);
        if (n < 0 || n > 170) throw new Error("fact");
        let f = 1;
        for (let i = 2; i <= n; i++) f *= i;
        return f;
      };
      // eslint-disable-next-line no-new-func
      const fn = new Function("__fact", `"use strict"; return (${s});`);
      const v = fn(__fact);
      if (typeof v !== "number" || !isFinite(v)) throw new Error("Kết quả không hợp lệ");
      return v;
    }

    function formatNum(n) {
      if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
      const t = Math.round(n * 1e12) / 1e12;
      return String(t);
    }

    function input(ch) {
      if (display === "Error") {
        display = "0";
        expr = "";
      }
      if (ch === "C") {
        expr = "";
        display = "0";
        lastResult = null;
        render();
        return;
      }
      if (ch === "CE") {
        display = "0";
        render();
        return;
      }
      if (ch === "BS") {
        if (display.length <= 1 || display === "0") display = "0";
        else display = display.slice(0, -1);
        render();
        return;
      }
      if (ch === "=") {
        try {
          const full = expr ? expr + display : display;
          const result = safeEval(full);
          history.push({ expr: full + " =", result: formatNum(result) });
          if (history.length > 50) history.shift();
          display = formatNum(result);
          expr = "";
          lastResult = result;
        } catch {
          display = "Error";
          expr = "";
        }
        render();
        return;
      }
      if (ch === "±") {
        if (display.startsWith("-")) display = display.slice(1);
        else if (display !== "0") display = "-" + display;
        render();
        return;
      }
      if ("+-*/^".includes(ch)) {
        const op = ch === "^" ? "^" : ch;
        if (expr && /[+\-*/^]$/.test(expr.trim()) && display === "0") {
          expr = expr.trim().slice(0, -1) + " " + op + " ";
        } else {
          expr = (expr || "") + display + " " + op + " ";
          display = "0";
        }
        render();
        return;
      }
      if (ch === ".") {
        if (!display.includes(".")) display += ".";
        render();
        return;
      }
      if (ch === "(" || ch === ")") {
        if (display === "0" && !expr) {
          expr = ch;
          display = "0";
        } else {
          expr = (expr || "") + display + ch;
          display = "0";
        }
        render();
        return;
      }
      if (ch === "%") {
        try {
          const v = safeEval(display) / 100;
          display = formatNum(v);
        } catch {
          display = "Error";
        }
        render();
        return;
      }
      // scientific function applied to display
      if (["sqrt", "sin", "cos", "tan", "log", "ln", "abs", "sq", "inv", "fact"].includes(ch)) {
        try {
          let v = safeEval(display);
          if (ch === "sqrt") v = Math.sqrt(v);
          else if (ch === "sin") v = Math.sin(v);
          else if (ch === "cos") v = Math.cos(v);
          else if (ch === "tan") v = Math.tan(v);
          else if (ch === "log") v = Math.log10(v);
          else if (ch === "ln") v = Math.log(v);
          else if (ch === "abs") v = Math.abs(v);
          else if (ch === "sq") v = v * v;
          else if (ch === "inv") {
            if (v === 0) throw new Error("div0");
            v = 1 / v;
          } else if (ch === "fact") {
            v = Math.floor(v);
            if (v < 0 || v > 170) throw new Error("fact");
            let f = 1;
            for (let i = 2; i <= v; i++) f *= i;
            v = f;
          }
          history.push({ expr: ch + "(" + display + ")", result: formatNum(v) });
          display = formatNum(v);
          lastResult = v;
        } catch {
          display = "Error";
        }
        render();
        return;
      }
      if (ch === "pi") {
        display = formatNum(Math.PI);
        render();
        return;
      }
      // digit
      if (display === "0" && ch !== ".") display = ch;
      else display += ch;
      render();
    }

    function mem(op) {
      const v = Number(display);
      if (!isFinite(v) && op !== "MC") return;
      if (op === "MC") memory = 0;
      else if (op === "MR") display = formatNum(memory);
      else if (op === "M+") memory += v;
      else if (op === "M-") memory -= v;
      else if (op === "MS") memory = v;
      render();
      onToast(op === "MC" ? "Đã xóa bộ nhớ" : "Bộ nhớ: " + formatNum(memory));
    }

    // keypad
    panel.querySelectorAll("[data-calc]").forEach((btn) => {
      btn.addEventListener("click", () => input(btn.dataset.calc));
    });
    panel.querySelectorAll("[data-mem]").forEach((btn) => {
      btn.addEventListener("click", () => mem(btn.dataset.mem));
    });

    document.getElementById("calcClose")?.addEventListener("click", () => close());
    document.getElementById("calcInsert")?.addEventListener("click", () => {
      try {
        const v = display === "Error" ? null : safeEval(display);
        if (v == null || !isFinite(v)) return onToast("Không có kết quả hợp lệ");
        onInsert(formatNum(v));
        onToast("Đã chèn vào ô: " + formatNum(v));
      } catch {
        onToast("Không chèn được");
      }
    });
    document.getElementById("calcInsertExpr")?.addEventListener("click", () => {
      const full = (expr + display).trim();
      if (!full || full === "0") return onToast("Biểu thức trống");
      // convert to spreadsheet formula
      let f = full.replace(/×/g, "*").replace(/÷/g, "/");
      if (!f.startsWith("=")) f = "=" + f;
      onInsert(f);
      onToast("Đã chèn công thức: " + f);
    });
    modeBtn?.addEventListener("click", () => {
      scientific = !scientific;
      render();
    });

    // drag
    const head = panel.querySelector(".calc-head");
    let drag = null;
    head?.addEventListener("mousedown", (e) => {
      if (e.target.closest("button")) return;
      const rect = panel.getBoundingClientRect();
      drag = { x: e.clientX - rect.left, y: e.clientY - rect.top };
      e.preventDefault();
    });
    window.addEventListener("mousemove", (e) => {
      if (!drag) return;
      panel.style.left = Math.max(0, e.clientX - drag.x) + "px";
      panel.style.top = Math.max(0, e.clientY - drag.y) + "px";
      panel.style.right = "auto";
      panel.style.bottom = "auto";
    });
    window.addEventListener("mouseup", () => { drag = null; });

    function open() {
      panel.classList.add("open");
      render();
    }
    function close() {
      panel.classList.remove("open");
    }
    function toggle() {
      if (panel.classList.contains("open")) close();
      else open();
    }

    // keyboard when open
    window.addEventListener("keydown", (e) => {
      if (!panel.classList.contains("open")) return;
      if (e.target.tagName === "INPUT" || e.target.tagName === "TEXTAREA") return;
      const k = e.key;
      if (k >= "0" && k <= "9") { e.preventDefault(); input(k); }
      else if (k === ".") { e.preventDefault(); input("."); }
      else if (k === "+" || k === "-" || k === "*" || k === "/") { e.preventDefault(); input(k); }
      else if (k === "Enter" || k === "=") { e.preventDefault(); input("="); }
      else if (k === "Escape") { e.preventDefault(); close(); }
      else if (k === "Backspace") { e.preventDefault(); input("BS"); }
      else if (k === "Delete") { e.preventDefault(); input("CE"); }
      else if (k === "(" || k === ")") { e.preventDefault(); input(k); }
      else if (k === "%") { e.preventDefault(); input("%"); }
    });

    render();
    return { open, close, toggle, input, getResult: () => lastResult };
  }

  global.CalcUI = { createCalculator };
})(typeof window !== "undefined" ? window : globalThis);
