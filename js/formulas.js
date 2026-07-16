/**
 * Lightweight spreadsheet formula engine
 * Supports: =A1, =A1+B2, =SUM(A1:B3), AVERAGE, COUNT, MIN, MAX, ROUND, IF, ABS, LEN, CONCAT
 */
(function (global) {
  "use strict";

  const CELL_REF = /^\$?([A-Za-z]+)\$?(\d+)$/;
  const RANGE_REF = /^\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$/;

  function colToIndex(letters) {
    let n = 0;
    const s = letters.toUpperCase();
    for (let i = 0; i < s.length; i++) {
      n = n * 26 + (s.charCodeAt(i) - 64);
    }
    return n - 1;
  }

  function parseNumber(v) {
    if (v == null || v === "") return null;
    if (typeof v === "number") return isFinite(v) ? v : null;
    const s = String(v).trim().replace(/,/g, "");
    if (s === "") return null;
    // percent
    if (s.endsWith("%")) {
      const n = Number(s.slice(0, -1));
      return isFinite(n) ? n / 100 : null;
    }
    const n = Number(s);
    return isFinite(n) ? n : null;
  }

  /**
   * @param {string} raw - cell raw value
   * @param {(r:number,c:number)=>string} getRaw - raw cell getter (no formula eval recursion issues via stack)
   * @param {Set<string>} [stack]
   */
  function evaluate(raw, getRaw, stack) {
    if (raw == null || raw === "") return { display: "", kind: "empty" };
    const s = String(raw);
    if (!s.startsWith("=")) {
      const num = parseNumber(s);
      if (num !== null && String(s).trim() !== "") {
        return { display: s, value: num, kind: "number" };
      }
      return { display: s, value: s, kind: "text" };
    }

    stack = stack || new Set();
    const formula = s.slice(1).trim();
    if (!formula) return { display: "#ERROR!", value: null, kind: "error" };

    try {
      const result = evalExpr(formula, getRaw, stack);
      if (result == null || (typeof result === "number" && !isFinite(result))) {
        return { display: "#NUM!", value: null, kind: "error" };
      }
      if (typeof result === "number") {
        const display =
          Math.abs(result) >= 1e10 || (Math.abs(result) > 0 && Math.abs(result) < 1e-6)
            ? result.toExponential(6)
            : formatNum(result);
        return { display, value: result, kind: "formula" };
      }
      if (typeof result === "boolean") {
        return { display: result ? "TRUE" : "FALSE", value: result, kind: "formula" };
      }
      return { display: String(result), value: result, kind: "formula" };
    } catch (e) {
      const msg = e && e.message ? e.message : "#ERROR!";
      return { display: msg.startsWith("#") ? msg : "#ERROR!", value: null, kind: "error" };
    }
  }

  function formatNum(n) {
    if (Number.isInteger(n)) return String(n);
    // trim floating noise
    const t = Math.round(n * 1e10) / 1e10;
    return String(t);
  }

  function evalExpr(expr, getRaw, stack) {
    // Tokenize & shunting-yard for + - * / ^ ( ) , functions
    const tokens = tokenize(expr);
    return evalTokens(tokens, getRaw, stack);
  }

  function tokenize(expr) {
    const tokens = [];
    let i = 0;
    const s = expr;
    while (i < s.length) {
      const ch = s[i];
      if (/\s/.test(ch)) {
        i++;
        continue;
      }
      if (/[0-9.]/.test(ch)) {
        let j = i + 1;
        while (j < s.length && /[0-9.]/.test(s[j])) j++;
        tokens.push({ type: "num", value: Number(s.slice(i, j)) });
        i = j;
        continue;
      }
      if (/[A-Za-z_$]/.test(ch)) {
        let j = i + 1;
        while (j < s.length && /[A-Za-z0-9_$]/.test(s[j])) j++;
        // range or cell? peek for :
        let word = s.slice(i, j);
        if (s[j] === ":") {
          let k = j + 1;
          while (k < s.length && /[A-Za-z0-9$]/.test(s[k])) k++;
          word = s.slice(i, k);
          tokens.push({ type: "range", value: word });
          i = k;
          continue;
        }
        // cell ref like A1
        if (CELL_REF.test(word) && !/^[A-Za-z]+$/.test(word)) {
          tokens.push({ type: "cell", value: word });
          i = j;
          continue;
        }
        // function name or bare identifier
        if (CELL_REF.test(word)) {
          tokens.push({ type: "cell", value: word });
        } else {
          tokens.push({ type: "func", value: word.toUpperCase() });
        }
        i = j;
        continue;
      }
      if (ch === '"') {
        let j = i + 1;
        let str = "";
        while (j < s.length && s[j] !== '"') {
          str += s[j++];
        }
        if (s[j] !== '"') throw new Error("#VALUE!");
        tokens.push({ type: "str", value: str });
        i = j + 1;
        continue;
      }
      if ("+-*/^(),:<>&!".includes(ch)) {
        // multi-char compare
        if ((ch === "<" || ch === ">" || ch === "!") && s[i + 1] === "=") {
          tokens.push({ type: "op", value: ch + "=" });
          i += 2;
          continue;
        }
        if (ch === "<" && s[i + 1] === ">") {
          tokens.push({ type: "op", value: "<>" });
          i += 2;
          continue;
        }
        tokens.push({ type: "op", value: ch });
        i++;
        continue;
      }
      throw new Error("#ERROR!");
    }
    return tokens;
  }

  function evalTokens(tokens, getRaw, stack) {
    // Convert func + ( ... ) into function calls via recursive parse
    let pos = 0;

    function peek() {
      return tokens[pos];
    }
    function next() {
      return tokens[pos++];
    }

    function parseComparison() {
      let left = parseAdd();
      while (peek() && peek().type === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(peek().value)) {
        const op = next().value;
        const right = parseAdd();
        if (op === "=") left = left == right; // eslint-disable-line eqeqeq
        else if (op === "<>") left = left != right; // eslint-disable-line eqeqeq
        else if (op === "<") left = Number(left) < Number(right);
        else if (op === ">") left = Number(left) > Number(right);
        else if (op === "<=") left = Number(left) <= Number(right);
        else if (op === ">=") left = Number(left) >= Number(right);
      }
      return left;
    }

    function parseAdd() {
      let left = parseMul();
      while (peek() && peek().type === "op" && (peek().value === "+" || peek().value === "-" || peek().value === "&")) {
        const op = next().value;
        const right = parseMul();
        if (op === "&") left = String(left ?? "") + String(right ?? "");
        else if (op === "+") left = Number(left) + Number(right);
        else left = Number(left) - Number(right);
      }
      return left;
    }

    function parseMul() {
      let left = parsePow();
      while (peek() && peek().type === "op" && (peek().value === "*" || peek().value === "/")) {
        const op = next().value;
        const right = parsePow();
        if (op === "*") left = Number(left) * Number(right);
        else {
          if (Number(right) === 0) throw new Error("#DIV/0!");
          left = Number(left) / Number(right);
        }
      }
      return left;
    }

    function parsePow() {
      let left = parseUnary();
      while (peek() && peek().type === "op" && peek().value === "^") {
        next();
        const right = parseUnary();
        left = Math.pow(Number(left), Number(right));
      }
      return left;
    }

    function parseUnary() {
      if (peek() && peek().type === "op" && peek().value === "-") {
        next();
        return -Number(parseUnary());
      }
      if (peek() && peek().type === "op" && peek().value === "+") {
        next();
        return Number(parseUnary());
      }
      return parsePrimary();
    }

    function parsePrimary() {
      const t = peek();
      if (!t) throw new Error("#ERROR!");

      if (t.type === "num") {
        next();
        return t.value;
      }
      if (t.type === "str") {
        next();
        return t.value;
      }
      if (t.type === "cell") {
        next();
        return resolveCell(t.value, getRaw, stack);
      }
      if (t.type === "range") {
        // bare range only valid inside functions
        next();
        return expandRange(t.value, getRaw, stack);
      }
      if (t.type === "func") {
        next();
        if (!peek() || peek().type !== "op" || peek().value !== "(") {
          // bare name treated as 0-arg? invalid
          throw new Error("#NAME?");
        }
        next(); // (
        const args = [];
        if (peek() && !(peek().type === "op" && peek().value === ")")) {
          args.push(parseArg());
          while (peek() && peek().type === "op" && peek().value === ",") {
            next();
            args.push(parseArg());
          }
        }
        if (!peek() || peek().type !== "op" || peek().value !== ")") throw new Error("#ERROR!");
        next();
        return callFunc(t.value, args);
      }
      if (t.type === "op" && t.value === "(") {
        next();
        const v = parseComparison();
        if (!peek() || peek().value !== ")") throw new Error("#ERROR!");
        next();
        return v;
      }
      throw new Error("#ERROR!");
    }

    function parseArg() {
      // argument may be a range token directly
      if (peek() && peek().type === "range") {
        const r = next();
        return expandRange(r.value, getRaw, stack);
      }
      // cell:cell written as cell op : cell — tokenizer merges ranges when contiguous
      return parseComparison();
    }

    return parseComparison();
  }

  function resolveCell(ref, getRaw, stack) {
    const m = String(ref).replace(/\$/g, "").match(CELL_REF);
    if (!m) throw new Error("#REF!");
    const c = colToIndex(m[1]);
    const r = parseInt(m[2], 10) - 1;
    if (r < 0 || c < 0) throw new Error("#REF!");
    const key = r + "," + c;
    if (stack.has(key)) throw new Error("#CYCLE!");
    stack.add(key);
    const raw = getRaw(r, c);
    const ev = evaluate(raw, getRaw, stack);
    stack.delete(key);
    if (ev.kind === "error") throw new Error(ev.display);
    if (ev.kind === "empty") return 0;
    if (ev.kind === "number" || ev.kind === "formula") return ev.value;
    const n = parseNumber(ev.value);
    return n !== null ? n : ev.value;
  }

  function expandRange(ref, getRaw, stack) {
    const m = String(ref).replace(/\$/g, "").match(RANGE_REF);
    if (!m) throw new Error("#REF!");
    let c1 = colToIndex(m[1]);
    let r1 = parseInt(m[2], 10) - 1;
    let c2 = colToIndex(m[3]);
    let r2 = parseInt(m[4], 10) - 1;
    if (r1 > r2) [r1, r2] = [r2, r1];
    if (c1 > c2) [c1, c2] = [c2, c1];
    const vals = [];
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const key = r + "," + c;
        if (stack.has(key)) throw new Error("#CYCLE!");
        stack.add(key);
        const raw = getRaw(r, c);
        const ev = evaluate(raw, getRaw, stack);
        stack.delete(key);
        if (ev.kind === "error") continue;
        if (ev.kind === "empty") continue;
        if (typeof ev.value === "number") vals.push(ev.value);
        else {
          const n = parseNumber(ev.value);
          if (n !== null) vals.push(n);
          else vals.push(ev.value);
        }
      }
    }
    return vals;
  }

  function flatten(args) {
    const out = [];
    for (const a of args) {
      if (Array.isArray(a)) out.push(...a);
      else out.push(a);
    }
    return out;
  }

  function nums(args) {
    return flatten(args)
      .map((v) => (typeof v === "number" ? v : parseNumber(v)))
      .filter((v) => v !== null && typeof v === "number");
  }

  function callFunc(name, args) {
    const n = name.toUpperCase();
    switch (n) {
      case "SUM":
        return nums(args).reduce((a, b) => a + b, 0);
      case "AVERAGE":
      case "AVG": {
        const a = nums(args);
        if (!a.length) throw new Error("#DIV/0!");
        return a.reduce((x, y) => x + y, 0) / a.length;
      }
      case "COUNT":
        return nums(args).length;
      case "COUNTA":
        return flatten(args).filter((v) => v !== "" && v != null).length;
      case "MIN": {
        const a = nums(args);
        if (!a.length) throw new Error("#VALUE!");
        return Math.min(...a);
      }
      case "MAX": {
        const a = nums(args);
        if (!a.length) throw new Error("#VALUE!");
        return Math.max(...a);
      }
      case "ROUND": {
        const a = nums(args);
        const digits = a[1] != null ? a[1] : 0;
        const f = Math.pow(10, digits);
        return Math.round(a[0] * f) / f;
      }
      case "ABS":
        return Math.abs(nums(args)[0] || 0);
      case "INT":
        return Math.floor(nums(args)[0] || 0);
      case "SQRT": {
        const v = nums(args)[0];
        if (v < 0) throw new Error("#NUM!");
        return Math.sqrt(v);
      }
      case "POWER":
      case "POW": {
        const a = nums(args);
        return Math.pow(a[0], a[1]);
      }
      case "IF": {
        const cond = Array.isArray(args[0]) ? args[0][0] : args[0];
        const truthy = cond === true || (typeof cond === "number" && cond !== 0) || (typeof cond === "string" && cond !== "" && cond !== "FALSE");
        return truthy ? (args[1] !== undefined ? args[1] : true) : args[2] !== undefined ? args[2] : false;
      }
      case "LEN": {
        const v = flatten(args)[0];
        return String(v ?? "").length;
      }
      case "UPPER":
        return String(flatten(args)[0] ?? "").toUpperCase();
      case "LOWER":
        return String(flatten(args)[0] ?? "").toLowerCase();
      case "TRIM":
        return String(flatten(args)[0] ?? "").trim();
      case "CONCAT":
      case "CONCATENATE":
        return flatten(args).map((v) => String(v ?? "")).join("");
      case "TRUE":
        return true;
      case "FALSE":
        return false;
      case "PI":
        return Math.PI;
      case "NOW":
        return new Date().toLocaleString();
      case "TODAY":
        return new Date().toLocaleDateString();
      default:
        throw new Error("#NAME?");
    }
  }

  /** Display value for a raw cell (formula → computed) */
  function displayValue(raw, getRaw) {
    return evaluate(raw, getRaw).display;
  }

  global.Formulas = {
    evaluate,
    displayValue,
    parseNumber,
    colToIndex,
  };
})(typeof window !== "undefined" ? window : globalThis);
