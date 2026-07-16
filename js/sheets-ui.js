/**
 * Google Sheets-style UI layer: menus, function panel, autocomplete
 */
(function () {
  "use strict";

  const A = () => window.SheetsApp;

  // ── Menus ───────────────────────────────────────────────
  function closeMenus() {
    document.querySelectorAll(".menu-dd.open").forEach((el) => el.classList.remove("open"));
    document.querySelectorAll(".menu-btn.open").forEach((el) => el.classList.remove("open"));
  }

  document.querySelectorAll(".menu-btn[data-menu]").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const id = btn.dataset.menu;
      const dd = document.getElementById(id);
      const was = dd.classList.contains("open");
      closeMenus();
      if (!was) {
        dd.classList.add("open");
        btn.classList.add("open");
      }
    });
  });
  document.addEventListener("click", () => closeMenus());

  function runAct(act) {
    const app = A();
    if (!app) return;
    closeMenus();
    const map = {
      new: () => app.newWorkbook(),
      open: () => document.getElementById("fileInput").click(),
      save: () => app.saveAs("auto"),
      "save-xlsx": () => app.saveAs("xlsx"),
      "save-csv": () => app.saveAs("csv"),
      "save-ods": () => app.saveAs("ods"),
      sample: () => app.loadSample(),
      undo: () => app.undo(),
      redo: () => app.redo(),
      cut: () => app.copySelection(true),
      copy: () => app.copySelection(false),
      paste: () => app.pasteSelection(),
      find: () => app.openFind(true),
      selectAll: () => app.selectAll(),
      freeze: () => {
        app.freezeHeader = !app.freezeHeader;
        app.toast(app.freezeHeader ? "Header frozen" : "Unfrozen");
      },
      gridlines: () => {
        app.showGrid = !app.showGrid;
      },
      showFx: () => {
        app.showFormulas = !app.showFormulas;
        app.toast(app.showFormulas ? "Showing formulas" : "Showing values");
      },
      zoomIn: () => app.setZoom((window._zoom || 1) + 0.1),
      zoomOut: () => app.setZoom((window._zoom || 1) - 0.1),
      zoom100: () => app.setZoom(1),
      theme: () => app.toggleTheme(),
      addRow: () => app.addRow(),
      addCol: () => app.addCol(),
      addSheet: () => app.addWorkbookSheet(),
      fxPanel: () => openFxPanel(true),
      calc: () => app.openCalculator(),
      bold: () => app.toggleStyle("bold"),
      italic: () => app.toggleStyle("italic"),
      underline: () => app.toggleStyle("underline"),
      alignLeft: () => app.applyStyleToSelection({ align: "left" }),
      alignCenter: () => app.applyStyleToSelection({ align: "center" }),
      alignRight: () => app.applyStyleToSelection({ align: "right" }),
      numCurrency: () => app.applyStyleToSelection({ numFmt: "currency" }),
      numPercent: () => app.applyStyleToSelection({ numFmt: "percent" }),
      clearFmt: () => app.applyStyleToSelection(app.emptyStyle()),
      sortAsc: () => app.sortByActiveCol(true),
      sortDesc: () => app.sortByActiveCol(false),
      "tool-sum": () => app.applyCalcTool("sum"),
      "tool-avg": () => app.applyCalcTool("avg"),
      "tool-pct": () => app.applyCalcTool("pct"),
      "tool-running": () => app.applyCalcTool("running"),
      "tool-vatAdd": () => app.applyCalcTool("vatAdd"),
      "tool-loan": () => app.applyCalcTool("loan"),
      "tool-evaluate": () => app.applyCalcTool("evaluate"),
      "tool-interest": () => app.applyCalcTool("interest"),
      "tool-compound": () => app.applyCalcTool("compound"),
      "tool-discount": () => app.applyCalcTool("discount"),
      help: () => document.getElementById("helpModal").classList.add("open"),
      insertRow: () => app.addRow(),
      insertCol: () => app.addCol(),
      deleteRow: () => app.delRow(),
      deleteCol: () => app.delCol(),
      clear: () => app.deleteSelection(false),
      autoSum: () => app.autoSum(),
    };
    if (map[act]) map[act]();
  }

  document.querySelectorAll(".menu-dd button[data-act]").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      runAct(btn.dataset.act);
    });
  });

  // Context menu extra acts
  document.getElementById("ctxMenu")?.querySelectorAll("button[data-act]").forEach((btn) => {
    // app.js already binds some; add fxPanel
    if (btn.dataset.act === "fxPanel") {
      btn.addEventListener("click", () => openFxPanel(true));
    }
  });

  // Toolbar extras
  document.getElementById("btnUndo")?.addEventListener("click", () => A()?.undo());
  document.getElementById("btnRedo")?.addEventListener("click", () => A()?.redo());
  document.getElementById("btnShareFake")?.addEventListener("click", () => A()?.saveAs("xlsx"));
  document.getElementById("btnOpenCalcTop")?.addEventListener("click", () => A()?.openCalculator());
  document.getElementById("btnLogo")?.addEventListener("click", () => A()?.openBackstage("new"));
  document.getElementById("fxLabelBtn")?.addEventListener("click", () => openFxPanel(true));
  document.getElementById("btnStar")?.addEventListener("click", function () {
    this.textContent = this.textContent === "☆" ? "★" : "☆";
    this.style.color = this.textContent === "★" ? "#f4b400" : "";
  });

  document.getElementById("zoomSelect")?.addEventListener("change", (e) => {
    A()?.setZoom(Number(e.target.value));
  });

  // Rename file
  const titleInput = document.getElementById("fileName");
  titleInput?.addEventListener("change", () => {
    if (A()) A().fileName = titleInput.value.trim() || "Untitled spreadsheet";
  });
  titleInput?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      titleInput.blur();
    }
  });

  // ── Function panel ──────────────────────────────────────
  const fxPanel = document.getElementById("fxPanel");
  const fxList = document.getElementById("fxList");
  const fxDetail = document.getElementById("fxDetail");
  const fxSearch = document.getElementById("fxSearch");
  const fxCats = document.getElementById("fxCats");
  const fxSuggest = document.getElementById("fxSuggest");
  const formulaInput = document.getElementById("formulaInput");
  const railFx = document.getElementById("railFx");
  const btnFxToggle = document.getElementById("btnFxToggle");

  let fxCat = "All";
  let fxSelIdx = 0;

  function openFxPanel(open) {
    if (!fxPanel) return;
    const on = open !== undefined ? open : !fxPanel.classList.contains("open");
    fxPanel.classList.toggle("open", on);
    railFx?.classList.toggle("on", on);
    btnFxToggle?.classList.toggle("on", on);
    if (on) renderFxList();
  }

  function renderCats() {
    if (!fxCats || !window.FxCatalog) return;
    fxCats.innerHTML = "";
    FxCatalog.CATS.forEach((c) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "fx-cat" + (c === fxCat ? " on" : "");
      b.textContent = c === "All" ? "All" : c;
      b.addEventListener("click", () => {
        fxCat = c;
        renderCats();
        renderFxList();
      });
      fxCats.appendChild(b);
    });
  }

  function showDetail(fn) {
    if (!fxDetail || !fn) return;
    fxDetail.innerHTML =
      `<strong>${fn.name}</strong>` +
      `<div>${escapeHtml(fn.sig)}</div>` +
      `<div style="margin-top:6px">${escapeHtml(fn.desc)}</div>` +
      `<div class="ex">${escapeHtml(fn.ex)}</div>`;
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function insertFunction(fn) {
    const app = A();
    if (!app || !fn) return;
    app.commitEdit();
    let formula = fn.ex;
    // Prefer range-aware insert for aggregate fns
    if (["SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX", "PRODUCT", "MEDIAN", "STDEV", "VAR"].includes(fn.name)) {
      app.insertFx(fn.name);
      return;
    }
    app.pushUndo();
    const pos = app.getActive();
    app.setCell(pos.r, pos.c, formula);
    app.refreshRange(pos.r, pos.c, pos.r, pos.c);
    app.markDirty();
    app.updateFormulaBar();
    app.toast(fn.name);
    hideSuggest();
  }

  function renderFxList() {
    if (!fxList || !window.FxCatalog) return;
    const q = fxSearch?.value || "";
    const items = FxCatalog.search(q, fxCat);
    fxList.innerHTML = "";
    if (!items.length) {
      fxList.innerHTML = `<div style="padding:16px;color:var(--muted);font-size:12px">No functions found</div>`;
      return;
    }
    items.forEach((fn, i) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "fx-item" + (i === 0 ? " active" : "");
      b.innerHTML =
        `<div class="name">${escapeHtml(fn.name)}</div>` +
        `<div class="sig">${escapeHtml(fn.sig)}</div>` +
        `<div class="desc">${escapeHtml(fn.desc)}</div>`;
      b.addEventListener("mouseenter", () => showDetail(fn));
      b.addEventListener("click", () => insertFunction(fn));
      fxList.appendChild(b);
    });
    showDetail(items[0]);
  }

  fxSearch?.addEventListener("input", () => renderFxList());
  document.getElementById("fxPanelClose")?.addEventListener("click", () => openFxPanel(false));
  document.getElementById("btnFxToggle")?.addEventListener("click", () => openFxPanel());
  railFx?.addEventListener("click", () => openFxPanel());
  document.getElementById("railCalc")?.addEventListener("click", () => A()?.openCalculator());

  // ── Formula autocomplete ────────────────────────────────
  function hideSuggest() {
    fxSuggest?.classList.remove("open");
  }

  function positionSuggest() {
    if (!fxSuggest || !formulaInput) return;
    const r = formulaInput.getBoundingClientRect();
    fxSuggest.style.left = r.left + "px";
    fxSuggest.style.top = r.bottom + 2 + "px";
    fxSuggest.style.width = Math.max(280, r.width * 0.5) + "px";
  }

  function getFormulaQuery(text) {
    if (!text || text[0] !== "=") return null;
    // last function token being typed
    const m = text.match(/=([A-Za-z][A-Za-z0-9]*)$/i);
    if (m) return m[1];
    const m2 = text.match(/([A-Za-z][A-Za-z0-9]*)$/);
    // only if looks like starting a function after =( or ,
    if (m2 && /[=(,]$/.test(text.slice(0, -m2[1].length).trimEnd().slice(-1) || text[0])) {
      return m2[1];
    }
    if (text === "=") return "";
    if (/^=[A-Za-z0-9]*$/i.test(text)) return text.slice(1);
    return null;
  }

  function showSuggest(query) {
    if (!fxSuggest || !window.FxCatalog) return;
    const items = FxCatalog.search(query || "", "All").slice(0, 12);
    if (!items.length) {
      hideSuggest();
      return;
    }
    fxSelIdx = 0;
    fxSuggest.innerHTML = "";
    items.forEach((fn, i) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "fx-item" + (i === 0 ? " active" : "");
      b.innerHTML =
        `<div class="name">${escapeHtml(fn.name)}</div>` +
        `<div class="sig">${escapeHtml(fn.sig)}</div>`;
      b.addEventListener("mouseenter", () => {
        fxSelIdx = i;
        fxSuggest.querySelectorAll(".fx-item").forEach((el, j) => el.classList.toggle("active", j === i));
        showDetail(fn);
      });
      b.addEventListener("mousedown", (e) => {
        e.preventDefault();
        applySuggest(fn);
      });
      fxSuggest.appendChild(b);
    });
    positionSuggest();
    fxSuggest.classList.add("open");
    // mirror to side panel
    if (fxPanel?.classList.contains("open")) {
      if (fxSearch) fxSearch.value = query || "";
      renderFxList();
      showDetail(items[0]);
    }
  }

  function applySuggest(fn) {
    if (!formulaInput || !fn) return;
    const text = formulaInput.value || "";
    let next;
    if (/^=[A-Za-z0-9]*$/i.test(text) || text === "=") {
      next = "=" + fn.name + "(";
    } else {
      // replace trailing name
      next = text.replace(/[A-Za-z][A-Za-z0-9]*$/, fn.name + "(");
      if (!next.startsWith("=")) next = "=" + fn.name + "(";
    }
    formulaInput.value = next;
    formulaInput.focus();
    formulaInput.setSelectionRange(next.length, next.length);
    hideSuggest();
    showDetail(fn);
  }

  formulaInput?.addEventListener("input", () => {
    const q = getFormulaQuery(formulaInput.value);
    if (q !== null) showSuggest(q);
    else hideSuggest();
  });

  formulaInput?.addEventListener("keydown", (e) => {
    if (!fxSuggest?.classList.contains("open")) return;
    const items = [...fxSuggest.querySelectorAll(".fx-item")];
    if (!items.length) return;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      fxSelIdx = Math.min(items.length - 1, fxSelIdx + 1);
      items.forEach((el, j) => el.classList.toggle("active", j === fxSelIdx));
      items[fxSelIdx].scrollIntoView({ block: "nearest" });
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      fxSelIdx = Math.max(0, fxSelIdx - 1);
      items.forEach((el, j) => el.classList.toggle("active", j === fxSelIdx));
    } else if (e.key === "Tab" || (e.key === "Enter" && e.ctrlKey)) {
      e.preventDefault();
      const name = items[fxSelIdx]?.querySelector(".name")?.textContent;
      const fn = name && FxCatalog.get(name);
      if (fn) applySuggest(fn);
    } else if (e.key === "Escape") {
      hideSuggest();
    }
  });

  formulaInput?.addEventListener("blur", () => {
    setTimeout(hideSuggest, 150);
  });

  window.addEventListener("resize", () => {
    if (fxSuggest?.classList.contains("open")) positionSuggest();
  });

  // Patch setZoom tracking
  const origSetZoom = () => {};
  document.getElementById("zoomIn")?.addEventListener("click", () => {
    const z = parseFloat(document.getElementById("zoomLabel")?.textContent) / 100 || 1;
    A()?.setZoom(z + 0.1);
    const zs = document.getElementById("zoomSelect");
    if (zs) zs.value = String(Math.round((z + 0.1) * 100) / 100);
  });

  // Init panel
  renderCats();
  renderFxList();
  // default open like Sheets sidebar tools
  openFxPanel(true);

  // Click outside calc already handled

  // Expose
  window.SheetsUI = { openFxPanel, runAct, insertFunction };
})();
