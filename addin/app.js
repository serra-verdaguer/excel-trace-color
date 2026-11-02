/*
 * Task pane UI + ribbon/keyboard command wiring (shared runtime).
 *
 * The trace view is a tree rooted at YOUR CELL, which stays pinned at the top.
 * Precedents (or dependents) open underneath it like folders, so you always
 * see where you are relative to where you started. Exactly one row is
 * highlighted, and it is always the cell Excel has selected.
 *
 *   ↑ ↓  move (Excel follows)     →  open / step into     ←  one step back
 *   Home your cell                Esc back to your cell and close the pane
 *
 * Built for big models: the list only draws the rows on screen, and cell
 * values/formulas are fetched lazily for the rows you scroll to.
 */
(function () {
  "use strict";
  const C = window.TCCore;
  const COLOR_KEY = "tc.colors";
  const ROW_H = 28; // keep in sync with --row-h
  const PAD_ROWS = 8; // rows drawn above/below the viewport

  const S = {
    dir: "precedents",
    root: null, // node for YOUR CELL
    rows: [], // visible nodes in order (flattened tree)
    cur: -1, // index into rows; -1 = YOUR CELL
    paused: false, // after Esc: pane closed, selection no longer followed
    expectSel: null, // selection we caused ourselves
    seq: 0,
    navTimer: null,
    autoTimer: null,
    fetchTimer: null,
  };
  const $ = (id) => document.getElementById(id);

  // ------------------------------------------------------------------ status

  function status(text, isErr) {
    const el = $("status");
    el.textContent = text || "";
    el.classList.toggle("err", !!isErr);
  }

  function busy(on) {
    for (const id of ["colorBtn", "colorWbBtn"]) $(id).disabled = on;
  }

  function errText(err) {
    return (err && (err.message || err.code)) || String(err);
  }

  // ------------------------------------------------------------------ colors

  function loadColors() {
    try {
      const saved = JSON.parse(localStorage.getItem(COLOR_KEY) || "{}");
      const out = Object.assign({}, C.DEFAULT_COLORS);
      for (const k of Object.keys(out)) if (C.isHexColor(saved[k])) out[k] = saved[k];
      if (saved.inputFill === "") out.inputFill = "";
      return out;
    } catch (e) {
      return Object.assign({}, C.DEFAULT_COLORS);
    }
  }

  function saveColors(colors) {
    try { localStorage.setItem(COLOR_KEY, JSON.stringify(colors)); } catch (e) { /* storage blocked */ }
  }

  function syncColorInputs() {
    const colors = loadColors();
    document.querySelectorAll("input[data-color]").forEach((inp) => { inp.value = colors[inp.dataset.color] || C.DEFAULT_COLORS[inp.dataset.color]; });
    $("fillInputs").checked = !!colors.inputFill;
  }

  async function runColor(scope) {
    busy(true);
    status("Color coding…");
    try {
      const s = await window.TCColor.colorCode(scope, loadColors(), (t) => status(t));
      status(
        `Colored ${s.label} in ${(s.ms / 1000).toFixed(1)}s — ` +
          `${C.formatCount(s.input)} inputs, ${C.formatCount(s.formula)} formulas, ` +
          `${C.formatCount(s.link)} sheet links, ${C.formatCount(s.external)} external links.`
      );
    } catch (err) {
      status("Color coding failed: " + errText(err), true);
    } finally {
      busy(false);
    }
  }

  // ------------------------------------------------------------- tree model

  const sameAddr = (a, b) => !!a && !!b && a.replace(/'/g, "").toLowerCase() === b.replace(/'/g, "").toLowerCase();

  function makeNode(area, sheet, parent) {
    const node = {
      address: area.address,
      sheet,
      local: area.local,
      cells: area.cells,
      value: area.value,
      formula: area.value !== undefined ? area.formula || null : undefined, // undefined = not fetched yet
      parent,
      depth: parent ? parent.depth + 1 : -1,
      children: null,
      expanded: false,
      loading: false,
      loop: false,
    };
    for (let p = parent; p; p = p.parent) if (sameAddr(p.address, node.address)) node.loop = true;
    return node;
  }

  function childrenFrom(res, parent) {
    let list = [];
    for (const g of res.groups) for (const a of g.areas) list.push(Object.assign({ sheet: g.sheet }, a));
    if (S.dir === "precedents" && res.originFormula) list = C.orderByFormula(list, res.originFormula, parent.sheet);
    return list.map((a) => makeNode(a, a.sheet, parent));
  }

  // "8 cells", or "whole columns" / "whole rows" for A:B / 3:5 style references.
  function sizeLabel(n) {
    if (/^\$?[A-Z]{1,3}:\$?[A-Z]{1,3}$/i.test(n.local)) return n.local.split(":")[0] === n.local.split(":")[1] ? "whole column" : "whole columns";
    if (/^\$?\d+:\$?\d+$/.test(n.local)) return n.local.split(":")[0] === n.local.split(":")[1] ? "whole row" : "whole rows";
    return `${C.formatCount(n.cells)} cells`;
  }
  const isWhole = (n) => /^\$?([A-Z]{1,3}|\d+):\$?([A-Z]{1,3}|\d+)$/i.test(n.local);

  // A node can be opened if it is a formula or a range (or not known yet).
  function canOpen(n) {
    if (n.loop) return false;
    if (n.cells > 1) return true;
    return n.formula !== null || S.dir === "dependents";
  }

  function flatten() {
    const out = [];
    const walk = (n) => {
      for (const c of n.children || []) {
        out.push(c);
        if (c.expanded) walk(c);
      }
    };
    if (S.root) walk(S.root);
    S.rows = out;
  }

  function traceKind(all) {
    if (S.dir === "precedents") return all ? "precedents" : "directPrecedents";
    return all ? "dependents" : "directDependents";
  }

  // Starts a fresh tree at `address` (or the selection).
  // Starts a fresh tree. `all` opens it level by level (each level indented
  // under the cell it feeds) instead of Excel's flat all-levels list, which
  // glues different levels together into one block.
  async function buildRoot(address, all) {
    await buildDirect(address);
    if (all && S.root) await openAllLevels();
  }

  async function openAllLevels() {
    const MAX_OPEN = 60; // Excel round trips; the rest stay openable with →
    const MAX_DEPTH = 6;
    let opened = 0;
    let queue = S.root.children.slice();
    status("Opening all levels…");
    while (queue.length && opened < MAX_OPEN) {
      const next = [];
      for (const n of queue) {
        if (opened >= MAX_OPEN) break;
        if (n.depth >= MAX_DEPTH || isWhole(n) || n.formula === null && n.cells === 1 || n.loop) continue;
        await expand(n, true);
        opened++;
        if (n.children) next.push(...n.children);
      }
      queue = next;
    }
    flatten();
    renderRows();
    renderDetail();
    const total = S.rows.length;
    status(`${C.formatCount(total)} ${S.dir} shown, level by level${queue.length ? " (→ opens deeper levels)" : ""}.`);
  }

  async function buildDirect(address) {
    const all = false;
    const seq = ++S.seq;
    status("Tracing…");
    const t0 = Date.now();
    try {
      const res = await window.TCTrace.trace(traceKind(all), address);
      if (seq !== S.seq) return;
      const { sheet, local } = C.splitAddress(res.origin);
      const root = makeNode({ address: res.origin, local, cells: res.originCells, value: res.originValue, formula: res.originFormula }, sheet, null);
      root.children = childrenFrom(res, root);
      root.expanded = true;
      S.root = root;
      S.cur = -1;
      flatten();
      $("tree").scrollTop = 0;
      renderHome();
      renderRows();
      renderDetail();
      const noun = S.dir === "precedents" ? "precedent" : "dependent";
      const n = root.children.length;
      status(`${C.formatCount(n)} ${all ? "" : "direct "}${noun}${n === 1 ? "" : "s"} · ${((Date.now() - t0) / 1000).toFixed(2)}s`);
    } catch (err) {
      if (seq !== S.seq) return;
      status(errText(err), true);
    }
  }

  async function expand(node, quiet) {
    if (!canOpen(node) || node.expanded) return;
    if (!node.children) {
      node.loading = true;
      renderRows();
      try {
        const res = await window.TCTrace.trace(traceKind(false), node.address);
        node.value = res.originCells > 1 ? node.value : res.originValue;
        node.formula = res.originCells > 1 ? node.formula : res.originFormula;
        node.children = childrenFrom(res, node);
      } catch (err) {
        node.children = [];
        status(errText(err), true);
      }
      node.loading = false;
    }
    if (!node.children.length) {
      if (!quiet) status(`${node.local} has no ${S.dir}.`);
      renderRows();
      renderDetail();
      return;
    }
    node.expanded = true;
    const keep = S.rows[S.cur];
    flatten();
    S.cur = keep ? S.rows.indexOf(keep) : S.cur;
    renderRows();
    renderDetail();
  }

  function collapse(node) {
    node.expanded = false;
    const keep = S.rows[S.cur];
    flatten();
    S.cur = keep ? S.rows.indexOf(keep) : -1;
    renderRows();
  }

  // ---------------------------------------------------------------- cursor

  function curNode() {
    return S.cur === -1 ? S.root : S.rows[S.cur];
  }

  // Moves the highlight; with `go`, Excel selects that cell too.
  function setCursor(i, go) {
    if (!S.root) return;
    S.cur = Math.max(-1, Math.min(S.rows.length - 1, i));
    scrollToCursor();
    renderHome();
    renderRows();
    renderDetail();
    if (go) {
      const addr = curNode().address;
      clearTimeout(S.navTimer);
      S.navTimer = setTimeout(() => goTo(addr), 50); // holding ↓ doesn't queue jumps
    }
  }

  function scrollToCursor() {
    const tree = $("tree");
    if (S.cur < 0) {
      tree.scrollTop = 0;
      return;
    }
    const top = S.cur * ROW_H;
    if (top < tree.scrollTop) tree.scrollTop = top;
    else if (top + ROW_H > tree.scrollTop + tree.clientHeight) tree.scrollTop = top + ROW_H - tree.clientHeight;
  }

  async function goTo(address) {
    S.expectSel = address;
    try {
      await window.TCTrace.goTo(address);
    } catch (err) {
      status("Could not go there: " + errText(err), true);
    }
    if (!S.paused) focusPane();
  }

  // The list keeps the keyboard so plain arrows work; retried because Excel
  // may take focus back right after selecting a cell.
  function focusPane() {
    const f = () => {
      try { window.focus(); } catch (e) { /* ignore */ }
      $("tree").focus({ preventScroll: true });
    };
    f();
    setTimeout(f, 80);
    setTimeout(f, 250);
  }

  // ------------------------------------------------------------- keyboard

  const keys = {
    down: () => setCursor(S.cur + 1, true),
    up: () => setCursor(S.cur - 1, true),
    home: () => setCursor(-1, true),
    end: () => setCursor(S.rows.length - 1, true),
    pageDown: () => setCursor(S.cur + 10, true),
    pageUp: () => setCursor(S.cur - 10, true),
    right: async () => {
      if (S.cur === -1) return setCursor(0, true);
      const n = curNode();
      if (!n) return;
      if (!n.expanded) await expand(n);
      else setCursor(S.cur + 1, true); // already open: step to its first child
    },
    // Every ← is one step closer to your cell (closing what you leave behind).
    left: () => {
      if (S.cur === -1) return;
      const n = curNode();
      const parent = n.parent;
      if (n.expanded) collapse(n);
      setCursor(parent === S.root ? -1 : S.rows.indexOf(parent), true);
    },
    exit: () => exitTrace(),
  };

  function onKey(e) {
    const t = e.target;
    if (t && t.tagName === "INPUT" && t.type !== "checkbox") return;
    const map = {
      ArrowDown: "down", ArrowUp: "up", ArrowRight: "right", Enter: "right", ArrowLeft: "left",
      Home: "home", End: "end", PageDown: "pageDown", PageUp: "pageUp", Escape: "exit",
    };
    const k = map[e.key];
    if (!k || !S.root) return;
    e.preventDefault();
    keys[k]();
  }

  async function exitTrace() {
    if (S.root) await goTo(S.root.address);
    S.paused = true;
    S.cur = -1;
    renderHome();
    renderRows();
    try { await Office.addin.hide(); } catch (e) { /* not supported */ }
    status("Stopped. Press Precedents or Dependents to trace again.");
  }

  // ------------------------------------------------------------- rendering

  function formulaEl(formula, sheet) {
    const el = document.createElement("code");
    el.className = "fx";
    for (const t of C.tokenizeFormula(formula, sheet)) {
      const span = document.createElement("span");
      span.textContent = t.text;
      if (t.kind !== "op") span.className = "tk-" + t.kind;
      if (t.kind === "num" && !/^[01](\.0*)?%?$/.test(t.text)) span.classList.add("hard");
      el.appendChild(span);
    }
    return el;
  }

  function renderHome() {
    const r = S.root;
    $("home").classList.toggle("cur", !!r && S.cur === -1);
    if (!r) return;
    $("homeAddr").textContent = `${r.sheet} › ${r.local}`;
    $("homeVal").textContent = r.cells > 1 ? `${C.formatCount(r.cells)} cells` : r.value || "";
    const fx = $("homeFx");
    fx.innerHTML = "";
    if (r.formula) fx.appendChild(formulaEl(r.formula, r.sheet));
  }

  function renderRows() {
    const tree = $("tree");
    const rowsEl = $("rows");
    $("empty").hidden = !!S.root;
    $("spacer").style.height = S.rows.length * ROW_H + "px";
    rowsEl.innerHTML = "";
    if (!S.root) return;
    if (!S.rows.length) {
      const noun = S.dir === "precedents" ? "precedents" : "dependents";
      const why = S.dir === "precedents" ? "It is a typed-in value or has no cell references." : "No formula uses this cell.";
      rowsEl.innerHTML = `<div class="empty"><b>No ${noun}.</b><br>${why}</div>`;
      return;
    }
    const first = Math.max(0, Math.floor(tree.scrollTop / ROW_H) - PAD_ROWS);
    const last = Math.min(S.rows.length, first + Math.ceil(tree.clientHeight / ROW_H) + 2 * PAD_ROWS);
    const missing = [];
    for (let i = first; i < last; i++) {
      const n = S.rows[i];
      rowsEl.appendChild(rowEl(n, i));
      if (n.cells === 1 && n.value === undefined && !n.fetching) missing.push(n);
    }
    if (missing.length) queueFetch(missing);
  }

  function rowEl(n, i) {
    const row = document.createElement("div");
    row.className = "row" + (i === S.cur ? " cur" : "");
    row.style.top = i * ROW_H + "px";
    row.style.paddingLeft = 6 + n.depth * 16 + "px";
    row.title = n.address;

    const tw = document.createElement("span");
    tw.className = "tw";
    tw.textContent = n.loading ? "…" : n.loop ? "↻" : n.expanded ? "▾" : canOpen(n) ? "▸" : "·";
    tw.addEventListener("click", (e) => {
      e.stopPropagation();
      setCursor(i, true);
      if (n.expanded) collapse(n);
      else expand(n);
    });
    row.appendChild(tw);

    // Sheet pill whenever the reference sits on a different sheet than its parent.
    if (n.sheet !== n.parent.sheet) {
      const pill = document.createElement("span");
      pill.className = "pill";
      pill.textContent = n.sheet;
      row.appendChild(pill);
    }
    const addr = document.createElement("span");
    addr.className = "addr";
    addr.textContent = n.local;
    row.appendChild(addr);

    const icons = document.createElement("span");
    icons.className = "icons";
    if (n.formula) {
      const notes = C.describeFormula(n.formula, n.sheet);
      icons.innerHTML = '<span class="ic-fx" title="Formula">ƒ</span>';
      if (notes.hardcoded.length) {
        const h = document.createElement("span");
        h.className = "ic-hard";
        h.title = "Typed-in number inside the formula";
        h.textContent = notes.hardcoded[0];
        icons.appendChild(h);
      }
      if (notes.external) icons.insertAdjacentHTML("beforeend", '<span class="ic-ext" title="Links to another file">⧉</span>');
    }
    if (n.loop) icons.insertAdjacentHTML("beforeend", '<span class="ic-loop" title="Circular reference">circular</span>');
    row.appendChild(icons);

    const val = document.createElement("span");
    val.className = "val";
    val.textContent = n.cells > 1 ? sizeLabel(n) : n.value || "";
    row.appendChild(val);

    row.addEventListener("click", () => setCursor(i, true));
    row.addEventListener("dblclick", () => (n.expanded ? collapse(n) : expand(n)));
    return row;
  }

  // Values for rows scrolled into view, fetched in one batch.
  function queueFetch(nodes) {
    for (const n of nodes) n.fetching = true;
    clearTimeout(S.fetchTimer);
    S.fetchTimer = setTimeout(async () => {
      const batch = S.rows.filter((n) => n.fetching && n.value === undefined).slice(0, 300);
      if (!batch.length) return;
      try {
        const out = await window.TCTrace.fetchCells(batch.map((n) => n.address));
        out.forEach((d, k) => { batch[k].value = d.value; batch[k].formula = d.formula; });
      } catch (e) {
        batch.forEach((n) => { n.value = ""; n.formula = null; });
      }
      batch.forEach((n) => { n.fetching = false; });
      renderRows();
      renderDetail();
    }, 60);
  }

  function renderDetail() {
    const n = curNode();
    const box = $("detail");
    box.hidden = !n;
    if (!n) return;

    // Path from your cell to here: "D10 ← K47 ← Page 2!D5"
    const chain = [];
    for (let p = n; p; p = p.parent) chain.unshift(p);
    const path = $("path");
    path.innerHTML = "";
    const span = document.createElement("span");
    chain.forEach((p, k) => {
      const label = (k === 0 || p.sheet !== chain[k - 1].sheet ? p.sheet + "!" : "") + p.local;
      const el = document.createElement(k === chain.length - 1 ? "b" : "span");
      el.textContent = label;
      span.appendChild(el);
      if (k < chain.length - 1) span.appendChild(document.createTextNode(S.dir === "precedents" ? "  ←  " : "  →  "));
    });
    path.appendChild(span);

    $("dAddr").textContent = `${n.sheet} › ${n.local}`;
    $("dVal").textContent = n.cells > 1 ? sizeLabel(n) : n.value || "";
    const fx = $("dFx");
    fx.innerHTML = "";
    if (n.formula) fx.appendChild(formulaEl(n.formula, n.sheet));

    const notes = $("dNotes");
    notes.innerHTML = "";
    const wrap = document.createElement("div");
    wrap.className = "notes";
    const note = (cls, text) => {
      const d = document.createElement("div");
      d.className = "note " + cls;
      d.textContent = text;
      wrap.appendChild(d);
    };
    if (n === S.root) note("n-info", "This is your cell. ↓ to walk through its " + S.dir + ".");
    if (n.parent && n.sheet !== n.parent.sheet) note("n-sheet", `Different sheet: ${n.sheet}`);
    if (n.loop) note("n-ext", "Circular reference: this cell is already above in the path.");
    if (n.formula) {
      const d = C.describeFormula(n.formula, n.sheet);
      for (const sh of d.sheets) note("n-sheet", `Pulls data from sheet ${sh}`);
      if (d.external) note("n-ext", "Pulls data from another file");
      if (d.hardcoded.length) note("n-hard", `Typed-in number in the formula: ${d.hardcoded.slice(0, 4).join(", ")}`);
      if (n !== S.root && !n.expanded) note("n-info", `ƒ Formula · → to open its ${S.dir}`);
    } else if (n.cells > 1) {
      note("n-info", isWhole(n) ? `Whole-column/row reference (e.g. a lookup table) · → to open its ${S.dir}` : `Range of ${C.formatCount(n.cells)} cells · → to open its ${S.dir}`);
    } else if (n.value !== undefined && n !== S.root) {
      note("n-hard", "Typed-in value (an input)");
    }
    if (wrap.childElementCount) notes.appendChild(wrap);
  }

  // Follow selection: clicking a cell in Excel highlights it if it's in the
  // tree, otherwise starts a new trace from it.
  function onSelectionChanged() {
    if (S.paused) return;
    clearTimeout(S.autoTimer);
    S.autoTimer = setTimeout(async () => {
      try {
        const addr = await window.TCTrace.getSelectionAddress();
        if (sameAddr(addr, S.expectSel)) return; // we moved there ourselves
        S.expectSel = null;
        if (S.root && sameAddr(addr, S.root.address)) return setCursor(-1, false);
        const i = S.rows.findIndex((n) => sameAddr(n.address, addr));
        if (i >= 0) return setCursor(i, false);
        buildRoot();
      } catch (e) { /* ignore */ }
    }, 200);
  }

  function setDir(dir) {
    S.dir = dir;
    document.querySelectorAll("[data-dir]").forEach((b) => b.classList.toggle("on", b.dataset.dir === dir));
  }

  // ---------------------------------------------------------------- commands

  async function showPane() {
    S.paused = false;
    try { await Office.addin.showAsTaskpane(); } catch (e) { /* already visible / unsupported */ }
  }

  function command(fn) {
    return async (event) => {
      try { await fn(); } finally { if (event && event.completed) event.completed(); }
    };
  }

  function wireCommands() {
    if (!Office.actions || !Office.actions.associate) return;
    const traceCmd = (dir, all) => command(async () => {
      await showPane();
      setDir(dir);
      await buildRoot(undefined, all);
      focusPane();
    });
    Office.actions.associate("tracePrecedents", traceCmd("precedents", false));
    Office.actions.associate("traceAllPrecedents", traceCmd("precedents", true));
    Office.actions.associate("traceDependents", traceCmd("dependents", false));
    Office.actions.associate("traceAllDependents", traceCmd("dependents", true));
    Office.actions.associate("colorCode", command(() => runColor("smart")));
    Office.actions.associate("colorCodeWorkbook", command(() => runColor("workbook")));
    Office.actions.associate("showPane", command(showPane));
    // Ctrl+Alt+arrows: the same moves, usable while you're in the Excel grid.
    const nav = (k) => command(async () => {
      if (!S.root || S.paused) { await showPane(); await buildRoot(); }
      await keys[k]();
    });
    Office.actions.associate("nextRef", nav("down"));
    Office.actions.associate("prevRef", nav("up"));
    Office.actions.associate("intoRef", nav("right"));
    Office.actions.associate("backRef", nav("left"));
  }

  // -------------------------------------------------------------------- boot

  function wireUi() {
    document.querySelectorAll("[data-dir]").forEach((b) =>
      b.addEventListener("click", () => {
        setDir(b.dataset.dir);
        buildRoot(S.root && S.root.address);
        focusPane();
      })
    );
    $("exitBtn").addEventListener("click", exitTrace);
    $("home").addEventListener("click", () => setCursor(-1, true));
    $("tree").addEventListener("scroll", () => requestAnimationFrame(renderRows));
    window.addEventListener("resize", () => renderRows());
    document.addEventListener("keydown", onKey);

    $("colorBtn").addEventListener("click", () => runColor("smart"));
    $("colorWbBtn").addEventListener("click", () => {
      if (confirm("Color code every sheet in this workbook? Existing font colors on numbers and formulas (and the background of inputs) will be replaced.")) runColor("workbook");
    });
    document.querySelectorAll("input[data-color]").forEach((inp) =>
      inp.addEventListener("change", () => {
        const colors = loadColors();
        colors[inp.dataset.color] = inp.value.toUpperCase();
        saveColors(colors);
        if (inp.dataset.color === "inputFill") $("fillInputs").checked = true;
      })
    );
    $("fillInputs").addEventListener("change", () => {
      const colors = loadColors();
      colors.inputFill = $("fillInputs").checked ? document.querySelector('input[data-color="inputFill"]').value.toUpperCase() : "";
      saveColors(colors);
    });
    $("matchFill").addEventListener("click", async (e) => {
      e.preventDefault();
      try {
        const fill = await Excel.run(async (context) => {
          const f = context.workbook.getActiveCell().format.fill;
          f.load("color");
          await context.sync();
          return f.color;
        });
        if (!C.isHexColor(fill)) {
          status("The selected cell has no solid background to copy.", true);
          return;
        }
        const colors = loadColors();
        colors.inputFill = fill.toUpperCase();
        saveColors(colors);
        syncColorInputs();
        status("Input background set to " + colors.inputFill + ".");
      } catch (err) {
        status("Could not read the cell: " + errText(err), true);
      }
    });
    $("resetColors").addEventListener("click", () => { saveColors(C.DEFAULT_COLORS); syncColorInputs(); });
    syncColorInputs();
  }

  Office.onReady(async (info) => {
    wireUi();
    wireCommands();
    renderDetail();
    if (info.host !== Office.HostType.Excel) {
      status("Open this add-in inside Excel.", true);
      return;
    }
    // Load with Excel from now on, so the tab and keyboard shortcuts are live.
    try { await Office.addin.setStartupBehavior(Office.StartupBehavior.load); } catch (e) { /* older hosts */ }
    try {
      Office.addin.onVisibilityModeChanged((args) => {
        if (args.visibilityMode === "Taskpane") S.paused = false;
      });
    } catch (e) { /* older hosts */ }
    try {
      await Excel.run(async (context) => {
        context.workbook.onSelectionChanged.add(onSelectionChanged);
        await context.sync();
      });
    } catch (e) { /* follow-selection unavailable */ }
    await buildRoot();
    focusPane();
  });
})();
