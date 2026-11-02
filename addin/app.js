/* Task pane UI + ribbon/keyboard command wiring (shared runtime). */
(function () {
  "use strict";
  const C = window.TCCore;
  const PAGE = 200; // rows rendered per sheet before "Show more"
  const COLOR_KEY = "tc.colors";

  const state = {
    dir: "precedents",
    result: null, // last trace
    history: [], // origins we drilled in from (for ← back)
    items: [], // rendered rows, in keyboard order: { area, sheet, el }
    active: -1,
    expectSel: null, // selection we caused ourselves (not a user click)
    autoTimer: null,
    navTimer: null,
    seq: 0, // ignores results of superseded traces
    paused: false, // after Esc: pane hidden, selection no longer followed
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

  // ------------------------------------------------------------------- trace

  const sameAddr = (a, b) => !!a && !!b && a.replace(/'/g, "").toLowerCase() === b.replace(/'/g, "").toLowerCase();

  function kind() {
    const all = $("allLevels").checked;
    if (state.dir === "precedents") return all ? "precedents" : "directPrecedents";
    return all ? "dependents" : "directDependents";
  }

  /*
   * opts.address : trace this address instead of the current selection
   * opts.dir     : switch direction first
   * opts.keepHistory : drill / back manage the history themselves
   *
   * Every cell you analyse is remembered, so "Back" always returns to the
   * previous one – also when Follow selection switched to a new cell.
   */
  async function runTrace(opts) {
    opts = opts || {};
    if (opts.dir) setDir(opts.dir);
    const prev = state.result && state.result.origin;
    const k = kind();
    const seq = ++state.seq;
    status(window.TCTrace.KINDS[k].label + "…");
    const t0 = Date.now();
    try {
      const res = await window.TCTrace.trace(k, opts.address);
      if (seq !== state.seq) return; // a newer trace started meanwhile
      if (!opts.keepHistory && prev && !sameAddr(prev, res.origin)) remember(prev);
      state.result = res;
      renderOrigin();
      render();
      const noun = state.dir === "precedents" ? "precedent" : "dependent";
      status(`${C.formatCount(res.totalAreas)} ${noun}${res.totalAreas === 1 ? "" : "s"} · ${((Date.now() - t0) / 1000).toFixed(2)}s`);
    } catch (err) {
      if (seq !== state.seq) return;
      state.items = [];
      $("results").innerHTML = "";
      const p = document.createElement("p");
      p.className = "error";
      p.textContent = errText(err);
      $("results").appendChild(p);
      status("Trace failed", true);
    }
  }

  // Formula with references, numbers etc. coloured (see TCCore.tokenizeFormula).
  function formulaEl(formula, sheet, cls) {
    const el = document.createElement("code");
    el.className = cls || "fx";
    for (const t of C.tokenizeFormula(formula, sheet)) {
      const span = document.createElement("span");
      span.textContent = t.text;
      if (t.kind !== "op") span.className = "tk-" + t.kind;
      if (t.kind === "num" && !/^[01](\.0*)?%?$/.test(t.text)) span.classList.add("hard");
      el.appendChild(span);
    }
    return el;
  }

  // Small badges: "↗ Page 2", "⚠ 2025", "external file".
  function badgesEl(formula, sheet, long) {
    const n = C.describeFormula(formula, sheet);
    const wrap = document.createElement("div");
    wrap.className = long ? "notes" : "badges";
    const add = (cls, text, title) => {
      const b = document.createElement("span");
      b.className = "badge " + cls;
      b.textContent = text;
      if (title) b.title = title;
      wrap.appendChild(b);
    };
    for (const sh of n.sheets) add("b-sheet", long ? `Takes data from another sheet: ${sh}` : `↗ ${sh}`, "Pulls from sheet " + sh);
    if (n.external) add("b-ext", long ? "Takes data from another file" : "↗ other file", "Links to another workbook");
    if (n.hardcoded.length) {
      const list = n.hardcoded.slice(0, 4).join(", ");
      add("b-hard", long ? `Typed-in number inside the formula: ${list}. Consider moving it to an input cell.` : `⚠ ${list}`, "Hard-coded number inside the formula");
    }
    return wrap.childElementCount ? wrap : null;
  }

  function remember(addr) {
    const h = state.history;
    if (!h.length || !sameAddr(h[h.length - 1], addr)) h.push(addr);
    if (h.length > 30) h.shift();
  }

  function renderOrigin() {
    const res = state.result;
    const { sheet, local } = C.splitAddress(res.origin);
    $("oSheet").textContent = sheet || "";
    $("oSheet").classList.toggle("has", !!sheet);
    $("oCell").textContent = local;
    $("oValue").textContent = res.originCells > 1 ? `${C.formatCount(res.originCells)} cells` : res.originValue || "";
    const of = $("oFormula");
    of.innerHTML = "";
    if (res.originCells === 1 && res.originFormula) {
      of.appendChild(formulaEl(res.originFormula, sheet, "fx big"));
      const b = badgesEl(res.originFormula, sheet, false);
      if (b) of.appendChild(b);
    }
    $("originBtn").title = `Go back to ${res.origin} (Esc)`;

    // Breadcrumb: the cells analysed before this one (last few).
    const crumbs = $("crumbs");
    crumbs.innerHTML = "";
    const SHOW = 3;
    if (state.history.length > SHOW) {
      const more = document.createElement("span");
      more.className = "sep";
      more.textContent = "… ›";
      crumbs.appendChild(more);
    }
    state.history.forEach((addr, i) => {
      if (i < state.history.length - SHOW) return;
      const b = document.createElement("button");
      b.textContent = addr;
      b.title = "Go back to " + addr;
      b.addEventListener("click", () => back(state.history.length - i));
      const sep = document.createElement("span");
      sep.className = "sep";
      sep.textContent = "›";
      crumbs.append(b, sep);
    });

    const prevCell = state.history[state.history.length - 1];
    $("backBtn").hidden = !prevCell;
    $("backBtn").textContent = prevCell ? `← Back to ${prevCell}` : "";
    updateReturn();

    const n = res.totalAreas ? ` (${C.formatCount(res.totalAreas)})` : " (0)";
    $("nPrec").textContent = state.dir === "precedents" ? n : "";
    $("nDep").textContent = state.dir === "dependents" ? n : "";
  }

  function matches(area, sheet, q) {
    if (!q) return true;
    return (
      sheet.toLowerCase().includes(q) ||
      area.local.toLowerCase().includes(q) ||
      (area.value != null && String(area.value).toLowerCase().includes(q)) ||
      (area.formula != null && area.formula.toLowerCase().includes(q))
    );
  }

  function render() {
    const res = state.result;
    const root = $("results");
    root.innerHTML = "";
    state.items = [];
    state.active = -1;
    if (!res) return;
    $("filter").hidden = res.totalAreas <= 15;
    const q = $("filter").hidden ? "" : $("filter").value.trim().toLowerCase();
    if (!res.groups.length) {
      const noun = state.dir === "precedents" ? "precedents" : "dependents";
      const hint = state.dir === "precedents" ? "It is a hard-coded value or has no cell references." : "No formula uses this cell.";
      root.innerHTML = `<div class="empty"><p><b>No ${noun}.</b></p><p>${hint}</p></div>`;
      return;
    }
    const originSheet = C.splitAddress(res.origin).sheet;
    // One chip per sheet: an overview of where the references live, and a jump link.
    const chips = document.createElement("div");
    chips.className = "chips";
    root.appendChild(chips);
    for (const g of res.groups) {
      const areas = g.areas.filter((a) => matches(a, g.sheet, q));
      if (!areas.length) continue;
      const h = document.createElement("div");
      h.className = "sheet-h";
      const chip = document.createElement("button");
      chip.className = "chip" + (g.sheet === originSheet ? " own" : "");
      chip.textContent = `${g.sheet} · ${C.formatCount(areas.length)}`;
      chip.title = `Jump to the ${g.sheet} references`;
      chip.addEventListener("click", () => {
        const idx = state.items.findIndex((x) => x.sheet === g.sheet);
        if (idx >= 0) activate(idx, true);
      });
      chips.appendChild(chip);
      const name = document.createElement("span");
      name.textContent = g.sheet;
      h.appendChild(name);
      if (g.sheet === originSheet) {
        const tag = document.createElement("span");
        tag.className = "tag";
        tag.textContent = "this sheet";
        h.appendChild(tag);
      }
      const n = document.createElement("span");
      n.className = "n";
      n.textContent = C.formatCount(areas.length);
      h.appendChild(n);
      root.appendChild(h);
      renderItems(root, h, g.sheet, areas, 0);
    }
    if (!state.items.length) root.innerHTML = `<div class="empty"><p>Nothing matches the filter.</p></div>`;
  }

  // Appends rows after `anchor`; keeps state.items in on-screen order.
  function renderItems(root, anchor, sheet, areas, start) {
    const end = Math.min(areas.length, start + PAGE);
    let after = anchor;
    const newItems = [];
    for (let i = start; i < end; i++) {
      const it = { area: areas[i], sheet, el: itemEl(areas[i], sheet) };
      after.after(it.el);
      after = it.el;
      newItems.push(it);
    }
    // Insert into the flat list right after the anchor's last item.
    const prev = state.items.findIndex((x) => x.el === anchor);
    const pos = prev >= 0 ? prev + 1 : itemsBefore(anchor);
    state.items.splice(pos, 0, ...newItems);
    newItems.forEach((it) => it.el.addEventListener("click", () => activate(state.items.indexOf(it), true)));
    newItems.forEach((it) => it.el.addEventListener("dblclick", () => drill(state.items.indexOf(it))));
    if (end < areas.length) {
      const more = document.createElement("button");
      more.className = "more";
      more.textContent = `Show ${C.formatCount(Math.min(PAGE, areas.length - end))} more (${C.formatCount(areas.length - end)} left)`;
      more.addEventListener("click", () => {
        const last = after;
        more.remove();
        renderItems(root, last, sheet, areas, end);
      });
      after.after(more);
    }
  }

  function itemsBefore(node) {
    let count = 0;
    for (const it of state.items) {
      if (it.el.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING) count++;
    }
    return count;
  }

  function itemEl(area, sheet) {
    const row = document.createElement("div");
    row.className = "item";
    row.title = area.address + (area.formula ? "\n" + area.formula : "") + "\nDouble-click or → to trace into it";
    const addr = document.createElement("span");
    addr.className = "addr";
    addr.textContent = area.local;
    const val = document.createElement("span");
    val.className = "val";
    val.textContent = area.cells > 1 ? `${C.formatCount(area.cells)} cells` : area.value !== undefined ? area.value : "";
    row.append(addr, val);
    if (area.formula) {
      row.appendChild(formulaEl(area.formula, sheet));
      const b = badgesEl(area.formula, sheet, false);
      if (b) row.appendChild(b);
    }
    return row;
  }

  // Highlights row i and (optionally) selects it in Excel.
  // i = -1 means the traced cell itself (the card at the top).
  function activate(i, go) {
    if (!state.result) return;
    i = Math.max(-1, Math.min(state.items.length - 1, i));
    if (state.active >= 0 && state.items[state.active]) state.items[state.active].el.classList.remove("active");
    state.active = i;
    $("origin").classList.toggle("here", i === -1);
    updateReturn();
    renderNow(i >= 0 ? state.items[i] : null);
    const address = i === -1 ? state.result.origin : state.items[i].area.address;
    if (i >= 0) {
      state.items[i].el.classList.add("active");
      state.items[i].el.scrollIntoView({ block: "nearest" });
    } else {
      $("results").scrollTop = 0;
    }
    if (go) {
      // Debounced, so holding ↓ doesn't queue dozens of jumps.
      clearTimeout(state.navTimer);
      state.navTimer = setTimeout(() => goTo(address), 60);
    }
  }

  // Explains the reference you're on: where it is, what it holds, what to watch.
  function renderNow(it) {
    const box = $("now");
    box.hidden = !it;
    if (!it) return;
    const a = it.area;
    const originLocal = C.splitAddress(state.result.origin).local;
    const otherSheet = it.sheet !== C.splitAddress(state.result.origin).sheet;
    $("nowLabel").textContent = `${state.dir === "precedents" ? "Feeds" : "Uses"} ${originLocal} ·`;
    $("nowAddr").textContent = (otherSheet ? it.sheet + " › " : "") + a.local;
    $("nowValue").textContent = a.cells > 1 ? `${C.formatCount(a.cells)} cells` : a.value || "";
    const f = $("nowFormula");
    const notes = $("nowNotes");
    f.innerHTML = "";
    notes.innerHTML = "";
    if (a.formula) {
      f.appendChild(formulaEl(a.formula, it.sheet, "fx big"));
      const b = badgesEl(a.formula, it.sheet, true);
      if (b) notes.appendChild(b);
    } else if (a.cells === 1) {
      notes.innerHTML = '<div class="notes"><span class="badge b-hard">Typed-in value (an input, no formula)</span></div>';
    }
    if (otherSheet) {
      const n = document.createElement("div");
      n.className = "notes";
      n.innerHTML = '<span class="badge b-sheet"></span>';
      n.firstChild.textContent = `On another sheet: ${it.sheet}`;
      notes.prepend(n);
    }
  }

  // Esc: back to your cell, hide the pane, stop following the selection.
  async function exitTrace() {
    if (state.result) await goTo(state.result.origin);
    state.paused = true;
    try { await Office.addin.hide(); } catch (e) { /* not supported */ }
    status("Tracing stopped. Press a Trace button to start again.");
  }

  // "Return to D10" shows while you're looking at one of its references.
  function updateReturn() {
    const btn = $("returnBtn");
    const away = state.result && state.active >= 0;
    btn.hidden = !away;
    if (away) btn.textContent = `↩ Return to ${C.splitAddress(state.result.origin).local}`;
  }

  async function goTo(address) {
    state.expectSel = address;
    try {
      await window.TCTrace.goTo(address);
    } catch (err) {
      status("Could not go there: " + errText(err), true);
    }
  }

  // Makes row i the new traced cell (→ / Enter / double-click).
  async function drill(i) {
    const it = state.items[i];
    if (!it || !state.result) return;
    if (sameAddr(it.area.address, state.result.origin) || state.history.some((h) => sameAddr(h, it.area.address))) {
      status("Circular reference: " + it.area.address + " is already in this path.", true);
      return;
    }
    state.history.push(state.result.origin);
    await goTo(it.area.address);
    await runTrace({ address: it.area.address, keepHistory: true });
    activate(-1, false); // you're now "at" the new cell; ↓ goes to its first reference
  }

  // Goes back `steps` levels in the drill-down path (← / Backspace / breadcrumb).
  async function back(steps) {
    if (!state.history.length) return;
    const from = state.result && state.result.origin;
    let target;
    for (let s = 0; s < (steps || 1) && state.history.length; s++) target = state.history.pop();
    await goTo(target);
    await runTrace({ address: target, keepHistory: true });
    // Back at the parent cell; mark the branch we just came out of.
    const idx = state.items.findIndex((x) => sameAddr(x.area.address, from));
    activate(-1, false);
    if (idx >= 0) {
      state.items[idx].el.classList.add("visited");
      state.items[idx].el.scrollIntoView({ block: "nearest" });
    }
  }

  function setDir(dir) {
    state.dir = dir;
    document.querySelectorAll("[data-dir]").forEach((b) => b.classList.toggle("on", b.dataset.dir === dir));
  }

  function onKey(e) {
    const t = e.target;
    if (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA") && t.type !== "checkbox") {
      if (e.key === "ArrowDown" && t.id === "filter") { activate(0, true); $("results").focus(); e.preventDefault(); }
      return;
    }
    switch (e.key) {
      case "ArrowDown": activate(state.active + 1, true); break;
      case "ArrowUp": activate(state.active - 1, true); break;
      case "Home": activate(0, true); break;
      case "End": activate(state.items.length - 1, true); break;
      case "PageDown": activate(state.active + 10, true); break;
      case "PageUp": activate(state.active - 10, true); break;
      case "ArrowRight":
      case "Enter": drill(state.active < 0 ? 0 : state.active); break;
      case "ArrowLeft":
      case "Backspace": if (state.active >= 0) activate(-1, true); else back(1); break;
      case "Escape": exitTrace(); break;
      default: return;
    }
    e.preventDefault();
  }

  async function onSelectionChanged() {
    if (!$("autoTrace").checked || state.paused) return;
    clearTimeout(state.autoTimer);
    state.autoTimer = setTimeout(async () => {
      try {
        const addr = await window.TCTrace.getSelectionAddress();
        if (sameAddr(addr, state.expectSel)) return; // we moved there ourselves
        if (state.result && sameAddr(addr, state.result.origin)) return;
        state.expectSel = null;
        runTrace();
      } catch (e) { /* ignore */ }
    }, 200);
  }

  // ---------------------------------------------------------------- commands

  async function showPane() {
    state.paused = false;
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
      $("allLevels").checked = all;
      await runTrace({ dir });
      $("results").focus();
    });
    Office.actions.associate("tracePrecedents", traceCmd("precedents", false));
    Office.actions.associate("traceAllPrecedents", traceCmd("precedents", true));
    Office.actions.associate("traceDependents", traceCmd("dependents", false));
    Office.actions.associate("traceAllDependents", traceCmd("dependents", true));
    Office.actions.associate("colorCode", command(() => runColor("smart")));
    Office.actions.associate("colorCodeWorkbook", command(() => runColor("workbook")));
    Office.actions.associate("showPane", command(showPane));
    // Arrow-key navigation that works while you're in the Excel grid.
    const ensureTrace = async () => { if (!state.result) { await showPane(); await runTrace(); } };
    Office.actions.associate("nextRef", command(async () => { await ensureTrace(); activate(state.active + 1, true); }));
    Office.actions.associate("prevRef", command(async () => { await ensureTrace(); activate(state.active - 1, true); }));
    Office.actions.associate("intoRef", command(async () => { await ensureTrace(); await drill(state.active < 0 ? 0 : state.active); }));
    // Ctrl+Alt+←: back to your cell → previous cell → stop tracing.
    Office.actions.associate("backRef", command(async () => {
      if (!state.result || state.paused) return;
      if (state.active >= 0) activate(-1, true);
      else if (state.history.length) await back(1);
      else await exitTrace();
    }));
  }

  // -------------------------------------------------------------------- boot

  function wireUi() {
    const retrace = () => state.result && runTrace({ address: state.result.origin, keepHistory: true });
    document.querySelectorAll("[data-dir]").forEach((b) => b.addEventListener("click", () => { setDir(b.dataset.dir); retrace(); }));
    $("allLevels").addEventListener("change", retrace);
    $("originBtn").addEventListener("click", () => activate(-1, true));
    $("returnBtn").addEventListener("click", () => activate(-1, true));
    $("backBtn").addEventListener("click", () => back(1));
    let ft;
    $("filter").addEventListener("input", () => { clearTimeout(ft); ft = setTimeout(render, 120); });
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
    if (info.host !== Office.HostType.Excel) {
      status("Open this add-in inside Excel.", true);
      return;
    }
    // Load with Excel from now on, so the tab and keyboard shortcuts are live
    // without opening the pane first.
    try { await Office.addin.setStartupBehavior(Office.StartupBehavior.load); } catch (e) { /* older hosts */ }
    // Reopening the pane (ribbon "Trace Pane") resumes tracing after Esc.
    try {
      Office.addin.onVisibilityModeChanged((args) => {
        if (args.visibilityMode === "Taskpane") state.paused = false;
      });
    } catch (e) { /* older hosts */ }
    try {
      await Excel.run(async (context) => {
        context.workbook.onSelectionChanged.add(onSelectionChanged);
        await context.sync();
      });
    } catch (e) { /* auto-trace unavailable */ }
    const ok = window.TCTrace.supported("dependents");
    status(ok ? "Ready." : "Ready. Some trace modes need a newer Excel (Microsoft 365).");
    runTrace(); // trace whatever cell is selected right now
  });
})();
