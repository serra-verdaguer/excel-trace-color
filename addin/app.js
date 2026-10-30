/* Task pane UI + ribbon/keyboard command wiring (shared runtime). */
(function () {
  "use strict";
  const C = window.TCCore;
  const PAGE = 200; // rows rendered per sheet before "Show more"
  const COLOR_KEY = "tc.colors";

  const state = {
    dir: "precedents",
    depth: "direct",
    result: null,
    navigating: false,
    autoTimer: null,
  };
  const $ = (id) => document.getElementById(id);

  // ------------------------------------------------------------------ status

  function status(text, isErr) {
    const el = $("status");
    el.textContent = text || "";
    el.classList.toggle("err", !!isErr);
  }

  function busy(on) {
    for (const id of ["traceBtn", "colorBtn", "colorWbBtn"]) $(id).disabled = on;
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
    document.querySelectorAll("input[data-color]").forEach((inp) => { inp.value = colors[inp.dataset.color]; });
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

  function kind(dir, depth) {
    if (dir === "precedents") return depth === "all" ? "precedents" : "directPrecedents";
    return depth === "all" ? "dependents" : "directDependents";
  }

  async function runTrace(opts) {
    opts = opts || {};
    if (opts.dir) setSeg("dir", opts.dir);
    if (opts.depth) setSeg("depth", opts.depth);
    const k = kind(state.dir, state.depth);
    busy(true);
    status(window.TCTrace.KINDS[k].label + "…");
    const t0 = Date.now();
    try {
      const res = await window.TCTrace.trace(k);
      state.result = res;
      $("originBtn").textContent = res.origin;
      $("originBtn").title = "Go back to " + res.origin;
      render();
      status(`${window.TCTrace.KINDS[k].label} of ${res.origin} · ${((Date.now() - t0) / 1000).toFixed(2)}s`);
    } catch (err) {
      $("results").innerHTML = "";
      const p = document.createElement("p");
      p.className = "error";
      p.textContent = errText(err);
      $("results").appendChild(p);
      status("Trace failed", true);
    } finally {
      busy(false);
    }
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
    if (!res) return;
    const q = $("filter").value.trim().toLowerCase();
    const noun = state.dir === "precedents" ? "precedents" : "dependents";
    if (!res.groups.length) {
      root.innerHTML = `<p class="empty">No ${noun} found for this cell.</p>`;
      $("summary").textContent = "";
      return;
    }
    $("summary").textContent = `${C.formatCount(res.totalCells)} cells · ${res.groups.length} sheet${res.groups.length > 1 ? "s" : ""}`;
    for (const g of res.groups) {
      const areas = g.areas.filter((a) => matches(a, g.sheet, q));
      if (!areas.length) continue;
      const det = document.createElement("details");
      det.className = "group";
      det.open = true;
      const sum = document.createElement("summary");
      const name = document.createElement("span");
      name.textContent = g.sheet;
      const sel = document.createElement("button");
      sel.className = "link small";
      sel.textContent = "select all";
      sel.title = "Select every listed range on this sheet";
      sel.addEventListener("click", async (e) => {
        e.preventDefault();
        await navigate(() => window.TCTrace.selectAll(g.sheet, areas.map((a) => a.local))).then((complete) => {
          if (complete === false) status("Selected the first batch only – the list is too long for one selection.");
        });
      });
      const count = document.createElement("span");
      count.className = "count";
      const cells = areas.reduce((s, a) => s + a.cells, 0);
      count.textContent = `${C.formatCount(cells)} cells`;
      sum.append(name, sel, count);
      det.appendChild(sum);
      const list = document.createElement("div");
      det.appendChild(list);
      renderRows(list, areas, 0, 0);
      root.appendChild(det);
    }
    if (!root.children.length) root.innerHTML = `<p class="empty">Nothing matches the filter.</p>`;
  }

  function renderRows(container, areas, start, depth, path) {
    const end = Math.min(areas.length, start + PAGE);
    for (let i = start; i < end; i++) container.appendChild(rowEl(areas[i], depth, path || []));
    if (end < areas.length) {
      const more = document.createElement("button");
      more.className = "more";
      more.textContent = `Show ${C.formatCount(Math.min(PAGE, areas.length - end))} more (${C.formatCount(areas.length - end)} left)`;
      more.addEventListener("click", () => {
        more.remove();
        renderRows(container, areas, end, depth, path);
      });
      container.appendChild(more);
    }
  }

  function rowEl(area, depth, path) {
    const wrap = document.createElement("div");
    const row = document.createElement("div");
    row.className = "row";
    row.style.setProperty("--depth", depth);
    row.title = area.address + (area.formula ? "\n" + area.formula : "");

    const exp = document.createElement("button");
    exp.className = "exp";
    exp.textContent = "▸";
    exp.title = `Show this range's direct ${state.dir}`;
    const addr = document.createElement("span");
    addr.className = "addr";
    addr.textContent = depth ? area.address : area.local;
    const detail = document.createElement("span");
    detail.className = "detail";
    if (area.cells > 1) {
      detail.textContent = `${C.formatCount(area.cells)} cells`;
    } else {
      if (area.value !== undefined) {
        const b = document.createElement("b");
        b.textContent = area.value;
        detail.appendChild(b);
      }
      if (area.formula) detail.appendChild(document.createTextNode("  " + area.formula));
    }
    row.append(exp, addr, detail);

    const children = document.createElement("div");
    children.className = "children";
    row.addEventListener("click", () => navigate(() => window.TCTrace.goTo(area.address)));
    exp.addEventListener("click", async (e) => {
      e.stopPropagation();
      if (children.childElementCount) {
        children.innerHTML = "";
        exp.textContent = "▸";
        return;
      }
      if (path.includes(area.address)) {
        status("Circular reference – " + area.address + " is already in this branch.");
        return;
      }
      exp.textContent = "…";
      try {
        const k = kind(state.dir, "direct");
        const sub = await window.TCTrace.trace(k, area.address);
        const flat = [];
        for (const g of sub.groups) for (const a of g.areas) flat.push(a);
        if (!flat.length) {
          exp.textContent = "·";
          exp.title = `No ${state.dir}`;
          return;
        }
        exp.textContent = "▾";
        renderRows(children, flat, 0, depth + 1, path.concat(area.address));
      } catch (err) {
        exp.textContent = "!";
        status(errText(err), true);
      }
    });
    wrap.append(row, children);
    return wrap;
  }

  // Selection changes we cause ourselves must not trigger auto-trace.
  async function navigate(fn) {
    state.navigating = true;
    try {
      return await fn();
    } catch (err) {
      status("Could not go there: " + errText(err), true);
    } finally {
      setTimeout(() => { state.navigating = false; }, 400);
    }
  }

  function setSeg(attr, value) {
    state[attr] = value;
    document.querySelectorAll(`[data-${attr}]`).forEach((b) => b.classList.toggle("on", b.dataset[attr] === value));
  }

  async function onSelectionChanged() {
    if (!$("autoTrace").checked || state.navigating) return;
    clearTimeout(state.autoTimer);
    state.autoTimer = setTimeout(() => runTrace(), 250);
  }

  // ---------------------------------------------------------------- commands

  async function showPane() {
    try { await Office.addin.showAsTaskpane(); } catch (e) { /* already visible / unsupported */ }
  }

  function command(fn) {
    return async (event) => {
      try { await fn(); } finally { if (event && event.completed) event.completed(); }
    };
  }

  function wireCommands() {
    if (!Office.actions || !Office.actions.associate) return;
    Office.actions.associate("tracePrecedents", command(async () => { await showPane(); await runTrace({ dir: "precedents", depth: "direct" }); }));
    Office.actions.associate("traceAllPrecedents", command(async () => { await showPane(); await runTrace({ dir: "precedents", depth: "all" }); }));
    Office.actions.associate("traceDependents", command(async () => { await showPane(); await runTrace({ dir: "dependents", depth: "direct" }); }));
    Office.actions.associate("traceAllDependents", command(async () => { await showPane(); await runTrace({ dir: "dependents", depth: "all" }); }));
    Office.actions.associate("colorCode", command(() => runColor("smart")));
    Office.actions.associate("colorCodeWorkbook", command(() => runColor("workbook")));
    Office.actions.associate("showPane", command(showPane));
  }

  // -------------------------------------------------------------------- boot

  function wireUi() {
    document.querySelectorAll("[data-dir]").forEach((b) => b.addEventListener("click", () => { setSeg("dir", b.dataset.dir); runTrace(); }));
    document.querySelectorAll("[data-depth]").forEach((b) => b.addEventListener("click", () => { setSeg("depth", b.dataset.depth); runTrace(); }));
    $("traceBtn").addEventListener("click", () => runTrace());
    $("originBtn").addEventListener("click", () => state.result && navigate(() => window.TCTrace.goTo(state.result.origin)));
    let ft;
    $("filter").addEventListener("input", () => { clearTimeout(ft); ft = setTimeout(render, 120); });
    $("colorBtn").addEventListener("click", () => runColor("smart"));
    $("colorWbBtn").addEventListener("click", () => {
      if (confirm("Color code every sheet in this workbook? Existing font colors on numbers and formulas will be replaced.")) runColor("workbook");
    });
    document.querySelectorAll("input[data-color]").forEach((inp) =>
      inp.addEventListener("change", () => {
        const colors = loadColors();
        colors[inp.dataset.color] = inp.value.toUpperCase();
        saveColors(colors);
      })
    );
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
    try {
      await Excel.run(async (context) => {
        context.workbook.onSelectionChanged.add(onSelectionChanged);
        await context.sync();
      });
    } catch (e) { /* auto-trace unavailable */ }
    const ok = window.TCTrace.supported("dependents");
    status(ok ? "Ready." : "Ready. Some trace modes need a newer Excel (Microsoft 365).");
  });
})();
