// Runs the real colour-coding pipeline against a small in-memory fake of the
// Excel JS API, including a payload limit to exercise adaptive chunking.
const test = require("node:test");
const assert = require("node:assert/strict");

globalThis.self = globalThis;
const C = require("../addin/lib/core.js");
globalThis.TCCore = C;
require("../addin/lib/colorcode.js");
const { colorRange } = globalThis.TCColor;

globalThis.Excel = {
  SpecialCellType: { constants: "Constants", formulas: "Formulas" },
  SpecialCellValueType: { numbers: "Numbers" },
};

function parseLocal(a) {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a);
  const r0 = +m[2] - 1, c0 = C.lettersToCol(m[1]);
  return m[3] ? { r0, c0, r1: +m[4] - 1, c1: C.lettersToCol(m[3]) } : { r0, c0, r1: r0, c1: c0 };
}

// Grid cells are raw values ("=formula", number, text) or { f, vt, nf } to set
// the formula's result type and the cell's number format explicitly.
const raw = (c) => (c && typeof c === "object" ? c.f : c);
const vtOf = (c) => {
  if (c && typeof c === "object" && c.vt) return c.vt;
  const v = raw(c);
  if (typeof v === "number" || (typeof v === "string" && v[0] === "=")) return "Double";
  return v === "" ? "Empty" : "String";
};
const nfOf = (c) => (c && typeof c === "object" && c.nf) || "General";

function fakeWorkbook(name, grid, opts) {
  opts = opts || {};
  const rows = grid.length, cols = grid[0].length;
  const color = grid.map((r) => r.map(() => null));
  const fill = grid.map((r) => r.map(() => null));
  const stats = { syncs: 0, maxLoad: 0, specialCalls: 0 };
  let pendingLoads = [];

  const paintInto = (grid2, rects) => ({
    set color(v) {
      for (const r of rects) for (let i = r.r0; i <= r.r1; i++) for (let j = r.c0; j <= r.c1; j++) grid2[i][j] = v;
    },
  });
  const paintRects = (rects) => paintInto(color, rects);
  const fillRects = (rects) => paintInto(fill, rects);

  function rangeAreas(rects) {
    return {
      isNullObject: rects.length === 0,
      cellCount: rects.reduce((s, r) => s + (r.r1 - r.r0 + 1) * (r.c1 - r.c0 + 1), 0),
      areaCount: rects.length,
      areas: {
        load() {},
        items: rects.map((r) => ({ rowIndex: r.r0, columnIndex: r.c0, rowCount: r.r1 - r.r0 + 1, columnCount: r.c1 - r.c0 + 1 })),
      },
      load() {},
      format: { font: paintRects(rects) },
    };
  }

  function range(r0, c0, nr, nc) {
    const rect = { r0, c0, r1: r0 + nr - 1, c1: c0 + nc - 1 };
    return {
      rowIndex: r0, columnIndex: c0, rowCount: nr, columnCount: nc, isNullObject: false,
      load(props) {
        if (/formulasR1C1|valueTypes|numberFormat/.test(props)) pendingLoads.push(this);
      },
      get formulasR1C1() { return this._read(raw); },
      get valueTypes() { return this._read(vtOf); },
      get numberFormat() { return this._read(nfOf); },
      _read(fn) {
        const out = [];
        for (let i = 0; i < nr; i++) out.push(grid[r0 + i].slice(c0, c0 + nc).map(fn));
        return out;
      },
      getSpecialCellsOrNullObject(type, valueType) {
        stats.specialCalls++;
        if (opts.noSpecialCells) throw new Error("not supported");
        const m = new C.RectMerger();
        for (let i = rect.r0; i <= rect.r1; i++)
          for (let j = rect.c0; j <= rect.c1; j++) {
            const v = raw(grid[i][j]);
            const isF = typeof v === "string" && v[0] === "=";
            const hit = type === "Formulas" ? isF : valueType === "Numbers" && typeof v === "number";
            if (hit) m.addRun(i, j, j);
          }
        return rangeAreas(m.rects());
      },
      format: { font: paintRects([rect]) },
    };
  }

  const sheet = {
    name,
    load() {},
    getRangeByIndexes: range,
    getRanges(addr) {
      const rects = addr.split(",").map(parseLocal);
      return { format: { font: paintRects(rects), fill: fillRects(rects) } };
    },
  };
  const context = {
    application: { suspendApiCalculationUntilNextSync() {}, suspendScreenUpdatingUntilNextSync() {} },
    async sync() {
      stats.syncs++;
      const loads = pendingLoads;
      pendingLoads = [];
      for (const l of loads) {
        const n = l.rowCount * l.columnCount;
        stats.maxLoad = Math.max(stats.maxLoad, n);
        if (opts.payloadLimit && n > opts.payloadLimit) throw Object.assign(new Error("payload"), { code: "ResponsePayloadSizeLimitExceeded" });
      }
    },
  };
  return { sheet, context, color, fill, stats, used: range(0, 0, rows, cols) };
}

const COLORS = { input: "BLUE", formula: "BLACK", link: "GREEN", external: "RED", inputFill: "YELLOW" };

function sampleGrid() {
  return [
    ["Revenue", 100, 200, "=RC[-1]*1.1", "=Inputs!R1C1"],
    ["Costs", 50, "=60", "=RC[-1]*1.1", "='[Budget.xlsx]P&L'!R2C2"],
    ["Note", "text", "", "=SUM(R1C2:R2C2)", "=R1C5+Model!R1C1"],
  ];
}

const expected = [
  [null, "BLUE", "BLUE", "BLACK", "GREEN"],
  [null, "BLUE", "BLUE", "BLACK", "RED"],
  [null, null, null, "BLACK", "BLACK"],
];

for (const mode of ["fast", "fallback"]) {
  test(`colors the convention correctly (${mode} path)`, async () => {
    const wb = fakeWorkbook("Model", sampleGrid(), { noSpecialCells: mode === "fallback" });
    const s = await colorRange(wb.context, wb.sheet, wb.used, COLORS);
    assert.deepEqual(wb.color, expected);
    assert.deepEqual(wb.fill, expected.map((r) => r.map((c) => (c === "BLUE" ? "YELLOW" : null))));
    assert.equal(s.input, 4); // 100, 200, 50 and the typed-in "=60"
    assert.equal(s.formula, 4);
    assert.equal(s.link, 1);
    assert.equal(s.external, 1);
  });
}

test("only numbers are coloured: dates, text and text/date formulas stay untouched", async () => {
  for (const mode of ["fast", "fallback"]) {
    const grid = [
      // label | typed date               | year header            | text formula                        | numeric formula
      ["Period", { f: 45657, nf: "dd/mm/yyyy" }, 2025, { f: '="FY"&RC[-1]', vt: "String" }, "=RC[-1]"],
      ["Date", { f: "=EOMONTH(RC[-1],12)", nf: "mmm-yy" }, { f: 0.25, nf: "0.0%" }, { f: "=Inputs!R1C1", vt: "String" }, { f: "=Inputs!R1C1", nf: "#,##0" }],
      ["Flag", { f: "=R1C2>0", vt: "Boolean" }, { f: 0.5, nf: "[h]:mm" }, { f: "=1/0", vt: "Error" }, ""],
    ];
    const wb = fakeWorkbook("Model", grid, { noSpecialCells: mode === "fallback" });
    await colorRange(wb.context, wb.sheet, wb.used, COLORS);
    assert.deepEqual(
      wb.color,
      [
        [null, null, "BLUE", null, "BLACK"],
        [null, null, "BLUE", null, "GREEN"],
        [null, null, null, "BLUE", null],
      ],
      mode
    );
  }
});

test("input background can be switched off", async () => {
  const wb = fakeWorkbook("Model", sampleGrid());
  await colorRange(wb.context, wb.sheet, wb.used, Object.assign({}, COLORS, { inputFill: "" }));
  assert.deepEqual(wb.color, expected);
  assert.ok(wb.fill.every((r) => r.every((c) => c === null)));
});

test("large sheet: chunked reads, payload back-off, compressed painting", async () => {
  const ROWS = 200000;
  const grid = new Array(ROWS);
  for (let i = 0; i < ROWS; i++) {
    grid[i] = [i, i * 2, "=RC[-2]+RC[-1]", "=Assumptions!R1C1*RC[-1]", i % 1000 === 0 ? "=[1]Ext!A1" : "=RC[-1]"];
  }
  const wb = fakeWorkbook("Data", grid, { payloadLimit: 20000 });
  const t0 = Date.now();
  const s = await colorRange(wb.context, wb.sheet, wb.used, COLORS);
  const ms = Date.now() - t0;
  assert.equal(s.input, ROWS * 2);
  assert.equal(s.link, ROWS);
  assert.equal(s.external, ROWS / 1000);
  assert.equal(s.formula, ROWS * 2 - ROWS / 1000);
  assert.ok(wb.stats.maxLoad <= 60000);
  assert.equal(wb.color[0][4], "RED");
  assert.equal(wb.color[1][4], "BLACK");
  assert.equal(wb.color[5][3], "GREEN");
  assert.equal(wb.color[ROWS - 1][0], "BLUE");
  console.log(`    200k rows x 5 cols in ${ms} ms, ${wb.stats.syncs} syncs`);
});
