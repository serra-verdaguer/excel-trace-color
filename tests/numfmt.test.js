// Applies number formats through a minimal fake of the Excel JS API.
const test = require("node:test");
const assert = require("node:assert/strict");

globalThis.self = globalThis;
const C = require("../addin/lib/core.js");
globalThis.TCCore = C;
require("../addin/lib/numfmt.js");
const { applyNumberFormat } = globalThis.TCNumFmt;

// A sheet as a map "r,c" -> format; areas are { r0, c0, rows, cols }.
function fakeExcel({ areas, active, activeValue = 1234, scalarOk = true, used = { r0: 0, c0: 0, rows: 50, cols: 10 } }) {
  const fmt = new Map();
  const writes = [];
  const rangeObj = (a) => ({
    rowIndex: a.r0, columnIndex: a.c0, rowCount: a.rows, columnCount: a.cols, isNullObject: false,
    load() {},
    worksheet: {
      getUsedRangeOrNullObject: () => rangeObj(used),
      getRangeByIndexes: (r0, c0, rows, cols) => rangeObj({ r0, c0, rows, cols }),
    },
    getIntersectionOrNullObject(o) {
      const r0 = Math.max(a.r0, o.rowIndex), c0 = Math.max(a.c0, o.columnIndex);
      const r1 = Math.min(a.r0 + a.rows, o.rowIndex + o.rowCount), c1 = Math.min(a.c0 + a.cols, o.columnIndex + o.columnCount);
      return r1 <= r0 || c1 <= c0 ? { isNullObject: true, load() {} } : rangeObj({ r0, c0, rows: r1 - r0, cols: c1 - c0 });
    },
    set numberFormat(v) {
      if (!Array.isArray(v) && !scalarOk) { pendingError = true; return; }
      writes.push({ a, scalar: !Array.isArray(v) });
      for (let i = 0; i < a.rows; i++) for (let j = 0; j < a.cols; j++) fmt.set(`${a.r0 + i},${a.c0 + j}`, Array.isArray(v) ? v[i][j] : v);
    },
  });
  let pendingError = false;
  const activeKey = `${active[0]},${active[1]}`;
  globalThis.Excel = {
    run: async (fn) => fn({
      application: { suspendScreenUpdatingUntilNextSync() {} },
      async sync() { if (pendingError) { pendingError = false; throw new Error("InvalidArgument"); } },
      workbook: {
        getSelectedRanges: () => ({ areas: { load() {}, items: areas.map(rangeObj) } }),
        getActiveCell: () => ({ load() {}, get numberFormat() { return [[fmt.get(activeKey) || "General"]]; }, values: [[activeValue]] }),
      },
    }),
  };
  return { fmt, writes };
}

test("one click formats the whole selection with a single write per area", async () => {
  const big = { r0: 0, c0: 0, rows: 1048576, cols: 3 }; // three whole columns
  const x = fakeExcel({ areas: [big, { r0: 5, c0: 8, rows: 2, cols: 2 }], active: [0, 0] });
  const f = await applyNumberFormat("number", {});
  assert.equal(f, '#,##0_);(#,##0);"–"_)');
  assert.equal(x.writes.length, 2);
  assert.ok(x.writes.every((w) => w.scalar));
  assert.equal(x.fmt.get("6,9"), f);
});

test("clicking again cycles decimals; another button restarts its own cycle", async () => {
  const x = fakeExcel({ areas: [{ r0: 0, c0: 0, rows: 4, cols: 4 }], active: [0, 0] });
  const n = C.numberFormatCycle("number");
  assert.equal(await applyNumberFormat("number"), n[0]);
  assert.equal(await applyNumberFormat("number"), n[1]);
  assert.equal(await applyNumberFormat("number"), n[2]);
  assert.equal(await applyNumberFormat("number"), n[0]);
  assert.equal(await applyNumberFormat("percent"), C.numberFormatCycle("percent")[0]);
  assert.equal(x.fmt.get("3,3"), C.numberFormatCycle("percent")[0]);
});

test("year: plain years vs real dates", async () => {
  fakeExcel({ areas: [{ r0: 0, c0: 0, rows: 1, cols: 5 }], active: [0, 0], activeValue: 2025 });
  assert.equal(await applyNumberFormat("year"), "0");
  assert.equal(await applyNumberFormat("year"), '0"A"');
  fakeExcel({ areas: [{ r0: 0, c0: 0, rows: 1, cols: 5 }], active: [0, 0], activeValue: 46022 });
  assert.equal(await applyNumberFormat("year"), "yyyy");
});

test("hosts without single-value writes get chunked arrays, clipped to the used range", async () => {
  const x = fakeExcel({ areas: [{ r0: 0, c0: 0, rows: 1048576, cols: 2 }], active: [0, 0], scalarOk: false, used: { r0: 0, c0: 0, rows: 120000, cols: 5 } });
  const f = await applyNumberFormat("percent", {});
  assert.ok(x.writes.length >= 2 && x.writes.every((w) => !w.scalar));
  const cells = x.writes.reduce((s, w) => s + w.a.rows * w.a.cols, 0);
  assert.equal(cells, 120000 * 2); // only the used part of the two columns
  assert.equal(x.fmt.get("119999,1"), f);
});
