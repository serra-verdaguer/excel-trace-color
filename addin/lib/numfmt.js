/*
 * One-click number formats on the current selection (see TCCore.numberFormatCycle).
 *
 * Scales to any selection: the format is assigned as a single value per
 * selected area, so Excel formats the whole block itself without the add-in
 * building a cell-by-cell array. Hosts that insist on a 2D array get one in
 * chunks, clipped to the used range so whole-column selections stay cheap.
 */
(function (root) {
  "use strict";
  const C = root.TCCore;
  const CHUNK = 50000; // cells per write on the fallback path

  // Date serials for 1900–2200 vs. plain years typed as numbers.
  const looksLikeDate = (v, fmt) => typeof v === "number" && (v > 3000 || /[dmy]/i.test(String(fmt).replace(/"[^"]*"/g, "")));

  async function fillArea(context, area, fmt) {
    area.load("rowIndex,columnIndex,rowCount,columnCount,worksheet/name");
    const used = area.worksheet.getUsedRangeOrNullObject(true);
    await context.sync();
    let target = area;
    if (!used.isNullObject) {
      const clip = area.getIntersectionOrNullObject(used);
      clip.load("isNullObject,rowIndex,columnIndex,rowCount,columnCount");
      await context.sync();
      if (clip.isNullObject) return;
      target = clip;
    }
    const ws = area.worksheet;
    const rowsPer = Math.max(1, Math.floor(CHUNK / target.columnCount));
    const row = new Array(target.columnCount).fill(fmt);
    for (let r = 0; r < target.rowCount; r += rowsPer) {
      const n = Math.min(rowsPer, target.rowCount - r);
      ws.getRangeByIndexes(target.rowIndex + r, target.columnIndex, n, target.columnCount).numberFormat = new Array(n).fill(row);
      await context.sync();
    }
  }

  async function applyNumberFormat(family, opts) {
    return Excel.run(async (context) => {
      const sel = context.workbook.getSelectedRanges();
      sel.areas.load("items");
      const cell = context.workbook.getActiveCell();
      cell.load("numberFormat,values");
      await context.sync();

      const current = cell.numberFormat[0][0];
      const fmt = C.nextNumberFormat(family, current, opts, { isDate: looksLikeDate(cell.values[0][0], current) });
      try { context.application.suspendScreenUpdatingUntilNextSync(); } catch (e) { /* not on all hosts */ }
      try {
        for (const a of sel.areas.items) a.numberFormat = fmt; // one value for the whole area
        await context.sync();
      } catch (err) {
        for (const a of sel.areas.items) await fillArea(context, a, fmt);
      }
      return fmt;
    });
  }

  root.TCNumFmt = { applyNumberFormat };
})(typeof self !== "undefined" ? self : this);
