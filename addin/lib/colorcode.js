/*
 * One-click financial-model colour convention, applied to NUMBERS ONLY:
 *   blue  = hard-coded numbers (incl. "=1000"-style typed numbers)
 *   black = formulas on the same sheet
 *   green = formulas linking to other sheets
 *   red   = formulas linking to other workbooks
 * Text, booleans, blanks, dates/times and formulas that return text or dates
 * keep whatever font colour they already have.
 *
 * Built for very large sheets:
 *   1. Excel's SpecialCells finds numeric constants and formulas server-side,
 *      so labels and empty space are never read.
 *   2. Those areas are read in adaptive chunks (number format; plus R1C1 formula
 *      and result type for formulas). Classification is cached per distinct
 *      formula / number format, so a column copied down 1M rows costs one check.
 *   3. Cells are compressed into rectangles and painted in batched unions.
 */
(function (root) {
  "use strict";
  const C = root.TCCore;
  const { CAT } = C;

  const START_BUDGET = 40000; // cells per read; halves automatically on payload errors
  const MIN_BUDGET = 500;
  const MAX_AREAS_FAST = 3000; // more fragmented than this -> scan used range instead
  const BATCHES_PER_SYNC = 80;
  const PROPS = {
    constants: ["numberFormat"],
    formulas: ["formulasR1C1", "valueTypes", "numberFormat"],
  };

  function suspend(context) {
    try { context.application.suspendApiCalculationUntilNextSync(); } catch (e) { /* older hosts */ }
    try { context.application.suspendScreenUpdatingUntilNextSync(); } catch (e) { /* not on web */ }
  }

  // Reads one block; splits it if Excel rejects the payload as too large.
  async function readBlock(context, sheet, block, props, state) {
    const range = sheet.getRangeByIndexes(block.r0, block.c0, block.rows, block.cols);
    range.load(props.join(","));
    try {
      await context.sync();
      const data = { rows: block.rows, cols: block.cols };
      if (props.includes("formulasR1C1")) data.formulas = range.formulasR1C1;
      if (props.includes("valueTypes")) data.valueTypes = range.valueTypes;
      data.numberFormat = range.numberFormat;
      return [{ data, block }];
    } catch (err) {
      if (state.budget <= MIN_BUDGET || block.rows * block.cols <= MIN_BUDGET) throw err;
      state.budget = Math.max(MIN_BUDGET, Math.floor(state.budget / 2));
      const parts = C.planChunks(block.r0, block.c0, block.rows, block.cols, state.budget);
      let out = [];
      for (const p of parts) out = out.concat(await readBlock(context, sheet, p, props, state));
      return out;
    }
  }

  async function paint(context, sheet, rects, color) {
    const batches = C.batchAddresses(rects);
    for (let i = 0; i < batches.length; i++) {
      sheet.getRanges(batches[i]).format.font.color = color;
      if ((i + 1) % BATCHES_PER_SYNC === 0) {
        suspend(context);
        await context.sync();
      }
    }
    await context.sync();
  }

  const toBlock = (a) => ({ r0: a.rowIndex, c0: a.columnIndex, rows: a.rowCount, cols: a.columnCount });

  // Finds the blocks worth reading: numeric constants and formulas via
  // SpecialCells, or the whole target if that is unavailable or too fragmented.
  async function planBlocks(context, target) {
    const whole = [{ block: toBlock(target), props: PROPS.formulas }];
    try {
      const nums = target.getSpecialCellsOrNullObject(Excel.SpecialCellType.constants, Excel.SpecialCellValueType.numbers);
      const fx = target.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
      nums.load("isNullObject,areaCount");
      fx.load("isNullObject,areaCount");
      await context.sync();
      const areaCount = (nums.isNullObject ? 0 : nums.areaCount) + (fx.isNullObject ? 0 : fx.areaCount);
      if (areaCount > MAX_AREAS_FAST) return whole;
      const AREA_PROPS = "items/rowIndex,items/columnIndex,items/rowCount,items/columnCount";
      if (!nums.isNullObject) nums.areas.load(AREA_PROPS);
      if (!fx.isNullObject) fx.areas.load(AREA_PROPS);
      await context.sync();
      const out = [];
      if (!nums.isNullObject) for (const a of nums.areas.items) out.push({ block: toBlock(a), props: PROPS.constants });
      if (!fx.isNullObject) for (const a of fx.areas.items) out.push({ block: toBlock(a), props: PROPS.formulas });
      return out;
    } catch (err) {
      return whole; // SpecialCells unavailable (or too complex) – scan everything.
    }
  }

  /*
   * Colour-codes one rectangular target on a sheet.
   *   target: Excel.Range (selection or used range)
   *   onProgress(text)
   * Returns counts per category.
   */
  async function colorRange(context, sheet, target, colors, onProgress) {
    const progress = onProgress || function () {};
    sheet.load("name");
    target.load("rowIndex,columnIndex,rowCount,columnCount");
    await context.sync();

    const stats = { input: 0, formula: 0, link: 0, external: 0, cells: target.rowCount * target.columnCount };
    const classify = C.makeClassifier(sheet.name);
    const state = { budget: START_BUDGET };
    const mergers = {};
    for (const k of [CAT.INPUT, CAT.FORMULA, CAT.LINK, CAT.EXTERNAL]) mergers[k] = new C.RectMerger();

    const jobs = await planBlocks(context, target);
    const totalCells = jobs.reduce((s, j) => s + j.block.rows * j.block.cols, 0);
    let doneCells = 0;
    for (const job of jobs) {
      const big = job.block;
      for (const chunk of C.planChunks(big.r0, big.c0, big.rows, big.cols, state.budget)) {
        for (const part of await readBlock(context, sheet, chunk, job.props, state)) {
          const counts = C.scanBlock(part.data, part.block.r0, part.block.c0, classify, mergers);
          stats.input += counts[CAT.INPUT] || 0;
          stats.formula += counts[CAT.FORMULA] || 0;
          stats.link += counts[CAT.LINK] || 0;
          stats.external += counts[CAT.EXTERNAL] || 0;
        }
        doneCells += chunk.rows * chunk.cols;
        progress(`${sheet.name}: reading ${Math.round((100 * doneCells) / totalCells)}%`);
      }
    }

    progress(`${sheet.name}: painting`);
    suspend(context);
    await paint(context, sheet, mergers[CAT.INPUT].rects(), colors.input);
    await paint(context, sheet, mergers[CAT.FORMULA].rects(), colors.formula);
    await paint(context, sheet, mergers[CAT.LINK].rects(), colors.link);
    await paint(context, sheet, mergers[CAT.EXTERNAL].rects(), colors.external);
    return stats;
  }

  function addStats(a, b) {
    for (const k of ["input", "formula", "link", "external", "cells"]) a[k] = (a[k] || 0) + (b[k] || 0);
    return a;
  }

  /*
   * scope: "smart"     – the selection if it is more than one cell, otherwise the active sheet
   *        "selection" | "sheet" | "workbook"
   */
  async function colorCode(scope, colors, onProgress) {
    colors = Object.assign({}, C.DEFAULT_COLORS, colors || {});
    const t0 = Date.now();
    const total = { sheets: 0 };
    let label = "";
    await Excel.run(async (context) => {
      const sel = context.workbook.getSelectedRange();
      sel.load("cellCount,address");
      const active = context.workbook.worksheets.getActiveWorksheet();
      active.load("name");
      await context.sync();

      let mode = scope;
      if (scope === "smart") mode = sel.cellCount > 1 ? "selection" : "sheet";

      const jobs = [];
      if (mode === "selection") {
        // Clip whole-column / whole-row selections to the data actually there.
        jobs.push({ sheet: active, target: () => sel.getIntersectionOrNullObject(active.getUsedRange(true)) });
        label = sel.address;
      } else if (mode === "sheet") {
        jobs.push({ sheet: active, target: () => active.getUsedRangeOrNullObject(true) });
        label = active.name;
      } else {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name,items/visibility");
        await context.sync();
        for (const ws of sheets.items) jobs.push({ sheet: ws, target: () => ws.getUsedRangeOrNullObject(true) });
        label = `${sheets.items.length} sheets`;
      }

      for (const job of jobs) {
        const target = job.target();
        target.load("isNullObject");
        try {
          await context.sync();
        } catch (e) {
          continue; // empty sheet
        }
        if (target.isNullObject) continue;
        addStats(total, await colorRange(context, job.sheet, target, colors, onProgress));
        total.sheets++;
      }
    });
    total.label = label;
    total.ms = Date.now() - t0;
    return total;
  }

  root.TCColor = { colorCode, colorRange };
})(typeof self !== "undefined" ? self : this);
