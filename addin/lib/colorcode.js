/*
 * One-click financial-model colour convention:
 *   blue  = hard-coded numbers (incl. "=1000"-style typed numbers)
 *   black = formulas on the same sheet
 *   green = formulas linking to other sheets
 *   red   = formulas linking to other workbooks
 *
 * Built for very large sheets:
 *   1. Excel's own SpecialCells does the bulk work server-side: every numeric
 *      constant turns blue and every formula turns black in two calls, without
 *      sending a single cell value to the add-in.
 *   2. Only formula cells are then read (as R1C1, in adaptive chunks) to find
 *      the green/red/blue exceptions. Classification is cached per distinct
 *      R1C1 formula, so a column copied down 1M rows is classified once.
 *   3. Exceptions are compressed into rectangles and painted in batched unions.
 */
(function (root) {
  "use strict";
  const C = root.TCCore;
  const { CAT } = C;

  const START_BUDGET = 60000; // cells per read; halves automatically on payload errors
  const MIN_BUDGET = 500;
  const MAX_AREAS_FAST = 3000; // more fragmented than this -> scan used range instead
  const BATCHES_PER_SYNC = 80;

  function suspend(context) {
    try { context.application.suspendApiCalculationUntilNextSync(); } catch (e) { /* older hosts */ }
    try { context.application.suspendScreenUpdatingUntilNextSync(); } catch (e) { /* not on web */ }
  }

  async function readBlock(context, sheet, block, state) {
    // Returns { values, block } – may split the block if Excel rejects the payload.
    const range = sheet.getRangeByIndexes(block.r0, block.c0, block.rows, block.cols);
    range.load("formulasR1C1");
    try {
      await context.sync();
      return [{ values: range.formulasR1C1, block }];
    } catch (err) {
      if (state.budget <= MIN_BUDGET || block.rows * block.cols <= MIN_BUDGET) throw err;
      state.budget = Math.max(MIN_BUDGET, Math.floor(state.budget / 2));
      const parts = C.planChunks(block.r0, block.c0, block.rows, block.cols, state.budget);
      let out = [];
      for (const p of parts) out = out.concat(await readBlock(context, sheet, p, state));
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

    // ---- Pass 1: server-side bulk colouring.
    let blocks = null;
    let fullScan = false;
    try {
      suspend(context);
      const nums = target.getSpecialCellsOrNullObject(Excel.SpecialCellType.constants, Excel.SpecialCellValueType.numbers);
      const fx = target.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
      nums.load("isNullObject,cellCount");
      fx.load("isNullObject,cellCount,areaCount");
      await context.sync();
      if (!nums.isNullObject) {
        nums.format.font.color = colors.input;
        stats.input = nums.cellCount;
      }
      if (fx.isNullObject) {
        await context.sync();
        return stats;
      }
      fx.format.font.color = colors.formula;
      stats.formula = fx.cellCount;
      if (fx.areaCount <= MAX_AREAS_FAST) {
        fx.areas.load("items/rowIndex,items/columnIndex,items/rowCount,items/columnCount");
      }
      await context.sync();
      if (fx.areaCount <= MAX_AREAS_FAST) {
        blocks = fx.areas.items.map((a) => ({ r0: a.rowIndex, c0: a.columnIndex, rows: a.rowCount, cols: a.columnCount }));
      }
    } catch (err) {
      // SpecialCells unavailable (or too complex) – classify everything by scanning.
      fullScan = true;
      stats.input = 0;
      stats.formula = 0;
    }
    if (!blocks) {
      blocks = [{ r0: target.rowIndex, c0: target.columnIndex, rows: target.rowCount, cols: target.columnCount }];
    }

    // ---- Pass 2: read formulas, find the exceptions.
    const mergers = {};
    mergers[CAT.LINK] = new C.RectMerger();
    mergers[CAT.EXTERNAL] = new C.RectMerger();
    mergers[CAT.INPUT] = new C.RectMerger();
    if (fullScan) mergers[CAT.FORMULA] = new C.RectMerger();

    const totalCells = blocks.reduce((s, b) => s + b.rows * b.cols, 0);
    let doneCells = 0;
    let hardcodedFormulas = 0;
    for (const big of blocks) {
      const chunks = C.planChunks(big.r0, big.c0, big.rows, big.cols, state.budget);
      for (const chunk of chunks) {
        const parts = await readBlock(context, sheet, chunk, state);
        for (const part of parts) {
          const counts = C.scanBlock(part.values, part.block.r0, part.block.c0, classify, mergers);
          stats.link += counts[CAT.LINK] || 0;
          stats.external += counts[CAT.EXTERNAL] || 0;
          if (fullScan) {
            stats.input += counts[CAT.INPUT] || 0;
            stats.formula += counts[CAT.FORMULA] || 0;
          } else {
            hardcodedFormulas += counts[CAT.INPUT] || 0;
          }
        }
        doneCells += chunk.rows * chunk.cols;
        progress(`${sheet.name}: reading ${Math.round((100 * doneCells) / totalCells)}%`);
      }
    }
    if (!fullScan) {
      // Exceptions were counted inside the black total.
      stats.formula -= stats.link + stats.external + hardcodedFormulas;
      stats.input += hardcodedFormulas;
    }

    // ---- Pass 3: paint exceptions (black first so the others win in full-scan mode).
    progress(`${sheet.name}: painting`);
    suspend(context);
    if (fullScan) await paint(context, sheet, mergers[CAT.FORMULA].rects(), colors.formula);
    await paint(context, sheet, mergers[CAT.INPUT].rects(), colors.input);
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
