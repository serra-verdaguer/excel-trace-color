/*
 * Precedent / dependent tracing on top of Excel's own dependency engine
 * (Range.getDirectPrecedents / getPrecedents / getDirectDependents / getDependents).
 *
 * Excel returns results as rectangles (not individual cells), so even a
 * SUM over a million rows comes back as one area – nothing is expanded
 * cell-by-cell on the add-in side.
 */
(function (root) {
  "use strict";
  const C = root.TCCore;

  const KINDS = {
    directPrecedents: { method: "getDirectPrecedents", api: "1.12", label: "Direct precedents" },
    precedents: { method: "getPrecedents", api: "1.14", label: "All precedents" },
    directDependents: { method: "getDirectDependents", api: "1.13", label: "Direct dependents" },
    dependents: { method: "getDependents", api: "1.15", label: "All dependents" },
  };
  const PREVIEW_LIMIT = 400; // single cells whose value/formula we fetch for the list

  function supported(kind) {
    try {
      return Office.context.requirements.isSetSupported("ExcelApi", KINDS[kind].api);
    } catch (e) {
      return false;
    }
  }

  function isNotFound(err) {
    return err && (err.code === "ItemNotFound" || /no (precedents|dependents)|cannot find/i.test(err.message || ""));
  }

  /*
   * Traces from `address` (e.g. "Sheet1!B12") or from the current selection.
   * Returns { origin, groups: [{ sheet, areas: [{ address, local, cells, value, formula }] }], totalCells, totalAreas }
   */
  async function trace(kind, address) {
    const k = KINDS[kind];
    if (!k) throw new Error("Unknown trace kind: " + kind);
    if (!supported(kind)) {
      throw new Error(`${k.label} needs ExcelApi ${k.api} (Microsoft 365 / Excel on the web). Update Excel to use this.`);
    }
    return Excel.run(async (context) => {
      let src;
      if (address) {
        const p = C.splitAddress(address);
        src = context.workbook.worksheets.getItem(p.sheet).getRange(p.local);
      } else {
        src = context.workbook.getSelectedRange();
      }
      src.load("address,cellCount");
      const first = src.getCell(0, 0);
      first.load("formulas,text");
      await context.sync();

      const f0 = first.formulas[0][0];
      const result = {
        kind,
        origin: src.address,
        originCells: src.cellCount,
        originValue: first.text[0][0],
        originFormula: typeof f0 === "string" && f0[0] === "=" ? f0 : null,
        groups: [],
        totalCells: 0,
        totalAreas: 0,
      };
      const wra = src[k.method]();
      wra.areas.load("items/worksheet/name,items/areas/items/address,items/areas/items/cellCount");
      try {
        await context.sync();
      } catch (err) {
        if (isNotFound(err)) return result;
        throw err;
      }

      const previews = [];
      for (const ra of wra.areas.items) {
        const group = { sheet: ra.worksheet.name, areas: [] };
        for (const r of ra.areas.items) {
          const local = C.splitAddress(r.address).local;
          const area = { address: r.address, local, cells: r.cellCount };
          group.areas.push(area);
          result.totalCells += r.cellCount;
          result.totalAreas++;
          if (r.cellCount === 1 && previews.length < PREVIEW_LIMIT) {
            const cell = ra.worksheet.getRange(local);
            cell.load("formulas,text");
            previews.push({ area, cell });
          }
        }
        result.groups.push(group);
      }
      if (previews.length) {
        await context.sync();
        for (const p of previews) {
          p.area.value = p.cell.text[0][0];
          const f = p.cell.formulas[0][0];
          if (typeof f === "string" && f[0] === "=") p.area.formula = f;
        }
      }
      // Origin sheet first, then alphabetical – the order you'd audit in.
      const originSheet = C.splitAddress(result.origin).sheet;
      result.groups.sort((a, b) => (a.sheet === originSheet ? -1 : b.sheet === originSheet ? 1 : a.sheet.localeCompare(b.sheet)));
      return result;
    });
  }

  async function goTo(address) {
    const p = C.splitAddress(address);
    await Excel.run(async (context) => {
      const ws = p.sheet ? context.workbook.worksheets.getItem(p.sheet) : context.workbook.worksheets.getActiveWorksheet();
      ws.activate();
      ws.getRange(p.local).select();
      await context.sync();
    });
  }

  // Selects every area of a group at once (chunked so huge unions don't fail).
  async function selectAll(sheet, locals) {
    const batches = C.batchAddresses(locals, 8000, 2000);
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getItem(sheet);
      ws.activate();
      ws.getRanges(batches[0]).select();
      await context.sync();
    });
    return batches.length === 1;
  }

  async function getSelectionAddress() {
    return Excel.run(async (context) => {
      const r = context.workbook.getSelectedRange();
      r.load("address");
      await context.sync();
      return r.address;
    });
  }

  root.TCTrace = { KINDS, trace, goTo, selectAll, supported, getSelectionAddress };
})(typeof self !== "undefined" ? self : this);
