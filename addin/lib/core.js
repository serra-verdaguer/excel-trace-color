/*
 * Pure helpers shared by the add-in and the unit tests (no Excel API here).
 * Loaded in the browser as window.TCCore, and via require() in Node tests.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.TCCore = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // Cell categories for the colour convention.
  const CAT = { NONE: 0, INPUT: 1, FORMULA: 2, LINK: 3, EXTERNAL: 4 };

  const DEFAULT_COLORS = {
    input: "#0000FF", // hard-coded numbers
    formula: "#000000", // formulas on the same sheet
    link: "#008000", // formulas pulling from other sheets
    external: "#FF0000", // formulas pulling from other workbooks
  };

  // ---------------------------------------------------------------- addresses

  function colToLetters(col) {
    // 0-based column index -> "A", "B", ..., "XFD"
    let s = "";
    let n = col + 1;
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function rectToA1(r) {
    // r = { r0, c0, r1, c1 } (0-based, inclusive)
    const a = colToLetters(r.c0) + (r.r0 + 1);
    if (r.r0 === r.r1 && r.c0 === r.c1) return a;
    return a + ":" + colToLetters(r.c1) + (r.r1 + 1);
  }

  function quoteSheet(name) {
    return "'" + String(name).replace(/'/g, "''") + "'";
  }

  // "Sheet 1!A1:B2" / "'My ''Sheet'''!A1" -> { sheet, local }
  function splitAddress(address) {
    const i = address.lastIndexOf("!");
    if (i < 0) return { sheet: null, local: address };
    let sheet = address.slice(0, i);
    if (sheet.length >= 2 && sheet[0] === "'" && sheet[sheet.length - 1] === "'") {
      sheet = sheet.slice(1, -1).replace(/''/g, "'");
    }
    return { sheet, local: address.slice(i + 1) };
  }

  // "A" -> 0, "XFD" -> 16383
  function lettersToCol(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  }

  // Groups rectangle addresses into comma-joined strings short enough for
  // worksheet.getRanges(). Long unions are where Excel gets slow or refuses.
  function batchAddresses(rects, maxChars, maxAreas) {
    maxChars = maxChars || 1500;
    maxAreas = maxAreas || 250;
    const out = [];
    let cur = [];
    let len = 0;
    for (const r of rects) {
      const a = typeof r === "string" ? r : rectToA1(r);
      if (cur.length && (len + a.length + 1 > maxChars || cur.length >= maxAreas)) {
        out.push(cur.join(","));
        cur = [];
        len = 0;
      }
      cur.push(a);
      len += a.length + 1;
    }
    if (cur.length) out.push(cur.join(","));
    return out;
  }

  // ------------------------------------------------------------ classification

  const STRING_LITERAL = /"(?:[^"]|"")*"/g;
  const ERROR_TOKENS = /#(?:REF!|NULL!|DIV\/0!|VALUE!|NAME\?|NUM!|N\/A|SPILL!|CALC!|GETTING_DATA)/gi;
  const EXTERNAL_QUOTED = /'[^']*\[[^\]]+\][^']*'!/;
  const EXTERNAL_BARE = /\[[^\[\]]+\][A-Za-z0-9_.À-￿]*!/;
  const SHEET_REF = /(?:'((?:[^']|'')+)'|([A-Za-z0-9_.À-￿]+(?::[A-Za-z0-9_.À-￿]+)?))!/g;
  // "=1000", "=12*1.05", "=-(3+4)/2" -> a typed-in number wearing a formula.
  const HARDCODED_FORMULA = /^=[\s\d.+\-*/^%()eE]*\d[\s\d.+\-*/^%()eE]*$/;

  function stripNoise(formula) {
    return formula.replace(STRING_LITERAL, '""').replace(ERROR_TOKENS, "0");
  }

  function classifyFormula(formula, sheetName) {
    const s = stripNoise(formula);
    if (EXTERNAL_QUOTED.test(s) || EXTERNAL_BARE.test(s)) return CAT.EXTERNAL;
    if (s.indexOf("!") >= 0) {
      const own = String(sheetName).toLowerCase();
      SHEET_REF.lastIndex = 0;
      let m;
      while ((m = SHEET_REF.exec(s)) !== null) {
        const name = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
        if (name.toLowerCase() !== own) return CAT.LINK;
      }
    }
    if (HARDCODED_FORMULA.test(s)) return CAT.INPUT;
    return CAT.FORMULA;
  }

  // True for number formats that display a date or time ("dd/mm/yyyy",
  // "mmm-yy", "[h]:mm", "[$-x-sysdate]dddd…"). Dates are numbers to Excel,
  // but they are labels/timelines, not inputs, so they are never coloured.
  function isDateFormat(fmt) {
    if (typeof fmt !== "string" || !fmt) return false;
    const s = fmt
      .replace(/"[^"]*"/g, "") // literal text
      .replace(/\\./g, "") // escaped characters
      .replace(/[_*]./g, "") // padding / fill
      .replace(/\[(?![hms]+\])[^\]]*\]/gi, "") // [Red], [$-409], [>100] – keep [h] [mm] [ss]
      .replace(/General/gi, "");
    return /[dmyhs]/i.test(s);
  }

  const NUMERIC_RESULT = { Double: true, Integer: true, Error: true };

  // Classifies one cell from what Excel returns for it:
  //   v  : formulasR1C1 entry (number / text / "=formula"); undefined for
  //        cells already known to be numeric constants
  //   vt : valueTypes entry (what a formula evaluates to)
  //   nf : numberFormat entry
  // Only numbers are coloured: text, booleans, blanks and anything shown as a
  // date/time are left untouched – constants and formula results alike.
  // Everything is cached: big models repeat the same R1C1 formula and the same
  // number format down whole columns, so most lookups are a Map hit.
  function makeClassifier(sheetName, maxCache) {
    const fxCache = new Map();
    const fmtCache = new Map();
    maxCache = maxCache || 100000;
    function isDate(nf) {
      if (nf === undefined || nf === "General") return false;
      let d = fmtCache.get(nf);
      if (d === undefined) {
        d = isDateFormat(nf);
        if (fmtCache.size >= maxCache) fmtCache.clear();
        fmtCache.set(nf, d);
      }
      return d;
    }
    return function classify(v, vt, nf) {
      if (v === undefined || typeof v === "number") return isDate(nf) ? CAT.NONE : CAT.INPUT;
      if (typeof v !== "string" || v.charCodeAt(0) !== 61 /* = */) return CAT.NONE;
      if (vt !== undefined && !NUMERIC_RESULT[vt]) return CAT.NONE;
      if (isDate(nf)) return CAT.NONE;
      let c = fxCache.get(v);
      if (c === undefined) {
        c = classifyFormula(v, sheetName);
        if (fxCache.size >= maxCache) fxCache.clear();
        fxCache.set(v, c);
      }
      return c;
    };
  }

  // ---------------------------------------------------- rectangle compression

  // Collects same-category cells into as few rectangles as possible:
  // horizontal runs per row, then identical runs on consecutive rows merge.
  // A 1M-row column of identical formulas becomes a single rectangle.
  class RectMerger {
    constructor() {
      this.open = new Map(); // "c0:c1" -> rect
      this.done = [];
    }
    addRun(row, c0, c1) {
      const key = c0 + ":" + c1;
      const rect = this.open.get(key);
      if (rect && rect.r1 === row - 1) {
        rect.r1 = row;
        return;
      }
      if (rect) this.done.push(rect);
      this.open.set(key, { r0: row, c0, r1: row, c1 });
    }
    rects() {
      return this.done.concat(Array.from(this.open.values()));
    }
  }

  // Scans a block of cells and feeds category runs into mergers.
  //   data     : { rows, cols, formulas?, valueTypes?, numberFormat? } – 2D arrays
  //              as returned by Excel (formulas = range.formulasR1C1). A plain
  //              2D array is accepted as shorthand for { formulas }.
  //   r0, c0   : absolute 0-based position of the block's top-left cell
  //   classify : function(value, valueType, numberFormat) -> CAT
  //   mergers  : { [CAT]: RectMerger } – only categories present are recorded
  function scanBlock(data, r0, c0, classify, mergers) {
    if (Array.isArray(data)) data = { formulas: data, rows: data.length, cols: data.length ? data[0].length : 0 };
    const F = data.formulas, VT = data.valueTypes, NF = data.numberFormat;
    const counts = {};
    for (let i = 0; i < data.rows; i++) {
      const fr = F && F[i], vr = VT && VT[i], nr = NF && NF[i];
      const absRow = r0 + i;
      let runCat = CAT.NONE;
      let runStart = 0;
      for (let j = 0; j <= data.cols; j++) {
        const cat = j < data.cols ? classify(fr && fr[j], vr && vr[j], nr && nr[j]) : CAT.NONE;
        if (cat !== runCat) {
          if (runCat !== CAT.NONE && mergers[runCat]) {
            mergers[runCat].addRun(absRow, c0 + runStart, c0 + j - 1);
            counts[runCat] = (counts[runCat] || 0) + (j - runStart);
          }
          runCat = cat;
          runStart = j;
        }
      }
    }
    return counts;
  }

  // Splits a big rectangle into blocks of at most `budget` cells.
  function planChunks(r0, c0, rows, cols, budget) {
    const colsPer = Math.max(1, Math.min(cols, budget));
    const rowsPer = Math.max(1, Math.floor(budget / colsPer));
    const out = [];
    for (let r = 0; r < rows; r += rowsPer) {
      for (let c = 0; c < cols; c += colsPer) {
        out.push({ r0: r0 + r, c0: c0 + c, rows: Math.min(rowsPer, rows - r), cols: Math.min(colsPer, cols - c) });
      }
    }
    return out;
  }

  // ---------------------------------------------------------------- utilities

  function formatCount(n) {
    return Number(n).toLocaleString("en-US");
  }

  function isHexColor(s) {
    return /^#[0-9A-Fa-f]{6}$/.test(s);
  }

  return {
    CAT,
    DEFAULT_COLORS,
    colToLetters,
    lettersToCol,
    rectToA1,
    quoteSheet,
    splitAddress,
    batchAddresses,
    classifyFormula,
    makeClassifier,
    isDateFormat,
    RectMerger,
    scanBlock,
    planChunks,
    formatCount,
    isHexColor,
  };
});
