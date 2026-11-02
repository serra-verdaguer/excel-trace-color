const test = require("node:test");
const assert = require("node:assert/strict");
const C = require("../addin/lib/core.js");
const { CAT } = C;

test("column letters round-trip", () => {
  for (const [i, s] of [[0, "A"], [25, "Z"], [26, "AA"], [701, "ZZ"], [702, "AAA"], [16383, "XFD"]]) {
    assert.equal(C.colToLetters(i), s);
    assert.equal(C.lettersToCol(s), i);
  }
});

test("rectToA1", () => {
  assert.equal(C.rectToA1({ r0: 0, c0: 0, r1: 0, c1: 0 }), "A1");
  assert.equal(C.rectToA1({ r0: 4, c0: 1, r1: 1048575, c1: 3 }), "B5:D1048576");
});

test("splitAddress handles quoted sheet names", () => {
  assert.deepEqual(C.splitAddress("Sheet1!A1:B2"), { sheet: "Sheet1", local: "A1:B2" });
  assert.deepEqual(C.splitAddress("'My ''P&L'' 2025'!C3"), { sheet: "My 'P&L' 2025", local: "C3" });
  assert.deepEqual(C.splitAddress("A1"), { sheet: null, local: "A1" });
});

test("classifyFormula – same sheet formulas are black", () => {
  assert.equal(C.classifyFormula("=SUM(A1:A10)", "Model"), CAT.FORMULA);
  assert.equal(C.classifyFormula("=R[-1]C*(1+R2C3)", "Model"), CAT.FORMULA);
  assert.equal(C.classifyFormula("=Model!A1+1", "Model"), CAT.FORMULA);
  assert.equal(C.classifyFormula("='model'!A1", "Model"), CAT.FORMULA);
  assert.equal(C.classifyFormula('="Total!"&A1', "Model"), CAT.FORMULA); // "!" inside a string
  assert.equal(C.classifyFormula("=IFERROR(A1,#N/A)+#REF!", "Model"), CAT.FORMULA);
  assert.equal(C.classifyFormula("=SUM(Table1[Sales])", "Model"), CAT.FORMULA);
  assert.equal(C.classifyFormula("=Table1[[#This Row],[Qty]]*2", "Model"), CAT.FORMULA);
});

test("classifyFormula – other sheets are green", () => {
  assert.equal(C.classifyFormula("=Inputs!B2", "Model"), CAT.LINK);
  assert.equal(C.classifyFormula("='Revenue Build'!B2*A1", "Model"), CAT.LINK);
  assert.equal(C.classifyFormula("=SUM(Jan:Dec!B2)", "Model"), CAT.LINK);
  assert.equal(C.classifyFormula("=Model!A1+Inputs!A1", "Model"), CAT.LINK);
  assert.equal(C.classifyFormula("='It''s'!A1", "Model"), CAT.LINK);
  assert.equal(C.classifyFormula("=SUM(Table1[Sales])+Inputs!A1", "Model"), CAT.LINK);
  assert.equal(C.classifyFormula("=Données!A1", "Model"), CAT.LINK);
});

test("classifyFormula – other workbooks are red", () => {
  assert.equal(C.classifyFormula("=[Budget.xlsx]Sheet1!A1", "Model"), CAT.EXTERNAL);
  assert.equal(C.classifyFormula("=[1]Sheet1!A1", "Model"), CAT.EXTERNAL);
  assert.equal(C.classifyFormula("='C:\\Data\\[Budget 2025.xlsx]P&L'!B4", "Model"), CAT.EXTERNAL);
  assert.equal(C.classifyFormula("='https://x.sharepoint.com/[B.xlsx]S'!A1*2", "Model"), CAT.EXTERNAL);
  assert.equal(C.classifyFormula("=[1]!MyName", "Model"), CAT.EXTERNAL);
});

test("classifyFormula – typed-in numbers are blue", () => {
  assert.equal(C.classifyFormula("=1000", "Model"), CAT.INPUT);
  assert.equal(C.classifyFormula("=12*1.05", "Model"), CAT.INPUT);
  assert.equal(C.classifyFormula("=-(3+4)/2", "Model"), CAT.INPUT);
  assert.equal(C.classifyFormula("=1E+6", "Model"), CAT.INPUT);
  assert.equal(C.classifyFormula("=A1*1.05", "Model"), CAT.FORMULA);
  assert.equal(C.classifyFormula("=PI()", "Model"), CAT.FORMULA);
});

test("makeClassifier handles raw values", () => {
  const f = C.makeClassifier("S");
  assert.equal(f(12.5), CAT.INPUT);
  assert.equal(f("label"), CAT.NONE);
  assert.equal(f(""), CAT.NONE);
  assert.equal(f(true), CAT.NONE);
  assert.equal(f("=Other!A1"), CAT.LINK);
  assert.equal(f("=Other!A1"), CAT.LINK); // cached
});

test("RectMerger merges identical runs on consecutive rows only", () => {
  const m = new C.RectMerger();
  for (let r = 0; r < 1000; r++) m.addRun(r, 2, 4);
  m.addRun(1001, 2, 4); // gap -> new rect
  m.addRun(5, 0, 0);
  const rects = m.rects().map(C.rectToA1).sort();
  assert.deepEqual(rects, ["A6", "C1002:E1002", "C1:E1000"]);
});

test("scanBlock builds minimal rectangles", () => {
  const values = [
    ["Label", 1, 2, "=A1", "=X!A1"],
    ["Label", 3, 4, "=A2", "=X!A2"],
    ["", "", 5, "=A3", "=[1]X!A3"],
  ];
  const mergers = {};
  for (const k of [CAT.INPUT, CAT.FORMULA, CAT.LINK, CAT.EXTERNAL]) mergers[k] = new C.RectMerger();
  const counts = C.scanBlock(values, 10, 1, C.makeClassifier("S"), mergers);
  const A = (k) => mergers[k].rects().map(C.rectToA1).sort();
  assert.deepEqual(A(CAT.INPUT), ["C11:D12", "D13"]);
  assert.deepEqual(A(CAT.FORMULA), ["E11:E13"]);
  assert.deepEqual(A(CAT.LINK), ["F11:F12"]);
  assert.deepEqual(A(CAT.EXTERNAL), ["F13"]);
  assert.deepEqual(counts, { [CAT.INPUT]: 5, [CAT.FORMULA]: 3, [CAT.LINK]: 2, [CAT.EXTERNAL]: 1 });
});

test("planChunks covers the block exactly", () => {
  const chunks = C.planChunks(3, 2, 1000, 70, 5000);
  let cells = 0;
  for (const c of chunks) {
    assert.ok(c.rows * c.cols <= 5000);
    cells += c.rows * c.cols;
  }
  assert.equal(cells, 70000);
  const wide = C.planChunks(0, 0, 2, 16384, 5000);
  assert.equal(wide.reduce((s, c) => s + c.rows * c.cols, 0), 2 * 16384);
});

test("batchAddresses respects limits", () => {
  const rects = [];
  for (let i = 0; i < 1000; i++) rects.push({ r0: i * 2, c0: 0, r1: i * 2, c1: 3 });
  const batches = C.batchAddresses(rects, 200, 1000);
  for (const b of batches) assert.ok(b.length <= 200);
  assert.equal(batches.join(",").split(",").length, 1000);
});

test("tokenizeFormula colours references, numbers and functions", () => {
  const t = (f) => C.tokenizeFormula(f, "Model").filter((x) => x.kind !== "op").map((x) => x.kind + ":" + x.text);
  assert.deepEqual(t("=+K47"), ["ref:K47"]);
  assert.deepEqual(t("='Page 2'!D5+2025"), ["xref:'Page 2'!D5", "num:2025"]);
  assert.deepEqual(t("=SUM($A$1:B9)*1.05"), ["fn:SUM(", "ref:$A$1:B9", "num:1.05"]);
  assert.deepEqual(t("=Model!A1+Inputs!B2"), ["ref:Model!A1", "xref:Inputs!B2"]);
  assert.deepEqual(t('=IF(A1>0,"Q4 2025",[1]Ext!C3)'), ["fn:IF(", "ref:A1", "num:0", 'str:"Q4 2025"', "ext:[1]Ext!C3"]);
  assert.deepEqual(t("=LOG10(A1)+Rate"), ["fn:LOG10(", "ref:A1"]);
  assert.deepEqual(t("=SUM(Jan:Dec!B2)"), ["fn:SUM(", "xref:Jan:Dec!B2"]);
  assert.equal(C.tokenizeFormula("=A1 + B1", "S").map((x) => x.text).join(""), "=A1 + B1");
});

test("describeFormula explains a step in plain terms", () => {
  assert.deepEqual(C.describeFormula("='Page 2'!D5+2025*12", "Model"), { sheets: ["Page 2"], external: false, hardcoded: ["2025", "12"] });
  assert.deepEqual(C.describeFormula("=IF(A1>0,A1*1,0)", "Model"), { sheets: [], external: false, hardcoded: [] });
  assert.equal(C.describeFormula("=[Budget.xlsx]P!A1", "Model").external, true);
});

test("orderByFormula follows the formula left to right", () => {
  const items = [
    { sheet: "Model", local: "C1:C3" },
    { sheet: "Inputs", local: "B2" },
    { sheet: "Model", local: "K47" },
    { sheet: "Model", local: "Z99" }, // not named in the formula
    { sheet: "Page 2", local: "D5" },
  ];
  const out = C.orderByFormula(items, "=+K47*'Page 2'!$D$5+SUM(C1:C3)+Inputs!B2", "Model").map((x) => x.sheet + "!" + x.local);
  assert.deepEqual(out, ["Model!K47", "Page 2!D5", "Model!C1:C3", "Inputs!B2", "Model!Z99"]);
  assert.deepEqual(C.a1ToRect("$B$2:A1"), { r0: 0, c0: 0, r1: 1, c1: 1 });
  assert.deepEqual(C.a1ToRect("C:C"), { r0: 0, c0: 2, r1: 1048575, c1: 2 });
});
