# Trace & Color: an Excel add-in

Trace **precedents and dependents** in very large workbooks, and **color-code a whole model in one click** using the standard financial-modeling convention.

| Color | Meaning |
|---|---|
| 🔵 **Blue** | Hard-coded inputs: typed numbers, and "fake formulas" like `=1000` or `=12*1.05` |
| ⚫ **Black** | Formulas that only use the same sheet |
| 🟢 **Green** | Formulas linking to **other sheets** |
| 🔴 **Red** | Formulas linking to **other workbooks** |

Text labels are left alone. You can change any of the four colors in the task pane, and the add-in remembers your choice.

<img src="docs/taskpane.png" width="320" alt="Trace pane">

---

## Install (about 1 minute, no coding)

The add-in is hosted on GitHub Pages, so there's nothing to download except one file.

1. Download **[`manifest.xml`](https://raw.githubusercontent.com/serra-verdaguer/excel-trace-color/main/manifest.xml)** (right-click → *Save link as…*).
2. Load it into Excel:
   - **Excel for Windows / Mac (Microsoft 365):** `Home` → `Add-ins` → `More Add-ins` → `My Add-ins` → **`Upload My Add-in`** → choose `manifest.xml`.
   - **Excel on the web:** same path: `Home` → `Add-ins` → `More Add-ins` → `Upload My Add-in`.
3. A new **Trace & Color** tab appears on the ribbon.

> If *Upload My Add-in* is greyed out, your IT department has blocked sideloading. Ask them to deploy `manifest.xml` through **Microsoft 365 admin center → Integrated apps** (it takes them about two minutes). Then everyone in the company gets the tab automatically.

<details>
<summary>Windows: permanent install via a shared folder (alternative)</summary>

1. Put `manifest.xml` in a folder and share it (right-click → *Properties* → *Sharing*). Copy the network path, e.g. `\\MY-PC\addins`.
2. Excel → `File` → `Options` → `Trust Center` → `Trust Center Settings` → `Trusted Add-in Catalogs`. Paste the path, tick **Show in Menu**, then OK.
3. Restart Excel → `Home` → `Add-ins` → `More Add-ins` → **`Shared Folder`** → *Trace & Color* → Add.
</details>

## Use

### Ribbon: Trace & Color tab

| Button | What it does | Shortcut |
|---|---|---|
| **Precedents** | Cells the selected cell reads from (one level) | `Ctrl+Alt+Shift+P` |
| **All Precedents** | Everything upstream, all levels | |
| **Dependents** | Cells that use the selected cell (one level) | `Ctrl+Alt+Shift+D` |
| **All Dependents** | Everything downstream, all levels | |
| **Trace Pane** | Opens the pane | `Ctrl+Alt+Shift+T` |
| **Color Code** | Colors the **selection** if more than one cell is selected, otherwise the **whole sheet** | `Ctrl+Alt+Shift+C` |
| **Color Workbook** | Colors every sheet | |

On Mac the shortcuts use `⌘ ⌥ ⇧` instead of `Ctrl Alt Shift`.

### Trace pane
- **Click** a result to jump to it. Click the cell name at the top to go **back** to where you started.
- **▸** drills one level further from that result, building a tree as you audit. Circular references are detected.
- **Filter** by sheet, address, value or formula, which helps when a total has thousands of inputs.
- **select all** selects every listed range on that sheet at once.
- **Auto** re-traces whenever you click another cell.
- Results are grouped by sheet. The current sheet is listed first.

## Built for huge workbooks

- **Tracing uses Excel's own dependency engine** (`getDirectPrecedents`, `getPrecedents`, `getDirectDependents`, `getDependents`). Results come back as *rectangles*, so `=SUM(A1:A1000000)` returns one row, not a million. Lists render 200 rows at a time, with a *Show more* button.
- **Color coding does the bulk work inside Excel.** `SpecialCells` turns every numeric constant blue and every formula black in two calls, without moving any cell data. The add-in then reads **only formula cells** (in R1C1, in chunks of about 60k cells) to find the green, red and blue exceptions.
- **One check per distinct formula.** A formula copied down 1M rows is the same R1C1 string, so it is checked once and then served from a cache.
- **Rectangle compression.** Cells of the same color merge into the fewest possible rectangles before painting, so a 1M-row column becomes one format call.
- **Adaptive chunking.** If Excel refuses a payload as too large, the chunk size halves and the add-in retries automatically. Calculation and screen updating are suspended while it paints.
- If `SpecialCells` isn't available, or the sheet is extremely fragmented, the add-in falls back to a full chunked scan.

## Requirements and limits
- Excel for **Microsoft 365** (Windows or Mac), or **Excel on the web**. Excel 2016/2019 perpetual versions don't support the APIs this uses.
  - *All Precedents* needs ExcelApi 1.14, and *All Dependents* needs 1.15. The pane tells you if your build is too old.
- Excel's API doesn't follow references **into other workbooks** when tracing. Those cells are still colored red by Color Code.
- Office add-ins can't undo their changes with **Ctrl+Z**. Save before running *Color Workbook* on a model you care about.
- Font color only. Fills, borders and number formats are never touched.

## Develop

```bash
npm test                 # unit tests + a 200k-row simulated sheet
node scripts/set-base-url.js https://<you>.github.io/<your-fork>   # repoint manifest after forking
```

Hosting: GitHub Pages serves the repo root (*Settings → Pages → Deploy from branch → `main` / root*). There is no build step: the add-in is plain HTML/JS in [`addin/`](addin/).

| File | Purpose |
|---|---|
| `manifest.xml` | Ribbon tab, buttons and URLs |
| `addin/lib/core.js` | Pure logic: formula classification, rectangle compression, chunk planning (unit-tested) |
| `addin/lib/colorcode.js` | Color-coding pipeline |
| `addin/lib/trace.js` | Precedent/dependent tracing and navigation |
| `addin/app.js` | Task pane UI, ribbon commands and keyboard shortcuts (shared runtime) |

MIT licensed.
