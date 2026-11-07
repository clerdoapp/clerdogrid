<p align="center"><img src="assets/clerdogrid-icon.png" width="128" alt="ClerdoGrid icon"></p>

<h1 align="center">ClerdoGrid</h1>

<p align="center"><strong>A free, open-source spreadsheet app by Clerdo.</strong><br>
Calculate, analyse and chart data, with your files staying on your own device.<br>Windows, macOS, Linux, iPad and the web.</p>

<p align="center">
<a href="https://clerdo.app">Website</a> ·
<a href="https://github.com/clerdoapp/clerdogrid/releases/latest">Download</a> ·
<a href="https://clerdoapp.github.io/clerdogrid/">Open in your browser</a> ·
<a href="LICENSE-MIT">MIT</a> / <a href="LICENSE-APACHE">Apache-2.0</a>
</p>

## What it does

- **Calculate:** about 500 worksheet functions across math, statistics, text, dates, lookup, finance, engineering, information and databases, with the same results, error values and edge cases people expect from a spreadsheet.
- **Modern formulas:** dynamic arrays that spill, `LET`, `LAMBDA` with named functions, `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `XLOOKUP`, `XMATCH`, `GROUPBY`, `PIVOTBY`, `TEXTSPLIT`, regular expressions, array constants and legacy array formulas.
- **A real calculation engine:** a dependency graph that only recalculates what changed, volatile functions, circular reference detection, iterative calculation, defined names (workbook and sheet scope), tables with structured references, 3D references across sheets and `INDIRECT`/`OFFSET`.
- **Edit fast:** typing with smart number, date, percent and currency recognition, formula autocomplete with argument hints and colored reference highlighting, point-and-click references, F4 to cycle `$`, the fill handle with series for numbers, dates, months, weekdays and numbered text, Flash Fill, Ctrl+Enter to fill a selection, copy and paste with other apps (text and formatted tables), paste special, drag to move, undo and redo.
- **Format:** fonts, sizes, colors, fills and gradients, every border style, alignment, wrap, rotation, indent, merge, custom number formats with sections, colors and conditions, cell styles and table styles.
- **Organise:** insert and delete rows, columns and cells with every reference kept right, hide, group, freeze panes, multiple sheets with colors, sort on several levels, filters, remove duplicates, text to columns, data validation with drop-down lists, find and replace, go to special.
- **Analyse:** conditional formatting (color scales, data bars, icon sets, top and bottom, duplicates, formulas), charts (column, bar, line, area, pie, doughnut, scatter, radar, stacked and combo), PivotTables, subtotals, Goal Seek, data tables and descriptive statistics.
- **Files:** opens and saves `.xlsx` (styles, formulas, merges, rules, validation, tables, notes, charts and pictures), opens `.xlsm` and `.xltx`, and opens and exports `.csv`. Print or save as PDF.
- **Private:** no account, no server, no tracking. Formulas never reach the internet.

The full list, with what's finished and what's basic, is in [docs/FEATURES.md](docs/FEATURES.md).

## Download

Get the latest build from the [Releases page](https://github.com/clerdoapp/clerdogrid/releases/latest):

| System | File |
|---|---|
| Windows | `.msi` installer, or the portable `.exe` / `.zip` |
| macOS | `.dmg` (Apple Silicon and Intel) |
| Linux | `.AppImage` (any distro), `.deb` (Ubuntu, Debian, Mint), `.rpm` (Fedora, openSUSE) or `.tar.gz` |
| iPad and web | Open [clerdoapp.github.io/clerdogrid](https://clerdoapp.github.io/clerdogrid/) in Safari, then **Share > Add to Home Screen** |

The builds aren't code-signed yet:

- **Windows:** in the SmartScreen warning, click **More info**, then **Run anyway**.
- **macOS:** right-click the app, choose **Open**, then **Open** again.
- **Linux AppImage:** make it executable first (`chmod +x`).

On iPad, the home-screen app works offline after the first visit and keeps a copy of unsaved work.

## Build from source

```bash
npm install
npm run dev     # web app at http://localhost:5173
npm test
npm run dist    # desktop installers for this system
```

More in [docs/development.md](docs/development.md) and [docs/architecture.md](docs/architecture.md).

## Contributing

Read [CONTRIBUTING.md](CONTRIBUTING.md) and [AGENTS.md](AGENTS.md). Report security problems as described in [SECURITY.md](SECURITY.md).

## Licence

The code is available under either the [MIT License](LICENSE-MIT) or the [Apache License 2.0](LICENSE-APACHE), at your choice. Third-party material is listed in [NOTICE](NOTICE) and [ASSETS.md](ASSETS.md). ClerdoGrid's name and icon belong to Clerdo.
