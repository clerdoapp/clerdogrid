# Features

**Done** works fully. **Basic** works for common cases, with limits noted. **Not yet** is planned.

## Calculation

| Feature | State | Notes |
|---|---|---|
| Worksheet functions | Done | 504 functions in 13 categories, checked against published examples and an independent spreadsheet |
| Dynamic arrays and spilling, `#` spill references, `#SPILL!` | Done | |
| `LET`, `LAMBDA`, named LAMBDA functions, `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY`, `ISOMITTED` | Done | |
| `GROUPBY`, `PIVOTBY` | Basic | One aggregation function at a time |
| Legacy array formulas (Ctrl+Shift+Enter) | Done | |
| Recalculation only of what changed | Done | Dependency graph with range buckets; no recursion, so long chains are safe |
| Volatile functions (`NOW`, `RAND`, `OFFSET`, `INDIRECT`…) | Done | |
| Circular references | Done | Shown in the status bar; iterative calculation can be turned on |
| Defined names, workbook and sheet scope | Done | |
| Tables and structured references (`Table1[Col]`, `[@Col]`, `#Headers`, `#Totals`) | Done | |
| 3D references (`Sheet1:Sheet3!A1`), union and intersection operators | Done | |
| Regular expressions (`REGEXTEST`, `REGEXEXTRACT`, `REGEXREPLACE`) | Done | Atomic groups and possessive quantifiers aren't supported |
| Exponential smoothing forecasts (`FORECAST.ETS` family) | Basic | Standard additive method; results can differ slightly from other apps |
| Functions that need the internet (`WEBSERVICE`, `STOCKHISTORY`, `IMAGE`) | Not supported | Return `#CONNECT!`: the app works offline by design |

## Editing

| Feature | State | Notes |
|---|---|---|
| Typing with number, currency, percent, date and time recognition | Done | US date order |
| Formula autocomplete, argument hints, colored references, point-and-click references, F4 | Done | |
| Fill handle: numbers, linear trends, dates by day or month, weekdays, months, quarters, numbered text, formulas | Done | |
| Flash Fill | Basic | Learns splits, joins and case changes of words |
| Copy, cut and paste within the app and with other apps (text and HTML tables) | Done | |
| Paste special: values, formulas, formats, transpose, operations, skip blanks, link | Done | |
| Drag to move or copy cells | Done | |
| Undo and redo | Done | 200 steps |
| Insert and delete rows, columns and cells with references kept right | Done | |
| Find and replace, go to, go to special | Done | |
| Keyboard shortcuts for navigation, selection, editing and formatting | Done | |

## Formatting

| Feature | State | Notes |
|---|---|---|
| Fonts, sizes, bold, italic, underline styles, strikethrough, super/subscript, colors | Done | |
| Fills, patterns, two-color gradients | Done | |
| All border styles and colors | Done | |
| Alignment, wrap, indent, rotation, vertical text, merge | Done | |
| Number formats including custom codes, sections, colors, conditions, fractions, elapsed time | Done | |
| Cell styles and table styles | Done | |
| Conditional formatting: values, text, dates, top/bottom, average, duplicates, formulas, color scales, data bars, icon sets | Done | |
| Automatic row height for larger fonts | Done | Wrapped text grows only for line breaks |

## Data and analysis

| Feature | State | Notes |
|---|---|---|
| Sort on several levels, by value or color | Done | |
| Filters with value lists, conditions and top N | Done | |
| Remove duplicates, text to columns, data validation with drop-down lists | Done | |
| Goal Seek, data tables, descriptive statistics | Done | |
| Subtotals and outline grouping | Basic | Collapse by hiding rows |
| PivotTables | Basic | Static summary, or a live `PIVOTBY` formula |
| Charts: column, bar, line, area, pie, doughnut, scatter, radar, stacked, 100% stacked, combo | Done | No 3D, bubble or stock charts yet |
| Pictures | Done | |
| Notes and hyperlinks | Done | |
| Sheet protection | Basic | Stops accidental edits; it isn't encryption |

## Files and platforms

| Feature | State | Notes |
|---|---|---|
| Open and save `.xlsx` | Done | Values, formulas, styles, merges, sizes, frozen panes, names, rules, validation, filters, tables, notes, links, charts, pictures, print setup |
| Open `.xlsm` and `.xltx` | Done | Macros are kept in the file and never run |
| Open and export `.csv` / `.tsv` | Done | |
| Print and save as PDF | Basic | Through the system print dialog |
| Windows, macOS, Linux installers | Done | Built by GitHub Actions on each release |
| iPad home-screen app and web | Done | Works offline after the first visit; keeps unsaved work |
| Dark mode, touch-friendly controls | Done | |
| Pivot caches, slicers, sparklines, threaded comments, shapes other than pictures | Not yet | Dropped when a file is opened |
| Right-to-left sheets | Basic | Setting is kept; the grid still draws left to right |
| Interface languages other than English | Not yet | |
