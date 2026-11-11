# Features

**Done** works fully. Notes say where behaviour depends on another app or on the network.

## Calculation

| Feature | State | Notes |
|---|---|---|
| Worksheet functions | Done | 514 functions; see [functions.md](functions.md) |
| Dynamic arrays and spilling, `#` spill references, `#SPILL!` | Done | |
| `LET`, `LAMBDA`, named LAMBDA functions, `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY`, `ISOMITTED` | Done | |
| `GROUPBY`, `PIVOTBY`, `PERCENTOF` | Done | Several aggregations at once, every optional argument |
| Trim-range operators `.:` `:.` `.:.` and `TRIMRANGE` | Done | |
| Legacy array formulas (Ctrl+Shift+Enter) | Done | |
| Recalculation only of what changed | Done | Dependency graph with range buckets; no recursion, so long chains are safe |
| Automatic, automatic except data tables, and manual calculation (F9, Shift+F9) | Done | |
| Iterative calculation, precision as displayed, 1904 date system | Done | |
| A1 and R1C1 reference styles | Done | |
| Volatile functions (`NOW`, `RAND`, `OFFSET`, `INDIRECT`…) | Done | |
| Circular references | Done | |
| Defined names, workbook and sheet scope | Done | |
| Tables and structured references | Done | |
| 3D references, union and intersection operators | Done | |
| Regular expressions (`REGEXTEST`, `REGEXEXTRACT`, `REGEXREPLACE`) | Done | Own engine with the full syntax: atomic groups, possessive quantifiers, lookbehind, recursion, conditionals, backtracking verbs, Unicode properties and full case folding |
| Forecasting (`FORECAST.ETS` family) | Done | Additive triple smoothing with season detection, all aggregation and completion options, confidence intervals and statistics |
| `FILTERXML` with XPath 1.0 and namespaces | Done | |
| Internet functions (`WEBSERVICE`, `STOCKHISTORY`, `IMAGE`, `IMPORTCSV`, `IMPORTTEXT`) | Done | Off until you allow them under File > Online functions; https only; on desktop they go through a checked relay that refuses private-network addresses |
| `TRANSLATE`, `DETECTLANGUAGE` | Done | Use the device's built-in translator when the system has one |
| Error checking with green triangles and fixes | Done | Six rules; ignored errors are saved in the file |

## Editing

| Feature | State | Notes |
|---|---|---|
| Typing with number, currency, percent, date and time recognition | Done | Follows the chosen region: date order, separators, currency, native digits |
| Formula autocomplete, argument hints, colored references, point-and-click references, F4 | Done | Argument descriptions in all 19 languages |
| Region-style formulas | Done | `;` lists and `,` decimals where the region uses them; files store the standard form |
| AutoComplete for values in a column | Done | |
| Fill handle: numbers, trends, dates, weekdays, months, quarters, numbered text, formulas | Done | |
| Flash Fill | Done | Learns from examples (word, number and case pieces, dates, initials); shows grey suggestions after two examples |
| Copy, cut and paste within the app and with other apps | Done | Copy visible cells only (Alt+;) |
| Paste special, drag to move or copy | Done | |
| Undo and redo | Done | 200 steps |
| Insert and delete rows, columns and cells with references kept right | Done | |
| Find and replace, go to, go to special | Done | |

## Formatting

| Feature | State | Notes |
|---|---|---|
| Fonts, sizes, styles, colors, fills, gradients, all borders, alignment, rotation, merge | Done | |
| Number formats including custom codes, sections, colors, conditions, locale codes | Done | |
| Cell styles and table styles | Done | |
| Conditional formatting, including color scales, data bars and icon sets | Done | |
| Cell checkboxes | Done | |

## Data and analysis

| Feature | State | Notes |
|---|---|---|
| Sort on several levels, filters, remove duplicates, text to columns, data validation | Done | |
| Goal Seek, data tables, descriptive statistics, Consolidate | Done | |
| Outline: group rows and columns, level and +/- buttons, Auto Outline, Subtotal | Done | |
| PivotTables | Done | Field list, every summary and show-values-as option, calculated fields and items, grouping, filters, layouts, styles, drill-down, GETPIVOTDATA, refresh; saved with their data cache so they stay live in other apps |
| PivotCharts | Done | |
| Slicers and timelines | Done | For tables and PivotTables, connected to several PivotTables |
| Charts | Done | 66 types: column, bar, line, area, pie, doughnut, pie of pie, bar of pie, scatter, bubble, radar, stock, surface and contour, 3D column, bar, line, area and pie, cylinder, cone, pyramid, histogram, Pareto, box and whisker, waterfall, funnel, treemap, sunburst, filled map, combo |
| Chart elements | Done | Trendlines with equation and R², error bars, up/down bars, high-low, drop and series lines, data table, full data label and axis options, secondary axes, per-point formatting, chart sheets, templates |
| Sparklines | Done | Line, column and win/loss |
| Insights and "Ask about your data" | Done | Questions typed in any of the 19 languages |
| Get Data from text/CSV and other workbooks, with Refresh | Done | |
| Pictures, shapes, text boxes, ink | Done | |

## Review and sharing

| Feature | State | Notes |
|---|---|---|
| Threaded comments, notes, hyperlinks | Done | |
| Spelling, accessibility check, performance check, workbook statistics | Done | |
| Sheet protection, ranges that stay editable, workbook structure protection | Done | Both password hash forms other apps use |
| Password-protected files | Done | Opens and saves encrypted .xlsx, .xlsb, .xls and .ods |
| Version history | Done | Local versions with preview, restore and names |
| Working together live | Done | Codes or QR to join, everyone's cursor on the sheet, follow a person, optional TURN relay; see [collaboration.md](collaboration.md) |
| Share links | Done | The workbook travels in the link, up to 2 MB |

## Page layout and printing

| Feature | State | Notes |
|---|---|---|
| Page setup: paper, orientation, margins, scaling and fit to pages, centering, page order | Done | |
| Headers and footers with all codes and pictures | Done | |
| Print area, print titles, manual page breaks | Done | |
| Page break preview and page layout view | Done | |
| Print preview and printing | Done | |
| PDF export | Done | Own writer: real embedded fonts, selectable text including Arabic and Chinese, Japanese and Korean, charts as vector drawings |

## View

| Feature | State | Notes |
|---|---|---|
| Freeze panes and split panes | Done | |
| Named sheet views and custom views | Done | Each sheet view keeps its own filter and sort order |
| Focus cell, navigation pane, full screen, zoom to selection | Done | |
| New window on the same workbook | Done | Both windows can edit |
| Dark mode, touch-friendly controls | Done | |
| Interface languages | Done | 19: English, Arabic, German, Spanish, French, Italian, Portuguese, Dutch, Swedish, Polish, Turkish, Russian, Ukrainian, Chinese, Japanese, Korean, Hindi, Indonesian and Hebrew |
| Right to left | Done | The whole app mirrors for Arabic and Hebrew; right-to-left sheets draw with column A on the right |

## Automate

| Feature | State | Notes |
|---|---|---|
| Scripts | Done | Run in a sandboxed worker without network access; saved in the file |
| Command-line tool and MCP server | Done | See [cli.md](cli.md) and [mcp.md](mcp.md) |

## Files and platforms

| Feature | State | Notes |
|---|---|---|
| `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xlam` | Done | Open and save |
| `.xlsb` | Done | Open and save |
| `.xls` (97–2003) | Done | Open and save, including charts, pictures and shapes; Excel 95 and older workbooks open too |
| `.ods`, `.ots` | Done | Open and save |
| `.csv`, `.tsv`, `.txt`, `.prn` | Done | Import wizard with delimiters, fixed width, encodings and column types |
| Export to PDF, web page, XML and JSON | Done | |
| Windows, macOS, Linux installers; iPad home-screen app and web | Done | Works offline after the first visit |

Details of what each file type keeps are in [file-formats.md](file-formats.md).
