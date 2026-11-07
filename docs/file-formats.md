# File formats

ClerdoGrid reads and writes files with its own code, written from the public ECMA-376 / ISO 29500 specification. File readers treat every file as untrusted: no remote fetches, no running of macros or scripts, and only `https:`, `mailto:` and in-workbook links.

## .xlsx (open and save)

| Kept | Notes |
|---|---|
| Values, formulas, cached results | Shared formulas are expanded; newer functions are written with the prefixes the format requires |
| Dynamic-array and legacy array formulas | |
| Styles | Fonts, fills, gradients, borders, alignment, number formats, protection |
| Sheets | Names, order, colors, hidden state, column widths, row heights, outline levels, frozen panes, zoom, gridlines |
| Merges, defined names, tables, filters | |
| Conditional formatting and data validation | Including color scales, data bars and icon sets |
| Notes, threaded comments, hyperlinks | Threads are also written as notes so other apps show them |
| Charts | Standard chart parts; newer chart kinds are saved as a column chart plus ClerdoGrid's own description |
| Pictures and shapes | In the sheet's drawing part |
| Print setup, sheet protection, edit ranges | |
| Macros (`vbaProject.bin`), custom XML | Kept and written back, never run |
| Scripts, cell checkboxes, ink | In ClerdoGrid's own part inside the file; other apps ignore it |

Not kept yet: pivot caches, slicers, sparklines, form controls and theme names.

## .csv / .tsv (open and export)

Comma, semicolon, tab and pipe are detected, as is a `sep=` first line. Quotes follow RFC 4180. Numbers, dates, percentages and TRUE/FALSE are recognised the way they are when typed; `="..."` keeps text as text. Exports are UTF-8 with a byte order mark.
