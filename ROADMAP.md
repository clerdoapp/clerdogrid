# Roadmap

Everything on the feature list is built. These are the edges that are left, roughly in order. Ideas are welcome as [issues](https://github.com/clerdoapp/clerdogrid/issues).

1. **Binary workbooks (.xlsb) from other apps:** read their color scales, data bars, icon sets, PivotTables, slicers and sparklines, and write those too (ClerdoGrid keeps them in its own part today, so they survive its own round trips).
2. **Older binary workbooks (.xls):** save 3D, stock and surface charts as themselves rather than the nearest simple chart; read charts in Excel 95 files and formulas in pre-97 worksheets (their values open today).
3. **Distinct count in PivotTables** is shown as Count by other apps, because the standard way to store it needs a data model.
4. **"Ask about your data":** understand number words ("top five") and more word forms in Slavic and Arabic.
5. **Function names in other languages:** today they stay in English everywhere, which keeps formulas portable; descriptions and argument help are translated.
6. **PDF text:** place marks on the right part of a ligature, and draw joined emoji sequences as one glyph in PDFs made by the command-line tool.
7. **Working together** needs a direct connection or your own TURN relay; there is no Clerdo server in between.
