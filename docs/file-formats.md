# File formats

ClerdoGrid reads and writes every file type with its own code, written from the public specifications: ECMA-376 / ISO 29500 and the published extensions for `.xlsx`, MS-XLSB for `.xlsb`, MS-XLS and MS-CFB for `.xls`, MS-OFFCRYPTO for passwords, OpenDocument 1.3 for `.ods`, and RFC 4180 for CSV. Every reader treats its file as untrusted: no remote fetches, no running of macros or scripts, and only `https:`, `mailto:` and in-workbook links.

| Type | Open | Save | Notes |
|---|---|---|---|
| `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, `.xlam` | Yes | `.xlsx` | Everything ClerdoGrid has, in the standard parts: PivotTables with caches, slicers, timelines, sparklines, extended charts, named sheet views, page setup. Macros are kept and never run |
| `.xlsb` | Yes | Yes | Formats with no binary record are kept in ClerdoGrid's own part; the save dialog says which |
| `.xls` (97–2003) | Yes | Yes | Charts, pictures, shapes, notes, rules, validation, filters, page setup and passwords. The save dialog warns about what the format can't hold (65,536 rows, 256 columns, newer chart types) |
| Excel 95 and older workbooks | Yes | No | Values and formatting; pre-97 worksheet formulas open as their values |
| `.ods`, `.ots` | Yes | Yes | Formulas both ways, styles, charts, pictures, notes, rules, validation, protection and passwords |
| `.csv`, `.tsv`, `.txt`, `.prn` | Yes | `.csv` | Import wizard: delimiters, fixed width, encodings, number region, column types |
| PDF, web page, XML, JSON | No | Export | |

Password-protected files: `.xlsx` and `.xlsb` use AES-256 with SHA-512 when saved; agile and standard encryption open. `.xls` uses RC4 and opens the older XOR form. `.ods` uses AES-256 and opens Blowfish.

Scripts, cell checkboxes, ink and Get Data queries go in ClerdoGrid's own parts inside the file; other apps ignore them.
