# Command-line tool

`clerdogrid` reads, edits, checks and converts workbooks from a terminal or a script, using the
same calculation engine and file code as the app. It is one JavaScript file that runs on plain
Node.js 20 or newer, with no other installs.

## Get it

Each release has `clerdogrid-cli-<version>.mjs` attached. Download it, then:

```bash
node clerdogrid-cli-0.5.0.mjs --version
```

On macOS and Linux you can make it a command:

```bash
chmod +x clerdogrid-cli-0.5.0.mjs
sudo mv clerdogrid-cli-0.5.0.mjs /usr/local/bin/clerdogrid
clerdogrid --help
```

To build it from source:

```bash
npm install
npm run cli                     # writes dist-cli/clerdogrid.mjs
node dist-cli/clerdogrid.mjs --help
```

## Commands

| Command | What it does |
| --- | --- |
| `info <file>` | Sheets, used ranges, cell and formula counts, tables, charts, defined names |
| `get <file> <ref>` | Prints a cell or range (`B4`, `A1:D9`, `Data!A1:C4`, `'My sheet'!A:A`, `3:5`, a defined name or a table name) |
| `set <file> <ref> <value>` | Changes one cell, recalculates and saves |
| `calc <file>` | Recalculates every formula and saves the results |
| `eval "<formula>"` | Calculates a formula on its own, or against a workbook with `--file` |
| `convert <in> <out>` | Changes format, chosen by the file extensions |
| `csv <file>` | Prints one sheet (or `--range`) as CSV |
| `find <file> <text>` | Lists cells whose shown text, value or formula contains the text |
| `validate <file>` | Lists error values, `#REF!`, unknown functions, unreadable formulas and circular references |
| `diff <a> <b>` | Lists cell-level differences between two workbooks |
| `new <out.xlsx> [template]` | Makes a workbook from a template (`new --list` shows them) |
| `functions [--search text]` | Lists worksheet functions |
| `describe <FUNCTION>` | Shows a function's arguments and what it returns |
| `mcp` | Runs the [MCP server](mcp.md) |

`clerdogrid help <command>` shows every option of a command. Most commands take `--json` for
output that scripts can read.

## Examples

```bash
clerdogrid info examples/sales.xlsx
clerdogrid get examples/sales.xlsx A1:E4
clerdogrid get examples/sales.xlsx E2 --formulas        # =SUM(B2:D2)
clerdogrid get examples/sales.xlsx Sales!B2:D13 --json

clerdogrid set budget.xlsx B4 2450                       # read as typed: number
clerdogrid set budget.xlsx C4 "=B4*12"                   # formula
clerdogrid set budget.xlsx A20 "'007" -o budget-2.xlsx   # text, saved to another file

clerdogrid calc report.xlsx
clerdogrid eval "=PMT(5%/12, 360, 250000)"
clerdogrid eval "=SUM(Sales!E2:E13)" --file examples/sales.xlsx

clerdogrid convert examples/sales.xlsx sales.csv
clerdogrid convert examples/sales.xlsx sales.json --formulas
clerdogrid convert data.csv data.xlsx
clerdogrid csv examples/todo.xlsx --sheet Tasks --range A1:D20

clerdogrid find examples/expenses.xlsx travel
clerdogrid validate quarterly.xlsx || echo "check the formulas"
clerdogrid diff before.xlsx after.xlsx --values-only

clerdogrid new invoice.xlsx invoice
clerdogrid functions --search lookup
clerdogrid describe XLOOKUP
```

A value that starts with `-` is read as an option; put `--` before it:
`clerdogrid set book.xlsx A1 -- -5`.

## Formats

| Extension | Read | Write | Notes |
| --- | --- | --- | --- |
| `.xlsx`, `.xlsm`, `.xltx` | yes | yes | Everything the app keeps, including charts and styles |
| `.csv`, `.txt` | yes | yes | One sheet. Writes the text each cell shows; `--raw` writes plain values |
| `.tsv` | yes | yes | Tab-separated, one sheet |
| `.json` | yes | yes | See below |

Run `clerdogrid convert --formats` to see what your build reads and writes: when the build
includes other format modules (for example older binary workbooks, OpenDocument, PDF or HTML
export), they are listed there and used by `convert`, `info` and the MCP server automatically.

JSON written by `convert` looks like this; with `--formulas`, formula cells hold `"=..."`:

```json
{ "sheets": [ { "name": "Sales", "range": "A1:E15", "rows": [["Month", "North"], ["Jan", 12400]] } ] }
```

Reading also accepts a plain array of rows (`[["a","b"],[1,2]]`) or an array of records
(`[{"name":"Pen","cost":2}]`, keys become the header row). Strings starting with `=` are formulas.

## Password-protected workbooks

When the build includes the encryption module (`convert --formats` says so), every command that
opens a file reads password-protected `.xlsx` files, and older `.xls` files with a password:

```bash
export CLERDOGRID_PASSWORD='the password'     # keeps it out of the process list
clerdogrid info locked.xlsx
clerdogrid get locked.xlsx A1:D10 --password 'the password'
clerdogrid convert open.xlsx locked.xlsx --encrypt 'new password'
```

`--encrypt` saves with AES-256 agile encryption. Saving a protected workbook again without
`--encrypt` writes it without a password.

## Exit codes

| Code | Meaning |
| --- | --- |
| 0 | Done |
| 1 | The command ran and found something: differences (`diff`), problems (`validate`), no match (`find`) |
| 2 | Bad usage: unknown command or option, a reference that doesn't exist, a formula that can't be read |
| 3 | File problem: missing, unreadable, not a workbook, or the output already exists (add `--force`) |
| 70 | Unexpected internal error |

Errors go to standard error as one line starting with `clerdogrid:`. Set `CLERDOGRID_DEBUG=1`
to also print the stack trace of an unexpected error.

## Notes

- Saving writes a temporary file next to the target and renames it, so a failed save never
  leaves half a file.
- `set` and the MCP write tools refuse to change protected sheets.
- The tool never connects to the network. Files are treated as untrusted input, as in the app.

## How it is built

The sources are in `src/cli/`: `main.ts` (commands), `core.ts` (operations shared with the MCP
server, all going through `GridStore` so edits behave as in the app), `io.ts` (file formats),
`mcp.ts` (the server) and `build.mjs`, which bundles everything with esbuild into
`dist-cli/clerdogrid.mjs`. The release workflow builds it and attaches it to each GitHub release as
`clerdogrid-cli-<version>.mjs`. Tests: `tests/cli.test.ts` (runs the built file on the example
workbooks), `tests/cli-mcp.test.ts` and `tests/cli-core.test.ts`.
