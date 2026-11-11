export const USAGE = `ClerdoGrid command-line tool

Usage: clerdogrid <command> [options]

Workbooks
  info <file>                       Sheets, sizes, defined names, tables and charts
  get <file> <ref|range>            Print cells (A1, B2:D9, Sheet2!A1:C4, a name or a table)
  set <file> <ref> <value|=formula> Change one cell and save
  calc <file>                       Recalculate every formula and save the results
  eval "<formula>"                  Calculate a formula (on its own or with --file)
  convert <in> <out>                Change format by file extension
  csv <file>                        Print a sheet as CSV
  find <file> <text>                List cells whose text, value or formula matches
  validate <file>                   Report error values, #REF!, unknown functions and circular references
  diff <a> <b>                      List cell-level differences between two workbooks
  new <out.xlsx> [template]         Make a workbook from a template (see --list)

Functions
  functions [--search text]         List worksheet functions
  describe <FUNCTION>               Show a function's arguments and what it does

Other
  mcp [--root dir]...               Run as a Model Context Protocol server over stdio
  help [command]                    Show help
  --version                         Print the version

Commands that open a file take --password <text> for password-protected workbooks
(or set CLERDOGRID_PASSWORD, which keeps it out of the process list).

Run "clerdogrid help <command>" for that command's options.
Exit codes: 0 done, 1 found differences/problems/no match, 2 bad usage, 3 file problem.`;

export const COMMAND_HELP: Record<string, string> = {
  info: `clerdogrid info <file> [--json]

Shows each sheet's used range, cell and formula counts, tables, charts, defined names
and document properties.`,
  get: `clerdogrid get <file> <ref|range> [options]

  --sheet <name>   Sheet to read when the reference has no sheet (default: the active sheet)
  --json           Print values, shown text and formulas as JSON
  --raw            Print raw values instead of formatted text
  --formulas       Print formulas where cells have them

A range prints as tab-separated rows.`,
  set: `clerdogrid set <file> <ref> <value|=formula> [options]

  --sheet <name>   Sheet for the reference
  -o, --out <file> Save to another file (default: change the file in place)
  --text           Store the value as text, without reading numbers, dates or formulas

The value is read as if typed into the cell: "=SUM(A1:A3)" is a formula, "12%" and "3/4/2026"
are numbers with a format, TRUE is a logical value. Prefix with ' to keep text as it is.
Formulas that depend on the cell are recalculated before saving.`,
  calc: `clerdogrid calc <file> [-o out] [--json]

Recalculates every formula and saves the workbook so the stored results are current.
Reports circular references.`,
  eval: `clerdogrid eval "<formula>" [--file <workbook>] [--sheet <name>] [--json]

Calculates a formula. With --file, references point into that workbook.`,
  convert: `clerdogrid convert <in> <out> [options]

Formats come from the file extensions: .xlsx (.xlsm, .xltx), .csv, .tsv, .json,
plus any other formats this build includes (run with --formats to list them).

  --sheet <name>   Sheet to write to a one-sheet format (CSV, TSV, JSON with --sheet)
  --range <range>  Only this range (CSV, TSV, JSON)
  --raw            Raw values instead of formatted text (CSV, TSV)
  --formulas       Write formulas instead of results (CSV, TSV, JSON)
  --force          Replace <out> if it exists
  --encrypt <pw>   Save the .xlsx with this password
  --formats        List the formats this build reads and writes`,
  csv: `clerdogrid csv <file> [--sheet <name>] [--range <range>] [--raw] [--formulas] [--tsv]

Prints one sheet as CSV (RFC 4180) on standard output.`,
  find: `clerdogrid find <file> <text> [options]

  --sheet <name>   Only this sheet
  --match-case     Match upper and lower case exactly
  --whole          Match the whole cell
  --no-formulas    Don't search formula text
  --limit <n>      Stop after n matches (default 1000)
  --json           Print matches as JSON`,
  validate: `clerdogrid validate <file> [--json]

Recalculates and lists cells showing error values, formulas with #REF!, unknown functions,
formulas that can't be read and circular references. Exits with 1 when it finds any.`,
  diff: `clerdogrid diff <a> <b> [--values-only] [--limit n] [--json]

Compares two workbooks cell by cell: added and removed sheets, added and removed cells,
changed formulas and changed values. Exits with 1 when they differ.`,
  new: `clerdogrid new <out.xlsx> [template] [--force] [--list]

Templates: run "clerdogrid new --list". Without a template the workbook is blank.`,
  functions: `clerdogrid functions [--search <text>] [--category <name>] [--json]

Lists worksheet functions with a short description.`,
  describe: `clerdogrid describe <FUNCTION> [--json]

Shows the function's arguments, what each one is for and what it returns.`,
  mcp: `clerdogrid mcp [--root <dir>]... [--read-only] [--max-bytes <n>] [--max-cells <n>]

Runs a Model Context Protocol server on standard input and output (JSON-RPC 2.0).
Tools can only open and save files under the --root folders (default: the current folder).

  --root <dir>      Allow files under this folder; repeat for more folders
  --read-only       Don't allow saving files
  --max-bytes <n>   Largest file to open or save (default 104857600)
  --max-cells <n>   Most cells one read or write may touch (default 100000)`,
};
