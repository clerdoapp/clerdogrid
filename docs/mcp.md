# MCP server

`clerdogrid mcp` runs a [Model Context Protocol](https://modelcontextprotocol.io) server, so an AI
assistant or agent can open, read, edit, check and save workbooks with ClerdoGrid's own engine.
It speaks JSON-RPC 2.0 over standard input and output (the MCP stdio transport), supports protocol
versions `2025-11-25`, `2025-06-18`, `2025-03-26` and `2024-11-05`, and offers tools only (no
resources or prompts). It never connects to the network.

## Add it to a client

Get the command-line tool first ([cli.md](cli.md)). Most MCP clients take a JSON config with a
command and arguments. Use absolute paths:

```json
{
  "mcpServers": {
    "clerdogrid": {
      "command": "node",
      "args": [
        "/path/to/clerdogrid-cli-0.5.0.mjs",
        "mcp",
        "--root", "/Users/me/Documents/Spreadsheets"
      ]
    }
  }
}
```

For a client with a command-line setup, the same thing is usually:

```bash
<client> mcp add clerdogrid -- node /path/to/clerdogrid-cli-0.5.0.mjs mcp --root /Users/me/Documents/Spreadsheets
```

To try it by hand:

```bash
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"test","version":"1"}}}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' \
  '{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"open_workbook","arguments":{"path":"examples/sales.xlsx"}}}' \
  | node dist-cli/clerdogrid.mjs mcp
```

## Options

| Option | Default | |
| --- | --- | --- |
| `--root <dir>` | the current folder | Files may only be opened and saved under this folder. Repeat for more folders. Relative paths start at the first one. |
| `--read-only` | off | Hides `save_workbook`; nothing is written to disk. |
| `--max-bytes <n>` | 104857600 (100 MB) | Largest file to open or save. |
| `--max-cells <n>` | 100000 | Most cells one read, write, format or export may touch. |

Paths are resolved with symbolic links followed, and anything that ends up outside the roots is
refused, including `..` and links pointing elsewhere. The server checks every tool's arguments
against its JSON Schema before running it. Log lines go to standard error so standard output
carries only protocol messages.

## Tools

Open a workbook first; the other tools take the `workbook` id it returns (`"wb1"`). They also
accept a file path and open it on first use. Changes stay in memory until `save_workbook`.

| Tool | What it does |
| --- | --- |
| `open_workbook` | Opens a file (optional `password`), or makes a new workbook (`create`, optional `template`) |
| `list_sheets` | Sheets with used ranges |
| `workbook_summary` | Sheets, counts, tables, charts, defined names, properties |
| `read_range` | Values, shown text and formulas of a range |
| `write_range` | Writes rows of values from a cell; strings are read as typed (`"=A1*2"`, `"12%"`) |
| `write_formula` | Puts a formula in a cell, or fills it over a range with references adjusted |
| `recalculate` | Recalculates everything and reports circular references |
| `evaluate_formula` | Calculates a formula without storing it, optionally against a workbook |
| `add_sheet`, `rename_sheet`, `delete_sheet` | Sheet changes; renames update formulas |
| `insert_rows`, `delete_rows` | Rows or columns (`axis: "columns"`); formulas, tables and charts move with the cells |
| `find` | Cells whose text, value or formula contains some text |
| `validate_workbook` | Error values, `#REF!`, unknown functions, unreadable formulas, circular references |
| `describe_function`, `search_functions` | The function reference |
| `create_chart` | A chart from a block of data (series names and categories found from text cells) |
| `set_format` | Number format, font, colours, alignment, wrapping, borders |
| `save_workbook` | Saves to its file or a new path; format from the extension; optional `password` for .xlsx |
| `export_csv` | A sheet or range as CSV text |
| `close_workbook` | Forgets an open workbook |

Each result is returned both as JSON text in `content` and as `structuredContent`. A tool that
fails (bad reference, unreadable formula, file outside the roots) returns `isError: true` with a
plain message the model can act on; protocol errors are kept for unknown methods and tools.

## Example exchange

```json
{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"write_range","arguments":{"workbook":"wb1","start":"G1","values":[["Share"],["=E2/$E$14"]]}}}
{"jsonrpc":"2.0","id":4,"result":{"content":[{"type":"text","text":"{\n  \"sheet\": \"Sales\",\n  \"range\": \"G1:G2\",\n  \"written\": 2\n}"}],"structuredContent":{"sheet":"Sales","range":"G1:G2","written":2},"isError":false}}
```

The server is in `src/cli/mcp.ts`; `tests/cli-mcp.test.ts` runs a full session against the built
file.
