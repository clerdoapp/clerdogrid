# Architecture

ClerdoGrid is one TypeScript code base that runs in the browser (and as a home-screen app on iPad) and on the desktop through Electron. There is no server.

```
src/engine     calculation: parser, evaluator, dependency graph, worksheet functions, function catalog
src/model      the workbook model shared by everything (sheets, cells, styles, rules, charts)
src/formats    .xlsx and .csv reading and writing, number formats
src/app        the editing layer: GridStore (selection, editing, undo), commands, operations, files
src/ui         React interface: grid canvas, formula bar, rail and panels, dialogs, charts
electron/      desktop shell (sandboxed window, file dialogs)
```

## Calculation
- **Parser** (`engine/parser.ts`) turns formula text into a tree and back. Formulas are stored in English with commas, like the file format.
- **Evaluator** (`engine/evaluator.ts`) walks the tree. References stay as references until a function or operator needs values, so `INDEX`, `OFFSET` and `A1:INDEX(...)` work. Arrays broadcast element by element; results bigger than one cell spill. `IF`, `IFERROR`, `CHOOSE`, `SWITCH`, `LET`, `LAMBDA` and `INDIRECT` are evaluated lazily inside the evaluator.
- **Engine** (`engine/engine.ts`) keeps a dependency index from each formula's references (with column buckets for ranges), finds everything downstream of a change, orders it with an explicit stack (no recursion, so long chains are safe), detects circular references and supports iterative calculation. Volatile functions recalculate every time.
- **Functions** live in `engine/functions/*.ts`, one file per category, each exporting `defs`. A function marks which arguments take single values; the engine lifts those over arrays automatically.

## Editing
`GridStore` (`app/store.ts`) owns the workbook, the selection, the cell editor state and undo. Every change runs through `store.change(label, scope, fn)`, which snapshots the affected cells, sheet or workbook first. Operations in `app/ops/` (structure, clipboard, fill, data, move) only use that API.

## Drawing
The grid is drawn on a canvas (`ui/render.ts`) in up to four panes for frozen rows and columns. Only visible cells are drawn, so a million rows cost nothing. The canvas sits inside a native scroll container, so touch and trackpad scrolling feel native.
