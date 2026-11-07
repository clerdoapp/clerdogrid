# Contributor rules

These apply to people and to automated agents working in this repository.

## Legal
- Write everything from public specifications (ECMA-376 / ISO 29500 for `.xlsx`, RFC 4180 for CSV) and your own design.
- Copy nothing from other companies' software, documents, help text, icons, templates or fonts.
- Add no GPL or AGPL code. New dependencies must be MIT, BSD, ISC, Apache-2.0 or similar, and go in `NOTICE` and `ASSETS.md`.
- The app and this repository don't name other spreadsheet apps or their makers. Use neutral words such as "other apps".

## Interface
- ClerdoGrid has its own layout: a rail, a tool panel, the formula bar, the grid and sheet tabs. Don't add a tabbed strip of tools.
- Interface text is plain and short. Say what a control does.

## Code
- TypeScript in strict mode. Keep the layers in [docs/architecture.md](docs/architecture.md).
- Cell changes go through `GridStore.change()` so they can be undone; the engine is told about every change.
- Worksheet functions return errors, they never throw. New functions go in `src/engine/functions/` with tests.
- Add comments only when the reason isn't clear from the code.
- Every file reader treats its input as untrusted: no `eval`, no remote fetches, and only `https:`, `mailto:` and `#` links.
- Keep unknown `.xlsx` parts on save where it's safe.

## Before you open a pull request
```bash
npm run typecheck
npm test
npm run build
```
New file-format work needs a save-and-reopen test in `tests/`.
