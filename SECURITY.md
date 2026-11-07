# Security

ClerdoGrid opens files from anywhere, so a bug in a file reader matters. Please report security problems privately through the repository's **Security > Report a vulnerability** page rather than in a public issue.

Include the version, your system, and a sample file if there is one.

## What the app does
- The desktop app runs with a sandboxed, isolated page and can only write to files you picked.
- Workbooks never leave your device. There is no account, server or tracking, and formulas can't reach the internet.
- Links in cells open only for `http`, `https` and `mailto`.
- Macros in `.xlsm` files are kept in the file but never run.
- Sheet protection uses the short password hash defined by the file format, so protected sheets work the same in other apps. Protection stops accidental edits; it isn't encryption.
