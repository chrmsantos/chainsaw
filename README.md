# Chainsaw Proposituras

Chainsaw Proposituras is a single-file Microsoft Word VBA module that standardizes formatting and structure of municipal legislative documents ("proposituras"). It performs typographic normalization, paragraph and header/footer fixes, light validation, and optional logging. The macro runs locally inside Word and requires no external services.

## Quick start

1. Open the Word document you want to process.
2. Press Alt+F11 to open the VBA editor.
3. Import `src/src/modChainsaw1.bas` (VBA Editor → File → Import File...).
4. Run the macro `StandardizeDocumentMain` (Alt+F8).

## Requirements

- Microsoft Word 2010 or newer (VBA7-compatible).
- Macros enabled and permission to run VBA macros.

## Optional configuration

Create an optional `chainsaw.config` file next to the document to override a few runtime defaults. Use simple `KEY=VALUE` lines. Unrecognized keys are ignored.

Example `chainsaw.config`:

```text
EnableLogging=true
EnableProgressBar=true
MaxSessionStampWords=17
```

## Files of interest

- `src/src/modChainsaw1.bas` — main macro module (import into Word).
- `assets/stamp.png` — optional header stamp image (placed next to the document to enable header stamping).

## Troubleshooting

- Compile errors: in the VBA editor run `Debug > Compile VBAProject` to list issues and exact line numbers.
- Missing `Sleep` or API-related errors: the module declares `Sleep` conditionally for VBA7; confirm the declaration at the top of `modChainsaw1.bas`.
- Protected/Read-only documents: save a writable copy before running the macro.

## Logging

Default log file: `C:\Temp\chainsaw_log.txt`. Use the `ViewLog` macro to open it.

## Security & privacy

- The macro runs locally without network calls or telemetry.
- Only enable macros from trusted locations and follow your organization's macro policies.

## Contributing & license

Contributions are welcome. See `CONTRIBUTORS.md` for contribution guidance. This project is licensed under GPL-3.0-or-later; see `LICENSE` for details.

Maintainer: Christian Martin dos Santos — <chrmsantos@gmail.com>
