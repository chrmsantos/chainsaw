# Chainsaw Proposituras# Chainsaw Proposituras



Chainsaw Proposituras is a Microsoft Word VBA macro collection to standardize and format legislative documents (proposituras). It provides automated formatting, validation, logging, and safe operations optimized for Word 2010+.## Overview



## Quick startChainsaw Proposituras disponibiliza um único módulo VBA que padroniza documentos legislativos municipais no Microsoft Word. O macro garante tipografia, espaçamento e numeração consistentes sem depender de serviços externos ou arquivos de configuração em tempo de execução.



1. Open the Word document you want to standardize.## Feature Highlights

2. Press Alt+F11 to open the VBA editor and import (or open) the `modChainsaw1.bas` module if not already present.

3. From the VBA editor, run the `StandardizeDocumentMain` subroutine.- Valida a estrutura do documento antes de aplicar qualquer transformação.

- Normaliza títulos, CONSIDERANDO, justificativas, anexos e parágrafos numerados.

## Requirements- Limpa espaços em branco redundantes e restabelece separação segura entre parágrafos.

- Injeta carimbo de cabeçalho a partir de `assets\stamp.png` quando disponível e garante numeração de páginas no rodapé.

- Microsoft Word 2010 or newer (module contains VBA7-compatible API declarations).- Mantém diálogos com o usuário seguros em ASCII por meio de helpers compartilhados.

- Macros enabled and access to the VBA project if needed.

## Repository Layout

<!--
  Fresh README for Chainsaw Proposituras
  Minimal, actionable, and intentionally concise.
-->

# Chainsaw Proposituras

A focused VBA macro for Microsoft Word that standardizes formatting and structure of municipal legislative documents ("proposituras").

<!--
  Fresh README for Chainsaw Proposituras
  Minimal, actionable, and intentionally concise.
-->

# Chainsaw Proposituras

A focused VBA macro for Microsoft Word that standardizes formatting and structure of municipal legislative documents ("proposituras").

Key points:

- Single-file VBA module: `src/src/modChainsaw1.bas` (import into Word's VBA editor).
- Safe defaults: behavior is initialized by `InitializeRuntimeConfigDefaults` and works out-of-the-box.
- Runs locally with no external network calls or telemetry.

---

## Quick start

1. Open the Microsoft Word document you want to process.
2. Press `Alt+F11` to open the VBA editor.
3. Import the module: `File → Import File...` → select `src/src/modChainsaw1.bas`.
4. Press `Alt+F8`, choose `StandardizeDocumentMain` and run it.

## Minimal requirements

- Microsoft Word 2010 or newer (supports VBA7). 64-bit and 32-bit Office are supported via conditional API declarations.
- Macros enabled in Word and permission to run VBA macros.

## Optional configuration

Place a `chainsaw.config` file next to the document to override a few runtime defaults. Format: simple `KEY=VALUE` lines. Unrecognized keys are ignored.

Example `chainsaw.config` (optional):

```text
EnableLogging=true
EnableProgressBar=true
MaxSessionStampWords=17
```

## Files of interest

- `src/src/modChainsaw1.bas` — main macro module (import into Word).
- `assets/stamp.png` — optional header stamp image (if present, Chainsaw can insert it into the document header).

## Troubleshooting

- Compilation errors: in the VBA editor run `Debug > Compile VBAProject` to get the list. Re-import `modChainsaw1.bas` if the module looks corrupted.
- Missing `Sleep` or API errors: the module declares `Sleep` conditionally for VBA7; ensure the declaration at the top of the file is present.
- Protected/Read-only documents: save the file and make it editable before running the macro.

## Logging

Default log file: `C:\Temp\chainsaw_log.txt`. Use the `ViewLog` macro to open the log.

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
EnableProgressBar=true

