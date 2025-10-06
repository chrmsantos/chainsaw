# CHAINSAW PROPOSITURAS

## v1.0.0-Beta2

*An open source VBA solution for standardization and advanced automation of legislative documents in Microsoft Word, developed specifically for Municipal Chambers and institutional environments.*

[![License](https://img.shields.io/badge/License-Apache%202.0%20Modified-blue.svg)](LICENSE)
![Word Version](https://img.shields.io/badge/Word-2010+-green.svg)
![Language](https://img.shields.io/badge/Language-VBA-orange.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)

## 📋 Table of Contents

- [Version News](#-version-news-100-beta2)
- [Main Features](#-main-features)
- [Project Structure](#-project-structure)
- [Installation](#-installation)
- [Configuration](#%EF%B8%8F-configuration)
- [Usage](#-usage)
- [Security](#-security)
- [Requirements](#-requirements)
- [Configuration Reference](#%EF%B8%8F-configuration-reference)
- [Architecture Overview](#-architecture-overview)
- [Troubleshooting](#-troubleshooting)
- [Roadmap](#-roadmap-planned)
- [Recent Refactor Summary](#-recent-refactor-summary)
- [Documentation](#-documentation)
- [Contributing](#-contributing)
- [License](#-license)

### Refactored Architecture (Beta2 Consolidation)

**Beta2 highlights:**

- `chainsaw.bas` is now a pure stub (only `ChainsawProcess`).
- Canonical pipeline lives in `modPipeline.RunChainsawPipeline`.
- Legacy `modMain.bas` removed (pipeline fully migrated).
- All formatting / spacing / numbering / separation logic centralized in `modFormatting`.
- Logging & backups remain disabled (stubs retained for compatibility; zero file writes).
- Structural validation placeholder isolated for future enhancement.
- Self-test harness preserved to detect regressions.
- Clear segregation of constants (`modConstants`) and messages (`modMessages`).

### Performance Optimizations

- **Batch processing:** Paragraphs processed in groups for better performance
- **Optimized operations:** Bulk Find/Replace, caching of frequently used objects
- **Memory management:** Smart garbage collection and minimal object creation
- **Compatibility preserved:** All optimizations keep compatibility with Word 2010+

### Enhanced Logging System

#### Logging system (stubbed in Beta2)

<!-- Retained descriptive bullets for future reinstatement -->
<!-- Detailed control: Configure log levels (ERROR, WARNING, INFO, DEBUG) -->
<!-- Performance tracking: Accurate execution time measurement -->
> Project structure shown below was updated after removal of the transitional `modMain.bas`. There is **no** `installation/` folder in this revision; former installer assets were consolidated or deferred.

```text
chainsaw/
├── assets/                      # Images and other assets (e.g. stamp.png)
├── config/                      # Normal.dotm + Ribbon/QAT export
│   ├── Normal.dotm
│   └── Word Personalizações.exportedUI
├── scripts/                     # Utility scripts (e.g. count-loc.ps1)
├── src/                         # VBA source modules
│   ├── chainsaw.bas             # Public entry stub (ChainsawProcess)
│   ├── modPipeline.bas          # Orchestrator pipeline
│   ├── modFormatting.bas        # Formatting & layout routines
│   ├── modReplacements.bas      # Text & semantic replacements
│   ├── modValidation.bas        # Structural / lexical checks
│   ├── modSafety.bas            # Defensive Word object wrappers
│   ├── modConfig.bas            # Configuration loading & defaults
│   ├── modMessages.bas          # User-facing strings
│   ├── modConstants.bas         # Stable constants (fonts, version)
│   ├── modErrors.bas            # Error/status reporting (no I/O)
│   ├── modSelfTest.bas          # Regression/self-test macro
│   ├── modUI.bas                # (Placeholder) UI helpers
│   └── modLog.bas               # No-op logging stubs
├── legacy_chainsaw_snapshot.bas # Archived pre-truncation monolith
├── ARCHITECTURE.md              # Detailed architecture notes
├── CHANGELOG.md                 # Changelog (Keep a Changelog style)
├── LICENSE                      # License
├── README.md                    # This file
└── SECURITY.md                  # Security policy
```

### Module Responsibilities

| Module | Responsibility | Example Procedure |
|--------|----------------|-------------------|
| chainsaw.bas | Public entry macro (stub only) | `ChainsawProcess` |
| modPipeline | Canonical pipeline orchestrator | `RunChainsawPipeline` |
| modFormatting | Formatting & special paragraphs | `FormatConsiderandoParagraphs` |
| modReplacements | Pattern / semantic replacements | `ApplyTextReplacements` |
| modValidation | Content & lexical validation | `ValidateContentConsistency` |
| modSafety | Safe wrappers for Word API | `SafeHasVisualContent` |
| modConfig | Config parsing & defaults | `modConfig_LoadConfiguration` |
| modLog | Stubbed logging API | `LogStepStart` |
| modErrors | Centralized error/status reporting (no I/O) | `ReportUnexpected` |
| modSelfTest | Lightweight regression/self-test macro | `ChainsawSelfTest` |
 
- **Flexible configuration:** Enable/disable logging by category

## 🚀 Main Features

- **Automatic standardization of legislative propositions:** Specific formatting for INDICAÇÕES, REQUERIMENTOS and MOÇÕES with institutional layout control.
- **Configurable content validation:** Consistency checks between header and content (can be disabled).
- **Smart cleanup of visual elements:** Automatic removal of hidden and inappropriate formatting (fully configurable).
- **(Planned) backup system:** Disabled this beta; keys retained.
- **Institutional formatting:** Header with logo, page numbering and standardized margins.
- **(Planned) structured logging:** Stub only (no file output this beta).
- **Enhanced interface:** Clear user messages and interactive validations.
- **Optimized performance:** Consolidated passes minimize duplication.
- **Security:** Integrity validation hooks & reduced file I/O surface.
- **Self-test macro:** `ChainsawSelfTest` collects metrics (paragraphs, words, chars, images) to detect unintended changes.

## 📁 Project Structure

See the updated structure in the earlier section (avoids duplication). `legacy_chainsaw_snapshot.bas` is retained only for audit/history and is **not** imported.

## 🔧 Installation

### Quick Install (Recommended)

1. Download the project (or copy the files to a trusted folder).
2. Import the required `.bas` modules into Word’s VBA editor (ALT+F11 → File > Import File...).
3. (Optional) Import ribbon customizations from `config/Word Personalizações.exportedUI`.

### Manual Installation

Manual steps depend on your Word setup. If you need an installer, we can add one later in `scripts/`.

## ⚙️ Configuration

The system loads settings from `chainsaw-config.ini` (placed alongside the document or in the expected configuration path). If the file is missing, safe defaults are applied.

### Quick Configuration

```ini
[GENERAL]
debug_mode = false
performance_mode = true
compatibility_mode = true

[VALIDATIONS]
validate_document_integrity = true
validate_proposition_type = true
check_word_version = true
min_word_version = 14.0
```

### File Locations

- Logs and backups: (inactive this beta) will reside beside the document when re-enabled.
- Assets: `assets/` (header image, etc.).
- Word UI customizations: `config/Word Personalizações.exportedUI`.
 
## 📖 Usage

### Basic Usage

1. Open a document in Microsoft Word.
2. Run the macro `ChainsawProcess` (or current orchestrator name in `chainsaw.bas`).
3. The system will process the document according to configuration.

### Key Shortcuts

- Alt + F8: Open macro list
- Ctrl + Shift + P: Custom shortcut (configurable)

## 🔒 Security

### Macro Configuration in Microsoft Word

To use CHAINSAW PROPOSITURAS safely:

1. **Configurações de Segurança:**
   - Arquivo → Opções → Central de Confiabilidade
   - Configurações de Macro → "Desabilitar todas as macros com notificação"

Checklist:

- ✅ Open and auditable source code
- ✅ No internet connection required
- ✅ Backup subsystem planned (disabled in this beta)
- ✅ Robust error handling

Para políticas corporativas, consulte [`SECURITY.md`](SECURITY.md).


- OS: Windows 7 or later
- Microsoft Word: 2010 or later
- Permissions: VBA macro execution enabled
- Disk Space: 50MB free

### Recommended

- Microsoft Word: 2016 or later
- RAM: 4GB or higher
- CPU: Intel/AMD 64-bit

## ⚙️ Configuration Reference

Below are selected, stable keys you can place in `chainsaw-config.ini` (section names accept Portuguese or English equivalents):

```ini
[INTERFACE]
dialog_ascii_normalization = true    ; true/false — fold accents & special chars in MsgBox text

[VALIDATIONS]
check_word_version = true            ; disable only for legacy environments
validate_proposition_type = true

[GENERAL]
debug_mode = false
performance_mode = true
compatibility_mode = true
```

Notes:

- Key names are case-insensitive; values: true/false/1/0.
- Portuguese section names also work (e.g., `[INTERFACE]` or `[INTERFACE]`, `[VALIDACOES]`).
- If a key is omitted, its safe default is used.

### Dialog ASCII Normalization

When enabled (`dialog_ascii_normalization = true`), all user-facing dialog strings are converted to an ASCII-safe form (accents replaced, smart quotes normalized) to avoid encoding issues on restricted systems. Set to `false` to retain original accents.

## 📚 Documentation

Project root files (selected):

- `modSelfTest.bas` – Macro `ChainsawSelfTest` (regression sanity)
- `modErrors.bas` – Minimal status/error centralization (no file writes in beta)
- `CONTRIBUTORS.md` – Contributors list
- `legacy_chainsaw_snapshot.bas` – Archived pre-refactor monolith (not loaded)
- `scripts/count-loc.ps1` – Utility to count active vs legacy LOC

Historical/legacy example or docs folders referenced earlier have been consolidated; examples can be added in a future `examples/` directory as needed.

## 🤝 Contributing

1. Fork o repositório
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

See `CONTRIBUTORS.md` for details on the contribution process.

## 📄 License

This project is licensed under the **Apache 2.0 Modified License (with clause 10)** - see [LICENSE](LICENSE) for details.

Note: Microsoft Word is proprietary software and requires its own license.

## 👨‍💻 Author

Christian Martin dos Santos - [chrmsantos](https://github.com/chrmsantos)

---

---

Built with ❤️ for the legislative community

## 🧩 Message Templating System

Dynamic user-facing messages use a lightweight placeholder system to avoid repetitive string concatenation and to simplify localization.

Placeholders format:

  {{KEY}}

Examples:

```vb
MSG_ERR_VERSION = "This tool requires Microsoft Word {{MIN}} or higher." & vbCrLf & _
                  "Current version: {{CUR}}" & vbCrLf & _
                  "Minimum version: {{MIN}}"
```

Helpers:

- ReplacePlaceholders(template, "KEY1", value1, "KEY2", value2, ...)
  Replaces each {{KEY}} with its corresponding value (converted to string). Odd trailing key without a value is ignored safely.

### ASCII Hardening of Dialog Text

Some environments (older Word builds / locale mismatches) raised compilation or rendering issues with certain Unicode characters (accented capitals, bullets •, ordinal indicators º). To guarantee reliability of the exported `.bas` module we applied an explicit ASCII hardening to several Portuguese messages:

- Accented letters were flattened (INDICAÇÃO → INDICACAO, MOÇÃO → MOCAO, ATENÇÃO → ATENCAO, CONSISTÊNCIA → CONSISTENCIA, etc.)
- Bullets (•) replaced with hyphens (-)
- Ordinal indicator º replaced with 'o'

Runtime readability is still acceptable; if future builds require restoring original accents, two approaches are possible:

1. Reintroduce accented literals directly in the constants (if your environment accepts them) and rely on `NormalizeForUI` to fold when `dialog_ascii_normalization = true`.
2. Maintain ASCII in constants and add a small helper that maps specific hardened words back to accented display forms right before `MsgBox`.

Given current goals (robust compilation across Word 2010+ and mixed encodings), we kept the source ASCII-safe by default. Open an issue if you want an optional accent-restoration layer added.

Usage example inside code:

```vb
Dim msg As String
msg = ReplacePlaceholders(MSG_ERR_VERSION, "MIN", Config.minWordVersion, "CUR", Application.Version)
MsgBox NormalizeForUI(msg), vbCritical, NormalizeForUI(TITLE_VERSION_ERROR)
```

Why double braces? They avoid conflicts with legacy %PLACEHOLDER% tokens that caused a compilation issue and are visually distinct from regular percent symbols sometimes present in legislative text.

All new dynamic dialogs should prefer ReplacePlaceholders over manual Replace() chains for maintainability.

## 📏 Code Size Metrics

Active VBA source (excluding legacy snapshot) currently totals approximately **1,826** lines across 13 active modules. The archived legacy snapshot plus removed transitional module previously exceeded 4,000 lines—illustrating the reduction and clearer separation of concerns.

To recompute metrics locally:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/count-loc.ps1
```

Include legacy snapshot lines as well:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/count-loc.ps1 -IncludeLegacy
```

These scripts are informational only; they perform no modifications.
