# CHAINSAW PROPOSITURAS

## v1.0.1-simplified (2025-10-07)

*An open source VBA solution for standardization and advanced automation of legislative documents in Microsoft Word, developed specifically for Municipal Chambers and institutional environments.*

[![License](https://img.shields.io/badge/License-GPLv3-blue.svg)](LICENSE)
![Word Version](https://img.shields.io/badge/Word-2010+-green.svg)
![Language](https://img.shields.io/badge/Language-VBA-orange.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)

## 📋 Table of Contents

- [Version News](#version-news)
- [Main Features](#-main-features)  
- [Project Structure](#-project-structure)  
- [Installation](#-installation)  
- [Configuration](#%EF%B8%8F-configuration)  
- [Usage](#-usage)  
- [Security](#-security)  
- [Configuration Reference](#%EF%B8%8F-configuration-reference)  
- [Documentation](#-documentation)  
- [Contributing](#-contributing)  
- [License](#-license)

### Version News

Latest (simplified):

- Reverted intentionally to a single monolithic module (`chainsaw_0.bas`) for easier maintenance in Word's VBA editor.
- All non-formatting subsystems removed: logging, backups, image/view protection, timing counters (FINAL PURGE COMPLETE).
- Configuration parser ignores deprecated sections (BACKUP / VISUAL_ELEMENTS / LOGGING).
- Formatting semantics fully preserved (paragraph indentation, fonts, numbering, special tokens, header/footer, hyphenation, replacements).
- Dead code and duplicate stubs pruned; image/view types eliminated.
- Final removal of all residual LogMessage calls and legacy backup references.
- Corrected error handler in first-paragraph formatting routine (previously referenced wrong function name on failure).

### Simplification Rationale

The previous multi-module architecture improved clarity but complicated distribution for users accustomed to a single importable `.bas` file. This edition keeps only the stable legislative formatting pipeline while removing auxiliary concerns (observability, image state capture, backups). Interfaces for those features were collapsed into inert stubs and then eliminated—reducing risk of partial, misleading behavior.

### Performance Notes

Core batching and safe font application routines remain; removed systems had negligible runtime contributions. Document processing speed should match or exceed earlier beta builds.

### Licensing Change

Project license switched to GNU GPL v3 (or later). Each source file may include an SPDX identifier:

`' SPDX-License-Identifier: GPL-3.0-or-later`

See `LICENSE` for the full text. Previous Modified Apache 2.0 terms no longer apply as of this version.

```text
chainsaw/
├── configurations/
│   ├── chainsaw-config.ini      # Runtime configuration (deprecated sections ignored)
│   └── stamp.png                # Header stamp asset
├── installation/                # Installer script & guidance
│   ├── install-chainsaw.ps1
│   └── install-config.ini
├── source/
│   ├── vba-modules/
│   │   └── chainsaw_0.bas       # Monolithic VBA module (all formatting logic)
│   └── testing-props/           # Sample documents
├── README.md
├── LICENSE
└── SECURITY.md
```

### Module Responsibilities

All prior module responsibilities were merged. Key logical regions inside `chainsaw_0.bas` are delimited with comment banners (configuration parsing, validations, formatting routines, replacements, cleanup). Backups, logging, image/view protection banners were removed.

## 🚀 Main Features

- Legislative formatting: standardized fonts, margins, indentation (2nd–4th paragraphs), numbering.
- Semantic paragraph handling: CONSIDERANDO, Justificativa, Anexo detection & formatting.
- Structural cleanup: whitespace normalization, duplicate blank line limiting, hidden element removal.
- Header/footer stamping: optional stamp image + page numbering.
- Hyphenation & replacements: controlled via configuration flags (deprecated sections ignored).
- Defensive guards: safe font application, error-resilient loops.

## 📁 Project Structure

Project intentionally uses a monolith for this simplified line—legacy modular artifacts were retired.

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


[VALIDATIONS]
validate_document_integrity = true
validate_proposition_type = true
check_word_version = true
min_word_version = 14.0
```

### File Locations

- Logs / backups: removed (no writes performed).
- Assets: `assets/` (header image, etc.).
- Word UI customizations: `config/Word Personalizações.exportedUI`.
 
## 📖 Usage

### Basic Usage

1. Open a document in Microsoft Word.
2. Import `chainsaw_0.bas` if not already present (VBA Editor → File → Import File...).
3. Run the macro `Chainsaw`.
4. The system applies all formatting steps sequentially.

### Key Shortcuts

- Alt + F8: Open macro list
- (Optional) Ribbon button mapped to `Chainsaw` macro.

## 🔒 Security

### Macro Configuration in Microsoft Word

To use CHAINSAW PROPOSITURAS safely:

1. **Configurações de Segurança:**
   - Arquivo → Opções → Central de Confiabilidade
   - Configurações de Macro → "Desabilitar todas as macros com notificação"

Checklist:

- ✅ Open and auditable source code
- ✅ No internet connection required
- ✅ No hidden telemetry / logging (all logging system removed)
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

```

Notes:

- Key names are case-insensitive; values: true/false/1/0.
- Portuguese section names also work (e.g., `[INTERFACE]` or `[INTERFACE]`, `[VALIDACOES]`).
- If a key is omitted, its safe default is used.

### Dialog ASCII Normalization

When enabled (`dialog_ascii_normalization = true`), all user-facing dialog strings are converted to an ASCII-safe form (accents replaced, smart quotes normalized) to avoid encoding issues on restricted systems. Set to `false` to retain original accents.

## 📚 Documentation

Project root files (selected):

Historical multi-module breakdown removed; refer to prior tags if needed.

Historical/legacy example or docs folders referenced earlier have been consolidated; examples can be added in a future `examples/` directory as needed.

## 🤝 Contributing

1. Fork o repositório
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

See `CONTRIBUTORS.md` for details on the contribution process.

## 📄 License

This project is licensed under the **GNU General Public License v3.0 or later (GPL-3.0-or-later)** – see [LICENSE](LICENSE) for details.

Note: Microsoft Word is proprietary software and requires its own license.

## 👨‍💻 Author

Christian Martin dos Santos - [chrmsantos](https://github.com/chrmsantos)

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

Current monolithic module ~5,700 lines (after subsystem removals) focused entirely on formatting and cleanup.
