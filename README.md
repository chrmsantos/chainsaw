# CHAINSAW PROPOSITURAS

## v1.0.0-Beta1

*An open source VBA solution for standardization and advanced automation of legislative documents in Microsoft Word, developed specifically for Municipal Chambers and institutional environments.*

[![License](https://img.shields.io/badge/License-Apache%202.0%20Modified-blue.svg)](LICENSE)
![Word Version](https://img.shields.io/badge/Word-2010+-green.svg)
![Language](https://img.shields.io/badge/Language-VBA-orange.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)

## 📋 Table of Contents

- [Version News](#-version-news-100-beta1)
- [Main Features](#-main-features)
- [Project Structure](#-project-structure)
- [Installation](#-installation)
- [Configuration](#️-configuration)
- [Usage](#-usage)
- [Security](#-security)
- [Requirements](#-requirements)
- [Configuration Reference](#-configuration-reference)
- [Contributing](#-contributing)
- [License](#-license)

## 🆕 Version News 1.0.0-Beta1

### Advanced Configuration System

- **External configuration file:** `chainsaw-config.ini` with over 100 settings
- **Granular control:** Enable/disable any system feature
- **15 configuration categories:** General, Validations, Backup, Formatting, Cleanup, Performance, etc.
- **Automatic configuration:** Loads default values if file not found

### Performance Optimizations

- **Batch processing:** Paragraphs processed in groups for better performance
- **Optimized operations:** Bulk Find/Replace, caching of frequently used objects
- **Memory management:** Smart garbage collection and minimal object creation
- **Compatibility preserved:** All optimizations keep compatibility with Word 2010+

### Enhanced Logging System

- **Detailed control:** Configure log levels (ERROR, WARNING, INFO, DEBUG)
- **Performance tracking:** Accurate execution time measurement
- **Flexible configuration:** Enable/disable logging by category

## 🚀 Main Features

- **Automatic standardization of legislative propositions:**
  Specific formatting for INDICAÇÕES, REQUERIMENTOS and MOÇÕES with institutional layout control.
- **Configurable content validation:**
  Consistency checks between header and content (can be disabled).
- **Smart cleanup of visual elements:**
  Automatic removal of hidden and inappropriate formatting (fully configurable).
- **Robust backup system:**
  Automatic backup before modifications, with emergency recovery.
- **Institutional formatting:**
  Header with logo, page numbering and standardized margins.
- **Detailed logging:**
  Logs with timestamps, severity levels and full traceability.
- **Enhanced interface:**
  Clear user messages and interactive validations.
- **Optimized performance:**
  Efficient processing even for large documents.
- **Advanced security:**
  Integrity validation, version check and failure protection.

## 📁 Project Structure

```text
chainsaw/
├── assets/                    # Assets (images, icons)
│   └── stamp.png              # Header/logo image
├── config/                    # Configuration and Word UI customizations
│   ├── Normal.dotm            # Word Normal template (customized)
│   └── Word Personalizações.exportedUI  # Ribbon/QAT export
├── scripts/                   # Scripts (currently empty)
├── src/                       # Source code
│   └── chainsaw.bas           # Main VBA module
├── LICENSE                    # Project license
├── README.md                  # This file
└── SECURITY.md                # Security policy
```

## 🔧 Installation

### Quick Install (Recommended)

1. Download the project (or copy the files to a trusted folder).
2. Import `src/chainsaw.bas` into Word’s VBA editor (ALT+F11).
3. Optionally import `config/Word Personalizações.exportedUI` into Word to add ribbon/QAT buttons.

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

As configuration evolves, this section will be expanded. See the Configuration Reference below for newly added flags.

### File Locations

- Logs and backups: same folder as the current document, or TEMP if unsaved.
- Assets: `assets/` (header image, etc.).
- Word UI customizations: `config/Word Personalizações.exportedUI`.

## 📖 Usage

### Basic Usage

1. Open a document in Microsoft Word
2. Run the macro `StandardizeDocumentMain`
3. The system will automatically process the document according to the configuration

### Key Shortcuts

- Alt + F8: Open macro list
- Ctrl + Shift + P: Custom shortcut (configurable)

## 🔒 Security

### Macro Configuration in Microsoft Word

To use CHAINSAW PROPOSITURAS safely:

1. **Configurações de Segurança:**
   - Arquivo → Opções → Central de Confiabilidade
   - Configurações de Macro → "Desabilitar todas as macros com notificação"

2. **Security Checks:**
  - ✅ Open and auditable source code
  - ✅ No internet connection required
  - ✅ Automatic backup before modifications
  - ✅ Robust error handling

Para políticas corporativas, consulte [`SECURITY.md`](SECURITY.md).

## 📋 Requirements

### Minimum

- OS: Windows 7 or later
- Microsoft Word: 2010 or later
- Permissions: VBA macro execution enabled
- Disk Space: 50MB free

### Recommended

- Microsoft Word: 2016 or later
- RAM: 4GB or higher
- CPU: Intel/AMD 64-bit

## � Configuration Reference

Below are selected, stable keys you can place in `chainsaw-config.ini` (section names accept Portuguese or English equivalents):

```ini
[INTERFACE]
dialog_ascii_normalization = true    ; true/false — fold accents & special chars in MsgBox text

[VALIDATIONS]
check_word_version = true            ; disable only for legacy environments
validate_document_integrity = true
validate_proposition_type = true
validate_content_consistency = true

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

Project root files:

- `SECURITY.md` – Security policies
- `CONTRIBUTORS.md` – Contributors list
- `installation/INSTALL.md` – Detailed installation & deployment guide

Historical/legacy example or docs folders referenced earlier have been consolidated; examples can be added in a future `examples/` directory as needed.

## 🤝 Contributing

Contributions are welcome! To contribute:

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
