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

## Configuration

```text

- A `chainsaw.config` file can be placed alongside the document to override runtime defaults. The loader reads simple `KEY=VALUE` lines.chainsaw/

- Runtime options are seeded from `InitializeRuntimeConfigDefaults` in `modChainsaw1.bas`.├── assets/

│   └── stamp.png

## Troubleshooting├── config/

│   ├── Normal.dotm

- If compilation errors mention missing Windows API functions (e.g. `Sleep`), ensure the module is opened in a VBA7-capable Office or the declarations are present near the top of `modChainsaw1.bas`.│   └── Word Personalizações.exportedUI

- Run `Debug > Compile VBAProject` in the VBA editor to get a full list of compile issues. Fix reported missing functions by searching for the referenced line or identifier.├── src/

- If Word reports permission or protection errors, ensure the document is saved and editable (not in Protected View).│   └── chainsaw_0.bas

├── README.md

## Logging├── CONTRIBUTORS.md

├── SECURITY.md

- Logs are written to `C:\Temp\chainsaw_log.txt` by default. Use `ViewLog` to open the log file.├── LICENSE

└── SPDX-LICENSE-IDENTIFIER.txt

## Contributing```



Contributions are welcome. Please follow the project's licensing (GPL-3.0-or-later) and include clear commit messages for any fixes.## Requirements



## License- Windows com Microsoft Word 2010 ou superior.

- Macros habilitadas no Trust Center ("Desabilitar todas as macros com notificação").

This project is licensed under GPL-3.0-or-later. See `LICENSE` for details.- Imagem opcional de cabeçalho em `assets\stamp.png`, armazenada ao lado do documento ativo.


## Setup

1. Baixe ou faça clone do repositório em uma pasta confiável.
2. Abra o documento alvo no Word.
3. Importe `chainsaw_0.bas` pelo Editor VBA (`Alt+F11` → `File` → `Import File`).
4. (Opcional) Importe `config/Word Personalizações.exportedUI` para adicionar atalhos na faixa de opções.

## Usage

1. Pressione `Alt+F8` e execute `StandardizeDocumentMain`.
2. Aguarde a conclusão da rotina; o macro aplica todas as etapas em sequência.
3. Caso `assets\stamp.png` esteja ausente, a etapa de carimbo é ignorada de forma segura.

## Behavior Defaults

- Fonte principal: Arial 12 pt, com margens e espaçamentos definidos pelas constantes do módulo.
- A limpeza de espaçamento remove linhas em branco repetidas e comprime espaços duplicados.
- Rotinas de cabeçalho/rodapé aplicam carimbo e numeração sempre que os recursos estiverem disponíveis.
- Nenhum arquivo de configuração externo é lido nesta edição.

## Security Notes

- O módulo executa localmente sem telemetria ou requisições de rede.
- Armazene o projeto em um Local Confiável e mantenha o modo de notificação antes da execução de macros.
- Consulte `SECURITY.md` para orientações de reporte e política de divulgação.

## Message Templates

Mensagens de interface usam placeholders `{{KEY}}`. `ReplacePlaceholders` substitui pares chave/valor e `NormalizeForUI` garante ASCII quando necessário.

## ASCII Hardening

Constantes de texto permanecem compatíveis com ASCII por padrão: acentos são suavizados, marcadores viram hifens e símbolos ordinais voltam para formas alfabéticas. Quando o ambiente permite, `NormalizeForUI` pode restaurar acentuação antes de mostrar um diálogo.

## Contributing

Consulte `CONTRIBUTORS.md` para conhecer fluxo de contribuição, estilo de código e expectativas de revisão.

## License

Distribuído sob GPL-3.0-or-later. Consulte `LICENSE` para os termos completos.

## Maintainer

Christian Martin dos Santos — <chrmsantos@gmail.com>

## v1.0.0-Beta3 (2025-10-10)

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
- [Usage](#-usage)
- [Security](#-security)
- [Documentation](#-documentation)
- [Contributing](#-contributing)
- [License](#-license)

### Version News

Latest:

- Kept a single monolithic module (`chainsaw_0.bas`) for easy import and maintenance in Word's VBA editor.
- Removed configuration control entirely: behavior now uses fixed, safe defaults (no INI parsing; no optional toggles).
- Header image is now resolved only from a fixed relative path to the active document folder: `assets\stamp.png`.
- Removed the deprecated ValidateParagraph routine and all dispatch branches that referenced it.
- Purged comments referring to removed/deprecated features to reduce noise; standardized key property casing.
- Minor safety fixes in helper functions (return values and alert handling) without changing semantics.
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
├── assets/
│   └── stamp.png                # Header stamp asset used in the document header
├── config/
│   ├── Normal.dotm             # Optional Word template
│   └── Word Personalizações.exportedUI  # Optional ribbon customizations
├── src/
│   └── chainsaw_0.bas          # Monolithic VBA module (all formatting logic)
├── README.md
├── CHANGELOG.md
├── LICENSE
└── SPDX-LICENSE-IDENTIFIER.txt
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

No runtime configuration is required (or loaded). This simplified build runs with fixed, safe defaults:

- Minimum Word version: 2010+
- Standard font: Arial 12 pt; line spacing and margins as per module constants
- Header image: resolved from `assets\stamp.png` relative to the active document’s folder
- Page numbers: added to the footer automatically

If `assets\stamp.png` is not found next to the document, the header image step is skipped safely.

## 📖 Usage

### Basic Usage

1. Open a document in Microsoft Word.
2. Import `chainsaw_0.bas` if not already present (VBA Editor → File → Import File...).
3. Run the macro `StandardizeDocumentMain`.
4. The system applies all formatting steps sequentially.

Note on header stamp:

- Place an image at `assets\stamp.png` in the same folder as your .docx. If it's missing, the header image step is skipped automatically.

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

## 📝 Notes on Dialog Text

User-facing dialog strings are normalized to ASCII at runtime via a helper to avoid encoding issues on older Word builds. This does not affect document content.

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

See [CONTRIBUTORS.md](CONTRIBUTORS.md) for details on the contribution process.

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
msg = ReplacePlaceholders(MSG_ERR_VERSION, "MIN", CStr(14), "CUR", Application.Version)
MsgBox NormalizeForUI(msg), vbCritical, NormalizeForUI(TITLE_VERSION_ERROR)
```

Why double braces? They avoid conflicts with legacy %PLACEHOLDER% tokens that caused a compilation issue and are visually distinct from regular percent symbols sometimes present in legislative text.

All new dynamic dialogs should prefer ReplacePlaceholders over manual Replace() chains for maintainability.

## 📏 Code Size Metrics

Current monolithic module ~5,700 lines (after subsystem removals) focused entirely on formatting and cleanup.
