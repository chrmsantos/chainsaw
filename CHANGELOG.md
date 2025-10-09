# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) and this project now follows Semantic Versioning where practical.

## [1.0.0-simplified-gpl] - 2025-10-07

### Changed

- License switched from Modified Apache 2.0 to GPL-3.0-or-later (see LICENSE). Added SPDX identifiers guidance file.
- README updated with new license badge and licensing change section.
- Monolithic module `chainsaw_0.bas` retains all formatting logic; deprecated subsystems (logging, backups, image/view protection) fully removed again after accidental reintroduction of some defaults.

### Removed

- Backup, image protection, logging default config assignments and stub functions (`BackupAllImages`, `RestoreAllImages`).

### Notes

- Future work: potential re-modularization (orchestrator + formatting + replacements + validation) can be performed without altering formatting semantics.

## [1.0.0-modular] - 2025-10-07

### New / Added

- Modular architecture completed: `ChainsawOrchestrator` pipeline, `modFormatting` as single formatting authority.
- `ARCHITECTURE.md` documentation.
- Smoke test module `modSmokeTests.bas` (`ChainsawSmokeTest`).
- Version constant `CHAINSAW_VERSION` exposed in orchestrator.
- Config flag `AUTO_RUN_SMOKE_TEST` (off by default) to auto-trigger smoke test when `debugMode` is true.

### Changes

- `StandardizeDocumentMain` deprecated and now forwards to `ChainsawRun`.
- All formatting calls explicitly qualified with `modFormatting.*`.
- Duplicate legacy formatting and helper implementations replaced with placeholder comments for traceability.

### Removed Items

- Remaining duplicate helper logic (pattern/numbering, replacements) from legacy modules.

### Migration Notes

- Replace any direct calls to legacy internal formatting helpers with public functions in `modFormatting`.
- To enable auto smoke test: add `AUTO_RUN_SMOKE_TEST=true` under [GERAL] / [GENERAL] in `chainsaw-config.ini` and ensure `DEBUG_MODE=true`.
