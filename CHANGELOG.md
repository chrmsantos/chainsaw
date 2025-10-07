# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Planned

* Structural validation implementation (currently placeholder).
* Optional lightweight logging/telemetry module (opt-in).
* Accent restoration layer (configurable) for dialogs.

## [1.0.0-beta3] - 2025-10-07

### Added (beta3)

* Legacy parity routine `FormatThirdAndFourthParagraphs` ensuring mandated indenting for 3rd & 4th substantive paragraphs.

### Changed (beta3)

* Documentation (README, ARCHITECTURE, SECURITY) aligned with simplified active module set.
* Reduced configuration surface (removed obsolete logging & compatibility keys).

### Removed (beta3)

* Legacy monolith (`chainsaw_old.bas`), backward stub (`chainsaw.bas`), logging stubs (`modLog.bas`), legacy snapshot references.

### Fixed (beta3)

* Purged stale references to removed modules/features.

### Notes (beta3)

* Formatting semantics preserved; structural and documentation simplification only.

## [1.0.0-beta2] - 2025-10-06

### Added (beta2)

* Final canonical orchestrator `modPipeline.RunChainsawPipeline` (logic fully inlined, legacy wrapper removed).
* Real structural validation hook placeholder separated for future enhancement.

### Changed (beta2)

* `chainsaw.bas` now a pure entry stub (private helpers removed; archived in `legacy_chainsaw_snapshot.bas`).
* `ARCHITECTURE.md` and `README.md` reflect removal of transitional `modMain.bas`.
* Consolidated duplicate formatting routines (hyphenation, watermark, header/footer, second paragraph helpers) into single implementations.

### Removed (beta2)

* Redundant duplicate formatting function definitions in `modFormatting.bas`.
* Deprecated `modMain.bas` physically removed (pipeline fully in `modPipeline`).

### Fixed (beta2)

* Eliminated residual duplication that could cause ambiguous references during future maintenance.

### Notes (beta2)

* Formatting semantics unchanged; refactor strictly architectural. Pre-truncation content preserved in `legacy_chainsaw_snapshot.bas`.

## [1.0.0-beta1] - 2025-10-06

### Added (beta1)

* Modular architecture: extracted formatting, replacements, validation, safety, config, and (stub) logging modules.
* Centralized formatting routines in `modFormatting` (migrated special section handlers: Considerando, Justificativa, Anexo patterns).
* ASCII normalization option for user dialog messages.
* `ARCHITECTURE.md` documentation file.

### Changed (beta1)

* Monolithic `chainsaw.bas` reduced to orchestrator responsibilities.
* Logging system replaced with no-op stubs (`modLog`) pending future reinstatement.
* Backup configuration flags marked deprecated and disabled by default.
* Default configuration rewritten for clarity; backup/logging now inert.
* README overhauled to reflect new module structure and disabled features.
* SECURITY policy updated (removal of active logging/backups, added threat model).

### Removed (beta1)

* Duplicate legacy orchestrator module (previous `modMain.bas`).
* Active logging & backup runtime behavior (retained keys for compatibility).

### Deprecated (beta1)

* Backup-related config flags: `autoBackup`, `backupBeforeProcessing`, `maxBackupFiles`, `backupCleanup`, `backupRetryAttempts`, `enableEmergencyBackup` (inactive).
* Logging-related config flags: `enableLogging`, `logLevel`, `logToFile`, `maxLogSizeMb` (stub only).

### Security (beta1)

* Reduced attack surface by disabling file write features (logs/backups).

### Notes (beta1)

* Formatting semantics intentionally unchanged per original simplification goal.
* Future beta will consider reinstating logging with structured, size-limited output and opt-in backups with retention.

---
Format: Keep chronological (newest on top after first release). Use Keep a Changelog style guidelines.
