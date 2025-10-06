# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

### Added (Unreleased)

- Finalized orchestrator rename: `modMain.bas` -> `modPipeline.bas`.

### Changed (Unreleased)

- `chainsaw.bas` remains a pure entry stub; pipeline logic consolidated in `modPipeline.bas`.
- `ARCHITECTURE.md` updated to reflect orchestrator rename.
- `README.md` module table updated (modPipeline supersedes legacy modMain).

### Removed (Unreleased)

- Accidental pasted legacy procedures from the former monolith out of `chainsaw.bas`.

### Notes (Unreleased)

- Next steps: migrate remaining view/backup legacy blocks or formally deprecate.

## [1.0.0-beta1] - 2025-10-06

### Added

- Modular architecture: extracted formatting, replacements, validation, safety, config, and (stub) logging modules.
- Centralized formatting routines in `modFormatting` (migrated special section handlers: Considerando, Justificativa, Anexo patterns).
- ASCII normalization option for user dialog messages.
- `ARCHITECTURE.md` documentation file.

### Changed

- Monolithic `chainsaw.bas` reduced to orchestrator responsibilities.
- Logging system replaced with no-op stubs (`modLog`) pending future reinstatement.
- Backup configuration flags marked deprecated and disabled by default.
- Default configuration rewritten for clarity; backup/logging now inert.
- README overhauled to reflect new module structure and disabled features.
- SECURITY policy updated (removal of active logging/backups, added threat model).

### Removed

- Duplicate legacy orchestrator module (previous `modMain.bas`).
- Active logging & backup runtime behavior (retained keys for compatibility).

### Deprecated

- Backup-related config flags: `autoBackup`, `backupBeforeProcessing`, `maxBackupFiles`, `backupCleanup`, `backupRetryAttempts`, `enableEmergencyBackup` (inactive).
- Logging-related config flags: `enableLogging`, `logLevel`, `logToFile`, `maxLogSizeMb` (stub only).

### Security

- Reduced attack surface by disabling file write features (logs/backups).

### Notes

- Formatting semantics intentionally unchanged per original simplification goal.
- Future beta will consider reinstating logging with structured, size-limited output and opt-in backups with retention.

---
Format: Keep chronological (newest on top after first release). Use Keep a Changelog style guidelines.
