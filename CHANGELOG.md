# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/) and this project now follows Semantic Versioning where practical.

## [1.0.0-Beta3] - 2025-10-10

### Changed

- License switched from Modified Apache 2.0 to GPL-3.0-or-later (see LICENSE). Added SPDX identifiers guidance file.
- README updated to reflect simplified, configuration-free behavior and header image path resolution.
- Monolithic module `chainsaw_0.bas` retains all formatting logic; removed configuration system and fixed defaults in code.
- Use Word alert enums (wdAlertsNone/wdAlertsAll) instead of booleans for DisplayAlerts.
- Standardized key property casing (Text/Size/Alignment/LeftIndent/FirstLineIndent) where applicable.

### Removed

- Configuration parser and all associated flags; legacy ValidateParagraph routine and its dispatch branches.
- Backup, image protection, logging default config assignments and stub functions (`BackupAllImages`, `RestoreAllImages`).

### Fixed

- SafeSetFont and SafeSetParagraphFormat now return True on success; stray MsgBox removed from the latter.
- Header image resolution simplified to a single relative path: `assets\\stamp.png` (resolved from the active document folder only). If not found, the step is skipped.

### Notes

- Future work: potential re-modularization (orchestrator + formatting + replacements + validation) can be performed without altering formatting semantics.
