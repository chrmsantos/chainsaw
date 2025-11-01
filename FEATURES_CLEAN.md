# FEATURES — Chainsaw Proposituras

A concise, categorized list of features implemented or exposed by the Chainsaw VBA module.

## Core

- Single-entry macro: `StandardizeDocumentMain` orchestrates the full standardization pipeline.
- Runtime configuration seeded by `InitializeRuntimeConfigDefaults` (stored in `runtimeConfig`).
- Local-only operation (no network telemetry).

## Formatting & typography

- Apply standard font and font size (defaults: Arial, 12pt).
- Normalize paragraph spacing, collapse redundant blank lines.
- Enforce page setup constants (margins, header/footer distances).
- Normalize document headings and title formatting.

## Paragraph handling

- Detect and format CONSIDERANDO sections.
- Detect and format Justificativa and Anexo headings.
- Normalize numbered paragraph formatting and list numbering.
- Ensure paragraph separation and apply batch processing for throughput.

## Header & footer

- Optional header stamp insertion using `assets/stamp.png` when present.
- Insert page numbers in the footer.

## Replacements & hyphenation

- Apply text replacements and specific paragraph replacements.
- Option to replace hyphens with em dashes.
- Remove manual line breaks where applicable.

## Image & visual elements

- Backup/restore image handling options (configurable).
- Protect images within a specified range.
- Remove hidden visual elements when configured.

## Validation & safety

- Document integrity checks: protection, read-only, minimum content.
- Word version check and compatibility guards.
- Progress reporting (status bar and optional progress bar).
- Undo grouping support to allow rolling back changes.

## Logging & diagnostics

- Configurable logging to `C:\Temp\chainsaw_log.txt` with log level control.
- Rolling log handling and basic rotation logic.

## Configuration & extensibility

- Optional `chainsaw.config` loader that accepts `KEY=VALUE` lines.
- `runtimeConfig` structure exposes many behavior flags for tuning.

## Performance & robustness

- Performance toggles: disable screen updating and display alerts.
- Batch paragraph operations and optimized loops for large documents.
- Safe property access helpers (`SafeGetCharacterCount`, `SafeSetFont`).
- Paragraph cache implementation to reduce repeated object access.

## Recovery & backup

- Emergency recovery helper to restore UI state on error.
- Document backup creation with retry logic before processing.

## Security & privacy

- No external network calls; intended for local, trusted use.
- Keep macros enabled only for Trusted Locations.

## Notes

- Some features are guarded by runtime flags in `runtimeConfig` and can be toggled via an optional config file.
- The module aims for maximum compatibility with Word 2010+ (VBA7) and includes conditional API declarations.

If you want this file moved to a different path or expanded into a formatted table, tell me where and I will update it.
