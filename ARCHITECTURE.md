# Chainsaw Architecture

## Overview

Chainsaw is a modular VBA toolkit for standardizing Microsoft Word documents related to legislative propositions. The legacy monolithic macro has been decomposed into clear, sandboxed modules with a thin orchestrator entry point.

## Modules

- **ChainsawOrchestrator.bas**: Public macro entry points and pipeline coordination.
- **modConfig.bas**: Loads and stores configuration (`Config` record) from INI or defaults.
- **modValidation.bas**: Document/environment validation (Word version, integrity, proposition type, content consistency).
- **modPerformance.bas**: Performance toggles (screen updating, event disabling, etc.).
- **modFormatting.bas**: Single authoritative implementation of all formatting, replacements, paragraph transformations, numbering normalization, header/footer insertion, hyphenation, and watermark removal.
- **modUtils.bas**: Shared utility helpers (timing, placeholder substitution, normalization helpers).
- **modOrquestration.bas**: Transitional legacy container. Now reduced to constants, some validation helpers and deprecated wrappers (`StandardizeDocumentMain`). Duplicate formatting code bodies replaced with placeholder comments referencing `modFormatting`.
- **chainsaw_0.bas**: Legacy monolith shell retained temporarily for backwards compatibility and audit trail (contains only placeholder comments for removed routines).

## Execution Flow

1. User runs `ChainsawRun` (or alias `Chainsaw`).
2. Orchestrator loads configuration (unless suppressed), performs validations, initializes performance optimizations.
3. Based on configuration flags, orchestrator delegates each formatting step to the corresponding public function in `modFormatting`.
4. Performance settings optionally restored (if performance toggle module implements restoration).
5. Status surfaced via `Application.StatusBar`.

## Key Design Choices

- **Single Source of Truth**: All formatting logic resides in `modFormatting`; legacy duplicates removed.
- **Explicit Qualification**: Orchestrator calls `modFormatting.*` to avoid accidental reference to deprecated stubs.
- **Backward Compatibility**: `StandardizeDocumentMain` preserved as a thin forwarder to `ChainsawRun`.
- **Private Helpers**: Internal pattern detection and numbering helpers are Private inside `modFormatting` to reduce public surface.
- **Fail-Soft Philosophy**: Non-critical failures (e.g., hyphenation not supported) do not abort the pipeline.

## Public API Surface

Primary user-facing macros:

- `ChainsawRun`
- `Chainsaw` (alias)
- `ChainsawLoadConfiguration`
- `ChainsawValidateActiveDocument`
- `ChainsawFormatActiveDocument`
- `StandardizeDocumentMain` (deprecated)

Primary callable formatting functions (if advanced users need granular control):

- `ApplyPageSetup`
- `ApplyStdFont`
- `ApplyStdParagraphs`
- `FormatFirstParagraph`
- `FormatSecondParagraph`
- `FormatConsiderandoParagraphs`
- `FormatJustificativaAnexoParagraphs`
- `FormatNumberedParagraphs`
- `EnableHyphenation`
- `RemoveWatermark`
- `InsertHeaderstamp`
- `InsertFooterstamp`
- `ApplyTextReplacements`
- `ApplySpecificParagraphReplacements`

(Consider exposing a single high-level `ApplyAllFormatting` wrapper later.)

## Deprecation Policy

- Legacy bodies replaced by comments retain historical trace until final purge milestone.
- A future cleanup pass may remove placeholder regions once stability confirmed.

## Future Enhancements

- Add automated smoke tests (macro) to count applied formatting transformations.
- Introduce granular logging toggled via configuration (lightweight, not the removed heavy logging system).
- Provide unit-test-like harness via a separate template document.
- Investigate isolating configuration parsing into its own class for easier future extension.

---
Generated on: 2025-10-07
