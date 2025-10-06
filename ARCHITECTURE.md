# Architecture Overview

## Purpose

This document explains the modular structure introduced in the 1.0.0 Beta refactor, replacing the earlier monolithic `chainsaw.bas` file. Formatting semantics were preserved exactly while reducing complexity and isolating responsibilities.

## Module Map

| Module | Responsibility | Key Public Procedures (Illustrative) |
|--------|----------------|--------------------------------------|
| chainsaw.bas | Entry point stub (public macro only) | ChainsawProcess |
| modPipeline.bas | Pipeline orchestrator (successor of transitional modMain) | RunChainsawPipeline |
| modFormatting.bas | All document + paragraph formatting rules | FormatFirstParagraph, FormatConsiderandoParagraphs |
| modReplacements.bas | Pattern-based global & targeted replacements | ApplyTextReplacements |
| modValidation.bas | Content & lexical integrity checks | ValidateContentConsistency |
| modSafety.bas | Defensive wrappers around Word object model | SafeHasVisualContent |
| modConfig.bas | INI config load + defaults | modConfig_LoadConfiguration |
| modLog.bas | No-op logging stubs | LogStepStart |

## Processing Pipeline (Conceptual)

1. Load configuration (INI or defaults)
2. Safety & validation checks (Word version, integrity, proposition type, optional content consistency)
3. Formatting phase:
   - Page setup, base font, paragraph normalization
   - Special sections (first/second paragraph, CONSIDERANDO, Justificativa, Anexo)
   - Numbering, hyphenation, watermark removal, header/footer
4. Replacement phase (global + specific paragraph types)
5. Final cleanup (spacing, structural normalization)
6. UI feedback (status, completion message)

## Design Principles

- Separation of concerns (formatting vs replacement vs validation)
- No hidden side-effects: each step returns quickly or fails safe
- Defensive access to Word objects (null/collection guards)
- Forward compatibility: deprecated features (logging, backups) kept as inert flags
- Minimal surface changes to ease diff review and auditing

## Error Handling Strategy

- Fail-soft: errors in non-critical formatting steps are caught and surfaced without aborting entire pipeline where safe
- Centralized safe wrappers (`modSafety`) for operations likely to raise runtime errors on certain Word versions

## Configuration Philosophy

- INI keys map directly to `ConfigSettings` fields (lowercased comparison)
- Deprecated keys retained to avoid runtime errors in existing user INI files
- Defaults emphasize safety & consistent formatting rather than raw speed

## Future Extensions (Roadmap Snippet)

- Reinstate logging (structured, size-limited)
- Reinstate backups with retention and optional encryption
- Add test harness (synthetic document generator) for regression validation
- Introduce performance profiling toggle separate from functional logging
- Optional accent restoration layer for dialog messages

## Glossary

- "Special Paragraphs": Domain-specific paragraphs (first, second, CONSIDERANDO blocks, Justificativa, Anexo sections)
- "Semantic Replacement": Replacement whose behavior depends on paragraph classification rather than pure text pattern

## Security Notes

- No file writes (logging/backups) in current beta to simplify threat surface
- All changes happen in-memory on the active document context

---
Document version: 1.0.0-Beta Refactor
