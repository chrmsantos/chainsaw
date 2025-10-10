# Contributors Guide

Thank you for considering a contribution to CHAINSAW PROPOSITURAS! This guide explains how to propose changes and what we expect to keep the project healthy and maintainable.

## Ways to Contribute

- Report bugs and suggest improvements via GitHub Issues
- Improve documentation (README, comments in code, examples)
- Optimize VBA routines, fix edge cases, or improve compatibility with older Word versions
- Add small, well-scoped features aligned with the project goals (see below)

## Development Principles

- Scope: Focus on document standardization in Word (VBA). Avoid adding unrelated features or external dependencies.
- Robustness: Favor defensive code and safe fallbacks that work across Word 2010+.
- Simplicity: Keep the single-module approach unless there is a compelling maintenance reason to split.
- Performance: Prefer batch operations and minimize UI thrashing; keep the code responsive (use DoEvents where helpful).

## Getting Started

1. Fork the repository on GitHub
2. Create a branch for your change
   - Example: `git checkout -b fix/paragraph-spacing`
3. Make your changes
   - Keep edits tight and focused; avoid unrelated formatting changes
   - Follow existing style and naming conventions (.Text, .Count, etc.)
4. Update documentation if behavior changes
5. Open a Pull Request
   - Describe the intent, the change, and any trade-offs
   - Reference related issues, if any

## Coding Style (VBA)

- Use explicit casing for Word/VBA members (e.g., `.Text`, `.Count`, `.Alignment`, `.LeftIndent`)
- Use `Option Explicit` and avoid undeclared variables
- Prefer helper functions for safe property access and formatting (e.g., `SafeSetFont`, `SafeGetCharacterCount`)
- Use constants and centralized messages for user-facing strings
- Avoid heavy logging or telemetry; this project aims for clean local execution

## Tests and Validation

- Manual validation steps are acceptable for this VBA project
- Run the macro on a few sample documents to verify:
  - Title formatting
  - Paragraph spacing and indentation
  - Header image behavior (assets\\stamp.png present/missing)
  - Footer page numbering present
  - No crashes in Word 2010+

## Pull Request Checklist

- [ ] Changes are scoped and documented
- [ ] Code compiles in Word VBA (2010+)
- [ ] README/CHANGELOG updated if needed
- [ ] No new external dependencies

## Code of Conduct

Please be respectful and constructive in issues and PRs. Harassment or discrimination of any kind is not tolerated.

## Licensing

By submitting a contribution, you agree that your work will be licensed under the project’s license (GPL-3.0-or-later). See LICENSE for details.

## Contact

If you have questions about contributions, open an issue or contact the maintainer via email at [chrmsantos@gmail.com](mailto:chrmsantos@gmail.com).
