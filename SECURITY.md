# Security Policy

Thank you for using chainsaw-proposituras. This project processes Microsoft Word documents using VBA macros. Please review and follow the guidelines below to use it safely.

## Supported Versions

- Microsoft Word 2010 (v14) and later on Windows
- This project targets Word desktop (no online support)

## Macro Security Guidance

- Enable macros only for trusted documents and locations.
- Recommended setting: “Disable all macros with notification” (File → Options → Trust Center → Trust Center Settings → Macro Settings).
- Prefer running from a trusted folder. Avoid enabling macros globally.

## Permissions and Scope


 File system access is limited to reading the active document and optional header image resources. Backup & logging writes were fully removed.
Current beta (logging & backup disabled):

- No backups are created.
- No log files are written.
- No document content is exported outside Word memory space.

Planned (future reinstatement):

- Backups (full file copies) stored beside source document or in temp path if unsaved.
- Logs: metadata only (timings, paragraph counts, warnings) — never full document dumps.
- Administrators will be able to configure retention & size limits.

## Reporting a Vulnerability

If you discover a potential security issue, please create a private issue or contact the maintainer. Include:

- Word version and Windows version
- Steps to reproduce
- Example document (sanitized, if possible)
- Any error messages or logs

We aim to acknowledge reports within 5 business days.

## Hardening Measures in Code

- Strict document editability & protection checks before formatting
- Disabled wrap-around in critical Find/Replace loops to avoid runaway edits
- Safe wrappers for Word object property access to reduce runtime errors
- ASCII normalization option reduces encoding-related macro load failures
- Graceful fallback: errors surface to user without forcing Word to close

Deprecated (temporarily removed): pre-format backup creation and detailed logging instrumentation.

## Known Limitations

- VBA macros inherit permissions of the host user account
- Disabled features may lead to reduced forensic traceability (no logs)
- When backup feature returns, unsaved documents will route backup to TEMP

## Threat Model (Concise)

| Asset | Threat | Mitigation |
|-------|--------|------------|
| Document content | Accidental destructive formatting | Validation + safety checks + no write until operations succeed |
| User environment | Macro abuse by tampered module | Open-source review + recommended trusted folder usage |
| Confidential data | Leakage via log/backups | Logging & backups disabled in current beta |
| Stability | Infinite replace loops | Guarded Find/Replace (no wrap, bounded iterations) |

## Deprecated / Disabled Features

| Feature | Previous Purpose | Current Status | Planned Return |
|---------|------------------|----------------|----------------|
| Backups | Pre-change recovery | Disabled | Yes (revised retention) |
| Logging | Execution trace & diagnostics | Stub only | Yes (configurable levels) |

Configuration keys for these features remain for backward compatibility but are ignored.

## Best Practices

- Keep Word and Windows updated
- Store project files in version control
- Review macros before enabling in corporate environments
- Use signed macros if required by your organization
