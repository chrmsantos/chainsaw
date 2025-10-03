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

- No internet access is required by the macros.
- File system access is limited to:
  - Creating backup copies next to your document (or TEMP if unsaved)
  - Writing plain-text log files in the same folder
  - Opening Windows Explorer to show the logs/backup folder on demand

## Data Handling

- Backups contain full copies of your documents; protect the folder accordingly.
- Log files contain file names and processing metadata, but not document contents (except small excerpts in warnings).
- Delete logs and backups if they include sensitive information you no longer need.

## Reporting a Vulnerability

If you discover a potential security issue, please create a private issue or contact the maintainer. Include:

- Word version and Windows version
- Steps to reproduce
- Example document (sanitized, if possible)
- Any error messages or logs

We aim to acknowledge reports within 5 business days.

## Hardening Measures in Code

- Backups created before changes; retry logic and permission checks
- Strict document editability checks before formatting
- Disabled wrap-around in critical Find/Replace loops to avoid runaway edits
- Logging with explicit levels; errors never crash Word intentionally

## Known Limitations

- VBA macros inherit permissions of the host user account
- Encoding in log files is ASCII-safe; non-ASCII characters are replaced with '?'
- Running on unsaved documents uses the TEMP folder for logs/backups

## Best Practices

- Keep Word and Windows updated
- Store project files in version control
- Review macros before enabling in corporate environments
- Use signed macros if required by your organization
