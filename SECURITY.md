# Security Policy

Thank you for helping keep CHAINSAW PROPOSITURAS and its users safe.

This project is a Microsoft Word VBA module that runs locally, without network access or telemetry. Even so, please report any security concerns you find.

## Supported Versions

We actively support the latest code on the `main` branch and the most recent tagged release.

Target runtime: Microsoft Word 2010 or later (Windows). Older Word versions are best-effort only.

## Reporting a Vulnerability

Prefer confidential disclosure by email:

- Email: [chrmsantos@gmail.com](mailto:chrmsantos@gmail.com)
- Subject: "[Security] chainsaw-proposituras"


Please include where possible:

- A clear description of the issue and potential impact
- Steps to reproduce (a minimal .docx sample, if applicable)
- Your environment (Word version, Windows version)
- The project version or commit hash (from README or module header)

If email is not an option, you may open a GitHub issue but avoid sharing sensitive data or private documents. Mark the issue with the Security label.

Acknowledgment and timelines (best-effort):

- Acknowledge within 7 days
- Triage and proposed path within 14 days
- Fix or mitigation target within 30 days for high-severity issues

## Scope

In scope:

- The VBA module(s) and documented usage flows
- Macro behavior (formatting, file/path handling, dialog prompts)

Out of scope:

- Vulnerabilities in Microsoft Word, Office, or Windows
- Third-party add-ins or templates not included in this repository

## Macro Security Guidance

 
- Only enable macros for documents from trusted sources
- Prefer storing this project in a Trusted Location (Word Trust Center)
- Recommended Trust Center settings: "Disable all macros with notification"
- Review imported VBA code before use in your environment

Paths used by the module:

- Header image is read from `assets\\stamp.png` relative to the active document folder; if missing, it is safely skipped. No network I/O is performed.

## Responsible Disclosure

We will credit reporters (by handle or name) upon request after a fix is shipped. If you prefer to remain anonymous, let us know.

Thank you for your help and professionalism.
