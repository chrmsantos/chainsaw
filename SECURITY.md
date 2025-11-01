# Security Policy

## Overview

This document describes the security expectations for the Chainsaw automation tooling. The project helps authors format Microsoft Word documents, so the macro and its companion scripts may run with the same privileges as the signed-in user. Treat the repository as sensitive code that can touch user documents and interact with file shares.

## Reporting Vulnerabilities

- Email the maintainer team at `seguranca@chrmsantos.dev` with the subject line "Security: Chainsaw".
- Provide reproduction steps, the affected version or commit SHA, and any proof-of-concept files or scripts.
- Do not open public issues for undisclosed vulnerabilities.
- Expect an acknowledgement within three business days and a remediation plan within ten.

## Coordinated Disclosure

We follow a responsible disclosure model. We will collaborate on a fix before announcing the issue publicly. Credit is granted to reporters who request it.

## Development Practices

- Enable `Option Explicit` in VBA modules and avoid untrusted dynamic code evaluation.
- Validate any file paths, URLs, or user-supplied text before acting on them.
- Keep logging free of sensitive content; redact names or document fragments when possible.
- Run scripts with the principle of least privilege—avoid elevated shells unless required.

## Dependency Management

- Track external dependencies (PowerShell modules, COM libraries) in `README.md` and update them quarterly.
- Use vendor signatures (e.g., Microsoft Office updates) and verify package hashes when available.
- Remove unused libraries and macros during routine maintenance to reduce the attack surface.

## Secure Distribution

- Publish releases through the official repository only.
- Sign release archives when possible and document verification steps.
- Provide a change log detailing security-impacting fixes.

## Incident Response

1. Triage the report and assign a severity rating.
2. Create a private patch branch with the proposed fix.
3. Request peer review focusing on security impact.
4. Notify affected users once a patched build is available.
5. Document lessons learned and update this policy if required.
