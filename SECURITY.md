# Security Policy

## Supported Versions

CHAINSAW PROPOSITURAS follows a best-effort security maintenance policy. Only the latest stable release is actively supported with security updates.

| Version      | Supported          |
| ------------ | ------------------ |
| 2.0.x        | :white_check_mark: |
| < 2.0        | :x:                |

## Macro Security Policy

### Security Considerations for VBA Macros

CHAINSAW PROPOSITURAS is a VBA-based solution that requires macro execution in Microsoft Word. The following security guidelines should be observed:

#### Recommended Security Settings

1. **Macro Security Level**: Configure Word to "Disable all macros with notification" (recommended) or "Disable all macros except digitally signed macros"
2. **Trusted Locations**: Add the directory containing chainsaw-fprops to Word's trusted locations
3. **Code Signing**: For enterprise deployments, consider digitally signing the VBA code
4. **Network Isolation**: Run on systems with appropriate network security controls

#### Security Features Built-in

- **Input Validation**: All user inputs and document content are validated before processing
- **Error Handling**: Comprehensive error handling prevents unexpected behavior
- **Backup Creation**: Automatic document backup before any modifications
- **Limited Scope**: Macro operations are restricted to document formatting only
- **No External Connections**: The macro does not connect to external services or networks
- **Read-Only Operations**: When possible, operations are performed in read-only mode

#### Enterprise Deployment Recommendations

1. **Group Policy**: Use Group Policy to manage macro security settings across the organization
2. **Antivirus Scanning**: Ensure antivirus software scans VBA content
3. **User Training**: Train users to verify macro source before enabling execution
4. **Regular Updates**: Keep Microsoft Office and CHAINSAW PROPOSITURAS updated
5. **Access Control**: Limit macro execution permissions to authorized users only

#### Code Review and Auditing

- All VBA code is open source and available for security review
- Regular security audits are encouraged before deployment
- Code changes are tracked through version control
- No obfuscated or hidden code is used

### Data Privacy

- The macro only processes local document content
- No data is transmitted to external servers
- Temporary files are cleaned up after processing
- Log files contain only technical information, no document content

## Reporting a Vulnerability

If you discover a security vulnerability in CHAINSAW PROPOSITURAS, please report it privately to the maintainer:

- Email: <chrmsantos@gmail.com>
- GitHub Issues: [https://github.com/chrmsantos/chainsaw-proposituras/issues](https://github.com/chrmsantos/chainsaw-proposituras/issues) (mark as "Security" and do **not** disclose details publicly)

Please include as much detail as possible to help reproduce and address the issue. You can expect a response within 7 business days. If the vulnerability is confirmed, a fix will be prioritized and released as soon as possible.

Thank you for helping keep CHAINSAW PROPOSITURAS secure!
