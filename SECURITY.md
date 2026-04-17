# Security Policy

[English](./SECURITY.md) | [日本語](./SECURITY_ja.md)

## Supported Versions

We support the latest version:

| Version | Supported          |
| ------- | ------------------ |
| 2.1.x   | :white_check_mark: |
| < 2.1   | :x:                |

## Reporting a Vulnerability

If you discover a security vulnerability in excel2md, please follow responsible disclosure practices:

### How to Report

Report security vulnerabilities via **GitHub Private Security Advisory**, regardless of severity.

1. **Do NOT** create a public GitHub Issue for security vulnerabilities, regardless of severity
2. Open a Private Security Advisory at:
   https://github.com/elvezjp/excel2md/security/advisories/new

For security-related questions that are **not** vulnerabilities (e.g., best practices, configuration), see the [Questions](#questions) section.

### What to Include

Please include the following information in your report:

- Description of the vulnerability
- Steps to reproduce the issue
- Potential impact and severity
- Suggested fix or mitigation
- Contact information (optional)

### Example Report

```
Subject: [SECURITY] Potential XXE vulnerability during Excel parsing

Description:
When processing a specially crafted Excel file, the openpyxl library may be
vulnerable to XML External Entity (XXE) attacks.

Steps to reproduce:
1. Create a malicious Excel file containing external entity references
2. Run excel2md against the file
3. Observe potential information disclosure

Impact:
An attacker could potentially read local files or cause denial of service.

Suggested fix:
Disable external entity processing in openpyxl configuration.
```

## Response Timeline

- **Initial response**: Within 48 hours
- **Status update**: Within 7 days
- **Resolution**: Depending on severity
  - Critical: Within 14 days
  - High: Within 30 days
  - Medium: Within 60 days
  - Low: Next release cycle

## Security Considerations

### File Processing

excel2md processes Excel files that may contain:

- Macros (`.xlsm` files)
- External links and references
- Embedded objects
- Formulas with potential side effects

**Recommendations:**

1. Only process Excel files from trusted sources
2. Verify files received from external sources before processing
3. Run excel2md in a sandboxed environment when processing untrusted files
4. Be cautious with files containing macros (though excel2md does not execute macros)

### Input Validation

excel2md includes the following security measures:

- Uses `read_only=True` mode in openpyxl to prevent file modification
- Uses `data_only=True` to avoid formula execution
- Limits cell processing with the `max_cells_per_table` option
- Sanitizes Markdown output to prevent injection attacks

### Output Security

When using generated Markdown files, be aware that:

- Hyperlinks from Excel files are preserved in the output
- Review generated Markdown before publishing
- Be cautious of malicious URLs from source Excel files
- Use `--hyperlink-mode text_only` to exclude URLs

### Dependencies

This project depends on:

- `openpyxl >= 3.1.5`: Excel file processing

We monitor security advisories for these dependencies and update as needed.

## Security Best Practices

Recommendations when using excel2md:

1. **Keep up to date**: Always use the latest version
2. **Verify input**: Inspect Excel files before processing
3. **Sandbox processing**: Use containers or VMs for untrusted files
4. **Validate output**: Review generated Markdown before use
5. **Limit permissions**: Run with minimal required privileges
6. **Monitor dependencies**: Keep openpyxl and other dependencies updated

## Known Security Limitations

1. **Macro detection**: excel2md does not execute macros but does not warn about their presence
2. **External links**: External links in Excel files are processed but not validated
3. **File size**: Very large files may cause memory issues. Use `max_cells_per_table` to limit
4. **Formula handling**: Formulas are displayed as values. Complex formulas are not validated

## Security Updates

Security updates are released as:

- Patch versions for minor issues (e.g., 2.0.1)
- Minor versions for significant issues (e.g., 2.1.0)
- Documented in CHANGELOG.md with `[SECURITY]` prefix

## Acknowledgments

We appreciate security researchers who report vulnerabilities responsibly. Those who report valid security issues will be acknowledged in:

- CHANGELOG.md (unless anonymity is preferred)
- Release notes for the fix

## Questions

For security-related questions that are not vulnerabilities, please:

- Create an Issue with the "security" label
- Contact the maintainers
