# Security Policy

## Supported Versions

We actively support the following versions of excel2md:

| Version | Supported          |
| ------- | ------------------ |
| 1.7.x   | :white_check_mark: |
| < 1.7   | :x:                |

## Reporting a Vulnerability

If you discover a security vulnerability in excel2md, please report it responsibly by following these steps:

### How to Report

1. **Do NOT** create a public GitHub issue for security vulnerabilities
2. Send a detailed report to the maintainers via:
   - Creating a private security advisory on GitHub (preferred)
   - Opening an issue with the label "security" (for less critical issues)

### What to Include

Please include the following information in your report:

- Description of the vulnerability
- Steps to reproduce the issue
- Potential impact and severity
- Any suggested fixes or mitigations
- Your contact information (optional)

### Example Report

```
Subject: [SECURITY] Potential XXE vulnerability in Excel parsing

Description:
When processing specially crafted Excel files, the openpyxl library
may be vulnerable to XML External Entity (XXE) attacks.

Steps to Reproduce:
1. Create a malicious Excel file with external entity references
2. Run excel2md on the file
3. Observe potential information disclosure

Impact:
An attacker could potentially read local files or cause denial of service.

Suggested Fix:
Disable external entity processing in openpyxl configuration.
```

## Response Timeline

- **Initial Response**: Within 48 hours
- **Status Update**: Within 7 days
- **Resolution**: Depends on severity
  - Critical: Within 14 days
  - High: Within 30 days
  - Medium: Within 60 days
  - Low: Next release cycle

## Security Considerations

### File Processing

excel2md processes Excel files which may contain:

- Macros (`.xlsm` files)
- External links and references
- Embedded objects
- Formulas with potential side effects

**Recommendations:**

1. Only process Excel files from trusted sources
2. Review files before processing if received from external sources
3. Run excel2md in a sandboxed environment when processing untrusted files
4. Be cautious with files containing macros (though excel2md itself doesn't execute them)

### Input Validation

excel2md includes several security measures:

- Uses `read_only=True` mode in openpyxl to prevent file modifications
- Uses `data_only=True` to avoid executing formulas
- Limits cell processing with `max_cells_per_table` option
- Sanitizes Markdown output to prevent injection attacks

### Output Security

When using the generated Markdown files:

- Be aware that hyperlinks from Excel files are preserved in output
- Review generated Markdown before publishing
- Be cautious of potentially malicious URLs in source Excel files
- Use `--hyperlink-mode text_only` to exclude URLs if concerned

### Dependencies

This project depends on:

- `openpyxl >= 3.1.5`: Excel file processing

We monitor security advisories for these dependencies and update when necessary.

## Security Best Practices

When using excel2md:

1. **Keep Updated**: Always use the latest version
2. **Review Input**: Inspect Excel files before processing
3. **Sandbox Processing**: Use containers or VMs for untrusted files
4. **Validate Output**: Review generated Markdown before use
5. **Limit Permissions**: Run with minimal necessary privileges
6. **Monitor Dependencies**: Keep openpyxl and other dependencies updated

## Known Security Limitations

1. **Macro Detection**: excel2md does not execute macros but does not warn about their presence
2. **External Links**: External links in Excel files are processed but not validated
3. **File Size**: Very large files may cause memory issues; use `max_cells_per_table` to limit
4. **Formula Processing**: Formulas are displayed as values; complex formulas are not validated

## Security Updates

Security updates will be released as:

- Patch versions (e.g., 1.7.1) for minor issues
- Minor versions (e.g., 1.8.0) for significant issues
- Documented in CHANGELOG.md with `[SECURITY]` prefix

## Acknowledgments

We appreciate security researchers who responsibly disclose vulnerabilities. Contributors who report valid security issues will be acknowledged in:

- CHANGELOG.md (unless they prefer to remain anonymous)
- Release notes for the fix

## Questions?

For security-related questions that are not vulnerabilities, please:

- Create an issue with the "security" label
- Reach out to the maintainers

Thank you for helping keep excel2md secure!
