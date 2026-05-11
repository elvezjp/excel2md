# Contributing to excel2md

[English](./CONTRIBUTING.md) | [日本語](./CONTRIBUTING_ja.md)

This document describes guidelines for contributing to the project.

## How to Contribute

### Reporting Bugs

If you find a bug, please create an Issue on GitHub with the following information:

- A clear and descriptive title
- Steps to reproduce the problem
- Expected behavior
- Actual behavior
- Sample Excel file (if possible)
- excel2md and Python versions
- Operating system

### Feature Requests

Feature requests are welcome! Please create an Issue with:

- A clear and descriptive title
- Detailed description of the proposed feature
- Use cases and benefits
- Related examples or mockups

### Pull Requests

1. **Fork the repository** and create a branch from `main` (format: username/YYYYMMDD-description)
   ```bash
   git checkout -b user/20260105-fix-feature
   ```

2. **Follow the coding style** of the existing codebase
   - Use meaningful variable and function names
   - Add comments for complex logic
   - Follow PEP 8 style guidelines

3. **Write tests** for your changes
   ```bash
   # Run tests
   uv run pytest v2.1.1/tests

   # Run tests with coverage
   uv run pytest v2.1.1/tests --cov=v2.1.1 --cov-report=html
   ```

4. **Update documentation** as needed
   - Update README.md for user-facing changes
   - Update spec.md for specification changes
   - Add examples for new features

5. **Commit your changes** with a clear commit message
   ```bash
   git commit -m "Add feature: description of your changes"
   ```

6. **Push to your fork** and submit a pull request
   ```bash
   git push origin user/20260105-fix-feature
   ```

7. **Wait for review** - maintainers will review the PR and may request changes

## Development Setup

### Prerequisites

- Python 3.10 or higher
- uv package manager

### Installation

```bash
# Install uv (if not already installed)
# Details: https://docs.astral.sh/uv/getting-started/installation/
curl -LsSf https://astral.sh/uv/install.sh | sh

# Clone your fork
git clone https://github.com/YOUR-USERNAME/excel2md.git
cd excel2md

# Install dependencies (including test dependencies)
uv sync --extra test
```

### Running Tests

```bash
# Run all tests
uv run pytest v2.1.1/tests

# Run a specific test file
uv run pytest v2.1.1/tests/test_csv_markdown.py

# Run with coverage
uv run pytest v2.1.1/tests --cov=v2.1.1 --cov-report=html
```

### Testing Your Changes

Before submitting a PR, make sure:

1. All existing tests pass
2. New features have new tests
3. Code coverage is maintained or improved
4. The tool works correctly with various Excel files

## Coding Guidelines

### Python Style

- Follow PEP 8 style guidelines
- Use type hints where appropriate
- Maximum line length: 100 characters (flexible for long strings)
- Use meaningful variable names

### Documentation

- Add docstrings to all public functions and classes
- Use clear and concise language
- Include examples in docstrings where helpful

### Commit Messages

- Use the present tense ("Add feature" not "Added feature")
- Use the imperative mood ("Move cursor to..." not "Moves cursor to...")
- Limit the first line to 72 characters
- Reference Issues and Pull Requests where relevant

Example:
```
Add CSV markdown description exclusion option

- Add --csv-include-description flag
- Update tests for new option
- Update documentation

Closes #123
```

## Versioning

When contributing:
- Focus on the latest version (`v2.1.1/`)
- Maintain backward compatibility where possible
- Document breaking changes clearly

## Code Review Process

1. Maintainers will review pull requests
2. Changes or questions may be requested
3. Once approved, the PR will be merged
4. Contributions will be acknowledged in release notes

## Community Guidelines

- Be respectful and inclusive
- Provide constructive feedback
- Help others when possible

## Questions

If you have questions about contributing, feel free to:
- Create an Issue with the "question" label
- Contact the maintainers
