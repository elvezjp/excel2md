# Contributing to excel2md

Thank you for your interest in contributing to excel2md! This document provides guidelines for contributing to the project.

## How to Contribute

### Reporting Bugs

If you find a bug, please create an issue on GitHub with the following information:

- A clear and descriptive title
- Steps to reproduce the issue
- Expected behavior
- Actual behavior
- Sample Excel file (if possible)
- Version of excel2md and Python
- Operating system

### Suggesting Enhancements

Enhancement suggestions are welcome! Please create an issue with:

- A clear and descriptive title
- Detailed description of the proposed feature
- Use cases and benefits
- Any relevant examples or mockups

### Pull Requests

1. **Fork the repository** and create your branch from `main`
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Follow the coding style** of the existing codebase
   - Use meaningful variable and function names
   - Add comments for complex logic
   - Follow PEP 8 style guidelines

3. **Write tests** for your changes
   ```bash
   # Run tests
   pytest v1.7/tests

   # Run tests with coverage
   pytest v1.7/tests --cov=v1.7 --cov-report=html
   ```

4. **Update documentation** if needed
   - Update README.md for user-facing changes
   - Update spec.md for specification changes
   - Add examples if introducing new features

5. **Commit your changes** with clear commit messages
   ```bash
   git commit -m "Add feature: description of your changes"
   ```

6. **Push to your fork** and submit a pull request
   ```bash
   git push origin feature/your-feature-name
   ```

7. **Wait for review** - maintainers will review your PR and may request changes

## Development Setup

### Prerequisites

- Python 3.9 or higher
- pip or uv package manager

### Installation

```bash
# Clone your fork
git clone https://github.com/YOUR-USERNAME/excel2md.git
cd excel2md

# Install dependencies
pip install openpyxl

# Install development dependencies
pip install pytest pytest-cov
```

### Running Tests

```bash
# Run all tests
pytest v1.7/tests

# Run specific test file
pytest v1.7/tests/test_csv_markdown.py

# Run with coverage
pytest v1.7/tests --cov=v1.7 --cov-report=html
```

### Testing Your Changes

Before submitting a PR, please ensure:

1. All existing tests pass
2. New tests are added for new features
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

- Use present tense ("Add feature" not "Added feature")
- Use imperative mood ("Move cursor to..." not "Moves cursor to...")
- Limit the first line to 72 characters or less
- Reference issues and pull requests when relevant

Example:
```
Add CSV markdown description exclusion option

- Add --csv-include-description flag
- Update tests for new option
- Update documentation

Closes #123
```

## Version Management

This project maintains multiple versions:

- `v1.7/`: Latest version with all features
- `v1.6/`: Previous version (for reference)
- `v1.5/`: Previous version (for reference)

When contributing:
- Focus on the latest version (`v1.7/`)
- Maintain backward compatibility when possible
- Document breaking changes clearly

## Code Review Process

1. A maintainer will review your pull request
2. They may request changes or ask questions
3. Once approved, your PR will be merged
4. Your contribution will be acknowledged in release notes

## Community Guidelines

- Be respectful and inclusive
- Provide constructive feedback
- Help others when possible
- Follow the code of conduct

## Questions?

If you have questions about contributing, feel free to:
- Create an issue with the "question" label
- Reach out to the maintainers

Thank you for contributing to excel2md!
