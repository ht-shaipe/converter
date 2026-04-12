# Contributing to File Converter

Thank you for your interest in contributing to the file_converter project! This document provides guidelines and instructions for contributing.

## Code of Conduct

Please be respectful and constructive in all interactions. We welcome contributors of all backgrounds and experience levels.

## How to Contribute

### Reporting Bugs

1. Check if the bug has already been reported in the Issues section
2. If not, create a new issue with:
   - A clear, descriptive title
   - Steps to reproduce the bug
   - Expected behavior vs actual behavior
   - Your environment (OS, Rust version, etc.)
   - Any relevant error messages or logs

### Suggesting Features

1. Check if the feature has already been suggested
2. Create a new issue with:
   - A clear description of the proposed feature
   - Use cases and benefits
   - Any potential implementation ideas

### Submitting Pull Requests

1. **Fork the repository** and clone it locally
2. **Create a branch** for your changes:
   ```bash
   git checkout -b feature/your-feature-name
   ```
3. **Make your changes** following the coding style guidelines
4. **Add tests** for new functionality
5. **Ensure all tests pass**:
   ```bash
   cargo test
   ```
6. **Run clippy** to check for common mistakes:
   ```bash
   cargo clippy
   ```
7. **Format your code**:
   ```bash
   cargo fmt
   ```
8. **Commit your changes** with clear, descriptive commit messages
9. **Push to your fork** and open a pull request

## Coding Guidelines

### Code Style

- Follow the [Rust API Guidelines](https://rust-lang.github.io/api-guidelines/)
- Use `cargo fmt` to format code consistently
- Run `cargo clippy` to catch common mistakes
- Write clear, self-documenting code with appropriate comments

### Documentation

- Add doc comments (`///`) to all public items
- Include examples in documentation when helpful
- Update README.md if adding new features or changing behavior

### Testing

- Write unit tests for new functionality
- Add integration tests for complex features
- Ensure test coverage is maintained or improved

### Error Handling

- Use the existing error types in `error.rs`
- Add new error variants as needed
- Provide clear, actionable error messages

## Development Setup

1. Install Rust (latest stable version recommended)
2. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/file_converter.git
   cd file_converter
   ```
3. Build the project:
   ```bash
   cargo build
   ```
4. Run tests:
   ```bash
   cargo test
   ```

## Release Process

Releases follow semantic versioning (MAJOR.MINOR.PATCH):

- **MAJOR**: Breaking changes
- **MINOR**: New features (backward compatible)
- **PATCH**: Bug fixes (backward compatible)

To prepare a release:

1. Update CHANGELOG.md with all changes
2. Update version numbers in Cargo.toml
3. Create a release tag
4. Publish to crates.io

## Questions?

Feel free to open an issue for any questions or discussions about contributing.

Thank you for helping make file_converter better!
