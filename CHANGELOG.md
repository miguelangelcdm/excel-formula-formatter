# Change Log

All notable changes to the "excel-formula" extension will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/), and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased] / 0.1.1
### Added
- Syntax diagnostics for `.xlf` documents: unbalanced parentheses/braces, unterminated string literals, trailing commas before `)` or `}`, and missing argument commas inside calls are flagged inline (issue #2).
- `run-syntax-tests.js` node test harness covering the checker, the issue #2 sample, and formatter round-trips.

### Fixed
- Accepting a function completion while typing `(` no longer produces an extra pair of parentheses; the typed `(` now closes the completion widget instead of committing the snippet (issue #3).

## [0.0.1]
### Added
- Syntax highlighting for functions, cell references, structured references, booleans, and Excel error literals.
- IntelliSense completions for frequently used Excel functions.
- Formatter that expands minimised formulas into indented, line-by-line layouts.
- Language configuration with indentation rules and editor defaults tailored for Excel formulas.
