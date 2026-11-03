# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- Dependency-free implementation using only the Go standard library.
- Streaming row writer with an internal buffered writer and automatic temporary
  file cleanup.
- Support for strings, booleans, signed and unsigned integers, floats,
  `time.Time`, `[]byte`, and `database/sql/driver.Valuer` values.
- Exported sentinel errors for `errors.Is` checks.
- Worksheet and column validation, including duplicate and invalid sheet names.
- Column widths via `WorksheetColumn.Width`.
- Correct OOXML content types, workbook relationships, styles part, date styles,
  XML escaping, one-based row references, and cell references.
- Unit, integration, fuzz, and benchmark tests.
- GitHub Actions CI matrix (Linux, macOS, Windows; Go 1.24, 1.25, 1.26), lint,
  module tidiness, benchmark smoke test, and `govulncheck`.
- Release workflow for SemVer tags and a manual tag workflow for patch, minor,
  and major releases.
- Makefile and golangci-lint configuration.

### Changed

- `Workbook.Commit` finalizes worksheets that have no pending rows
  automatically.
- Zip files are written through a temporary file in the destination directory
  and then renamed.
- The output directory is created automatically when it does not exist.
- `Workbook.Cleanup` is idempotent and is called after a successful commit.

### Fixed

- Invalid `r="0"` row numbers in worksheet XML.
- Missing `styles.xml` part and its relationships/content type.
- Worksheet content type incorrectly pointing at the workbook content type.
- XML special characters in cell values and sheet names.
- Panic in `DefineColumns` when adding a header cell failed.
- Date serials losing sub-second precision and depending on the time zone.
- Correct `driver.Valuer` handling, including typed nil and `nil` values.

## [0.0.0] - 2018-04-01

Initial public repository.
