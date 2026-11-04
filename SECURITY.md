# Security policy

## Supported versions

Security fixes are provided for the latest release line. The package has no
runtime dependencies; the main risks are malformed output, resource exhaustion
when reading untrusted workbooks, and unsafe temporary-file handling.

## Reporting a vulnerability

Please do **not** open a public issue for a security problem. Use GitHub's
private vulnerability reporting:

<https://github.com/zignd/xlsx/security/advisories/new>

Include the affected version, a minimal reproduction, and any suggested fix.
You can expect an acknowledgement within a few days. Once a fix is available, a
coordinated disclosure and release will be published.

## Reader safety

`Open` and `Validate` enforce configurable limits for ZIP entry count,
uncompressed size, rows, columns and cell length. They reject absolute paths,
`..` path segments, backslashes, NUL bytes and duplicate ZIP entries. Encrypted
workbooks and XLSB packages are rejected with sentinel errors rather than
parsed.

The reader is designed for untrusted input, but callers should still apply
their own request-level limits and timeouts. Use `WithReadContext` to cancel
large reads. XML is parsed iteratively with `encoding/xml`; nesting does not
recurse into the Go stack, and total decompressed size is capped by
`MaxUncompressedSize`.

## Formula injection

Strings are written as inline or shared strings, never as formulas. A string
beginning with `=` is therefore data, not executable spreadsheet content.
Formulas are only written when a caller explicitly sets `Cell.Formula` or calls
`Cell.SetFormula`. If a workbook is later exported to CSV, apply the usual CSV
formula-injection escaping at that boundary.

## Temporary files

The writer creates a private `0o700` subdirectory inside the configured temp
directory and removes it after a successful commit unless `KeepTempFiles` is
enabled. The reader creates private temporary files for disk-backed shared
strings and removes them on `Reader.Close`.
