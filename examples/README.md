# Runnable examples

This directory contains standalone programs that demonstrate the package with
real `main` functions. They are separate from the `Example*` functions in the
package tests: those are documentation tests, while these programs can be run
directly with `go run`.

All examples are part of the same Go module (`github.com/zignd/xlsx`), so they
run without a separate `go.mod` or `replace` directive.

## Prerequisites

- Go 1.24 or newer
- Run the commands from the repository root

## Examples

| Example | What it shows |
| --- | --- |
| [`basic`](basic) | Writing and reading a small keyed workbook with strings, numbers, booleans, floats and dates. |
| [`read`](read) | Validating and streaming an existing workbook without loading a whole sheet into memory. |
| [`complex`](complex) | A larger workbook with multiple sheets, styles, formulas, merged cells, hyperlinks, freeze panes, auto-filters, shared strings and batched streaming writes. |

## Quick start

```sh
# 1. Write a small workbook.
go run ./examples/basic -out /tmp/people.xlsx

# 2. Read it back in a memory-flat stream.
go run ./examples/read -in /tmp/people.xlsx

# 3. Generate the feature-rich example.
go run ./examples/complex -out /tmp/dashboard.xlsx -rows 10000
```

Each example writes or reads `.xlsx` files only when you run it. Generated files
are not committed.

## Notes

- The examples use only the public API of `github.com/zignd/xlsx`.
- The `complex` example intentionally exercises most writer features; it is a
  good starting point to adapt for your own reports.
- If you run an example from outside the repository root, use an absolute
  `-out`/`-in` path or run `go run github.com/zignd/xlsx/examples/basic` from a
  module that requires this package.

## Troubleshooting

### `compile: version "goX" does not match go tool version "goY"`

This is a local Go installation problem, not an example error. It means the
`GOROOT` environment variable points at a different Go installation than the
`go` binary on your `PATH`.

Check it with:

```sh
which -a go
go env GOROOT GOVERSION
```

If `GOROOT` and `GOVERSION` disagree, either unset `GOROOT` so Go can detect its
own root:

```sh
unset GOROOT
go run ./examples/complex -out /tmp/dashboard.xlsx -rows 10000
```

or point both at the same installation:

```sh
export GOROOT="$(env -u GOROOT go env GOROOT)"
export PATH="$GOROOT/bin:$PATH"
```

To fix it permanently, remove the stale `export GOROOT=...` line from your shell
profile (for example `~/.zshrc`) and open a new terminal.

