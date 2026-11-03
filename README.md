# xlsx

[![CI](https://github.com/zignd/xlsx/actions/workflows/ci.yml/badge.svg?branch=master)](https://github.com/zignd/xlsx/actions/workflows/ci.yml)
[![Go Reference](https://pkg.go.dev/badge/github.com/zignd/xlsx.svg)](https://pkg.go.dev/github.com/zignd/xlsx)
[![Go Report Card](https://goreportcard.com/badge/github.com/zignd/xlsx)](https://goreportcard.com/report/github.com/zignd/xlsx)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](#license)

`xlsx` is a small, dependency-free Go library for creating Excel `.xlsx`
workbooks **incrementally**. Rows are buffered only until you call
`CommitRows`, then streamed to temporary files on disk, so memory usage stays
flat even for very large spreadsheets.

## Features

- Streaming writer: commit rows in batches instead of holding a whole sheet in
  memory.
- Keyed columns with an automatically generated header row.
- Strings, booleans, integers, unsigned integers, floats, `time.Time`, `[]byte`,
  and `database/sql/driver.Valuer` values (including `sql.Null*`).
- Correct XML escaping, date serials and cell references.
- Multi-sheet workbooks, column widths, deterministic XML, and no third-party
  dependencies.
- Tested on Linux, macOS, and Windows with Go 1.24, 1.25, and 1.26.

## Installation

```sh
go get github.com/zignd/xlsx@latest
```

Requires Go 1.24 or newer.

## Quick start

```go
package main

import (
	"log"

	"github.com/zignd/xlsx"
)

func main() {
	workbook := xlsx.NewWorkbook("report.xlsx")
	worksheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Sales"})

	for i := 0; i < 1_000_000; i++ {
		row, err := worksheet.AddRow()
		if err != nil {
			log.Fatal(err)
		}

		product, err := row.AddCell()
		if err != nil {
			log.Fatal(err)
		}
		product.Value = "Keyboard"

		amount, err := row.AddCell()
		if err != nil {
			log.Fatal(err)
		}
		amount.Value = 49.90

		// Flush this batch to disk. Repeat CommitRows to keep memory flat.
		if err := worksheet.CommitRows(); err != nil {
			log.Fatal(err)
		}
	}

	// Worksheets without pending rows are finalized automatically.
	if err := workbook.Commit(); err != nil {
		log.Fatal(err)
	}
}
```

`CommitRows` can be called once per row (as above) or once per batch. Calling it
per batch is usually faster because it amortizes the append/flush work:

```go
const batchSize = 10_000

for i := 0; i < total; i++ {
	row, _ := worksheet.AddRow()
	// ... add cells ...
	if (i+1)%batchSize == 0 {
		if err := worksheet.CommitRows(); err != nil {
			log.Fatal(err)
		}
	}
}
if err := worksheet.CommitRows(); err != nil { // flush the final partial batch
	log.Fatal(err)
}
if err := workbook.Commit(); err != nil {
	log.Fatal(err)
}
```

## Keyed columns

Defining columns writes a header row and lets you address cells by key. Missing
cells are written as blank cells so column positions stay stable.

```go
worksheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Sales"})

err := worksheet.DefineColumns([]*xlsx.WorksheetColumn{
	{Key: "product", Value: "Product", Width: 30},
	{Key: "amount", Value: "Amount"},
	{Key: "sold_at", Value: "Sold at", Width: 22},
})
if err != nil {
	log.Fatal(err)
}

row, err := worksheet.AddRow()
if err != nil {
	log.Fatal(err)
}

product, err := row.AddCellWithKey("product")
if err != nil {
	log.Fatal(err)
}
product.Value = "Keyboard"

amount, err := row.AddCellWithKey("amount")
if err != nil {
	log.Fatal(err)
}
amount.Value = 49.90

if err := worksheet.CommitRows(); err != nil {
	log.Fatal(err)
}
if err := workbook.Commit(); err != nil {
	log.Fatal(err)
}
```

## Supported values

| Go type | Written as |
| --- | --- |
| `string`, `[]byte`, named string and byte-slice types | Inline string |
| `bool`, named bool types | Excel boolean |
| Signed and unsigned integers | Number |
| `float32`, `float64` | Number |
| `time.Time` | Excel date serial with a `yyyy-mm-dd hh:mm:ss` style |
| `database/sql/driver.Valuer` | The value returned by `Value()` |
| `nil` | Blank cell |

Unsupported values (structs other than `time.Time`, maps, `NaN`, infinities,
and so on) return an error from `CommitRows`, wrapped with
`ErrUnsupportedType`.

Dates are written using their wall-clock components (the location is ignored),
which is how Excel stores date-times. The 1900 date system is used.

## API overview

```go
func NewWorkbook(filePath string) *Workbook

func (wb *Workbook) AddWorksheet(opts *WorksheetOptions) *Worksheet
func (wb *Workbook) Commit() error
func (wb *Workbook) Cleanup() error
func (wb *Workbook) HasPendingWorksheets() bool
func (wb *Workbook) Worksheets() []*Worksheet

func (ws *Worksheet) DefineColumns(columns []*WorksheetColumn) error
func (ws *Worksheet) AddRow() (*Row, error)
func (ws *Worksheet) CommitRows() error
func (ws *Worksheet) Commit() error

func (r *Row) AddCell() (*Cell, error)
func (r *Row) AddCellWithKey(key string) (*Cell, error)
```

Errors that callers are expected to handle are exported as sentinels and can be
checked with `errors.Is`:

```go
if err := workbook.Commit(); errors.Is(err, xlsx.ErrPendingRows) {
	// some worksheet still has rows waiting to be committed
}
```

See the [package documentation](https://pkg.go.dev/github.com/zignd/xlsx) for
the complete API and examples.

## Design notes and limitations

- The library is a **writer only**; it does not read XLSX files.
- A worksheet is written to a temporary directory while rows are appended, then
  packaged into a ZIP archive when `Workbook.Commit` runs. The temporary files
  are removed automatically after a successful commit; call `Cleanup` if you
  abandon a workbook.
- `Workbook.Commit` finalizes worksheets that have no pending rows, but it does
  **not** flush pending rows for you. Call `CommitRows` first.
- Strings are always written as inline strings. Formulas and rich text are not
  supported.
- `Workbook`, `Worksheet`, and `Row` are not safe for concurrent use. Use one
  workbook per goroutine; independent workbooks can be written concurrently.
- Worksheet names are validated when the workbook is committed (31 characters
  maximum, no `[]:*?/\\`, no leading/trailing apostrophe, unique,
  case-insensitive).
- Excel's hard limits are enforced: 1,048,576 rows and 16,384 columns per
  worksheet.

## Benchmarks

Run the benchmarks with:

```sh
make bench
# or
go test -run '^$' -bench . -benchmem ./...
```

The suite covers cell-reference generation, XML escaping, date conversion,
row buffering, `CommitRows`, and full workbook commits. Results depend on the
machine and storage; the benchmark smoke test in CI only verifies that they
compile and run.

## Development

```sh
make test       # go test ./...
make test-race  # race detector
make cover      # coverage profile + function report
make bench      # benchmarks
make lint       # golangci-lint
make check      # tidy + format + vet + race tests + lint
```

CI runs the test matrix, lint, module-tidiness, benchmark, and vulnerability
checks on every push and pull request. Dependabot keeps the Go module and
GitHub Actions up to date.

## Versioning and releases

The project follows [Semantic Versioning](https://semver.org/). Versions are
regular Go module tags, so they are published to
[pkg.go.dev](https://pkg.go.dev/github.com/zignd/xlsx) automatically.

To release from the GitHub UI, run the **Tag release** workflow and choose
`patch`, `minor`, or `major` (optionally with a prerelease suffix), or provide an
exact tag such as `v1.2.3`. The workflow runs the tests, computes or validates
the tag, creates an annotated tag, and pushes it. The tag then triggers
**Release**, which validates the SemVer tag, builds, tests, scans for
vulnerabilities, and creates the GitHub release with generated notes.

The same flow can be driven locally:

```sh
git tag -a v1.2.3 -m "Release v1.2.3"
git push origin v1.2.3
```

## License

MIT. See [LICENSE](LICENSE).
