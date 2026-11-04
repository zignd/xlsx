# xlsx

[![CI](https://github.com/zignd/xlsx/actions/workflows/ci.yml/badge.svg?branch=master)](https://github.com/zignd/xlsx/actions/workflows/ci.yml)
[![Go Reference](https://pkg.go.dev/badge/github.com/zignd/xlsx.svg)](https://pkg.go.dev/github.com/zignd/xlsx)
[![Go Report Card](https://goreportcard.com/badge/github.com/zignd/xlsx)](https://goreportcard.com/report/github.com/zignd/xlsx)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](#license)

`xlsx` is a dependency-free Go library for creating and reading Excel `.xlsx`
workbooks **incrementally**. Writing buffers a batch of rows and streams them to
temporary files; reading streams worksheet rows with `xml.Decoder`. Neither
direction needs to hold a whole workbook in memory.

## Features

- Streaming writer with configurable batching, temp directory, buffer size and
  ZIP compression.
- Streaming reader with row-bounded memory and in-memory, disk-backed or
  automatic shared-string storage.
- Keyed columns, typed cell helpers, formulas, styles, hyperlinks, merged cells,
  freeze panes, auto-filters, row heights and column widths.
- Strings, booleans, integers, unsigned integers, floats, `time.Time`, `[]byte`,
  and `database/sql/driver.Valuer` values (including `sql.Null*`).
- Structural validation for packages, relationships, content types, cell
  references, styles and shared strings.
- Context cancellation, Excel limits, ZIP-bomb protection and clear sentinel
  errors.
- No third-party runtime dependencies.
- Tested on Linux, macOS, and Windows with Go 1.24, 1.25, and 1.26.

## Installation

```sh
go get github.com/zignd/xlsx@latest
```

Requires Go 1.24 or newer.

## Examples

Runnable programs live in [`examples/`](examples). Each has its own `README.md`
with instructions and expected output:

```sh
# Small keyed workbook, then read it back.
go run ./examples/basic -out /tmp/people.xlsx
go run ./examples/read -in /tmp/people.xlsx

# A feature-rich dashboard with styles, formulas, merges and 10k rows.
go run ./examples/complex -out /tmp/dashboard.xlsx -rows 10000
```

The `Example*` functions in the root tests are documentation examples for
`go test`; the programs under `examples/` are standalone runnable commands.

## Writing

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
		product.SetString("Keyboard")

		amount, err := row.AddCell()
		if err != nil {
			log.Fatal(err)
		}
		amount.SetFloat(49.90)

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

`CommitRows` can be called once per row or once per batch. Batching is usually
faster:

```go
const batchSize = 10_000

for i := 0; i < total; i++ {
	row, _ := worksheet.AddRow()
	if err := row.Append("Keyboard", 49.90); err != nil {
		log.Fatal(err)
	}
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

### Keyed columns

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

product, _ := row.AddCellWithKey("product")
product.SetString("Keyboard")

amount, _ := row.AddCellWithKey("amount")
amount.SetFloat(49.90)

if err := worksheet.CommitRows(); err != nil {
	log.Fatal(err)
}
```

### Styles, formulas, hyperlinks and layout

```go
style, err := workbook.AddStyle(xlsx.Style{
	Bold:         true,
	NumberFormat: "#,##0.00",
	FillColor:    "FFF2CC",
	Horizontal:   "center",
})
if err != nil {
	log.Fatal(err)
}

worksheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Data"})
if err := worksheet.FreezePanes(1, 0); err != nil { // freeze the header row
	log.Fatal(err)
}
if err := worksheet.SetAutoFilter("A1:D3"); err != nil {
	log.Fatal(err)
}
if err := worksheet.MergeCells("A4:B4"); err != nil {
	log.Fatal(err)
}

row, _ := worksheet.AddRow()
amount, _ := row.WriteFloat(1234.5)
amount.SetStyle(style)
amount.SetHyperlink("https://example.com")

formula, _ := row.AddCell()
formula.SetFormula("=SUM(B1:B2)")
```

Formulas are written without the leading `=` and the workbook asks Excel to
recalculate on open.

## Reading

```go
reader, err := xlsx.Open("report.xlsx")
if err != nil {
	log.Fatal(err)
}
defer reader.Close()

for _, info := range reader.Sheets() {
	fmt.Println(info.Name, info.Hidden)
}

sheet, err := reader.OpenSheet("Sales")
if err != nil {
	log.Fatal(err)
}
defer sheet.Close()

for sheet.Next() {
	row := sheet.Row()
	name, ok := row.CellAt(0)
	if !ok {
		continue
	}
	value, err := name.String()
	if err != nil {
		log.Fatal(err)
	}
	fmt.Println(row.Index, value)
}
if err := sheet.Err(); err != nil {
	log.Fatal(err)
}
```

`OpenReaderAt(io.ReaderAt, size)` supports embedding and object storage.
Shared strings use `SharedStringsAuto` by default: small tables stay in memory
and large tables are stored in temporary files and read on demand.

## Validation

```go
if err := xlsx.Validate("report.xlsx"); err != nil {
	log.Fatal(err)
}
```

`Validate` checks ZIP integrity, unsafe entry names, content types, root and
worksheet relationships, required parts, XML well-formedness, sheet names, cell
references, row/column ordering, style IDs and shared-string indexes. It streams
worksheets and therefore stays memory-flat.

## Options and temporary files

```go
workbook, err := xlsx.NewWorkbookWithOptions("report.xlsx", xlsx.Options{
	TempDir:       "/var/tmp/xlsx",
	BufferSize:    128 * 1024,
	Compression:   xlsx.CompressionBestSpeed,
	TempDirMode:   0o700,
	FileMode:      0o600,
	SharedStrings: true,
	Context:       ctx,
})
```

- `TempDir` must exist and be writable. The library creates a private
  `xlsx-*` subdirectory inside it and never deletes the parent directory.
- `KeepTempFiles` leaves the private directory in place for debugging.
- `NewWriter(io.Writer, Options)` writes the final ZIP to an arbitrary writer.
- `Workbook.Close`/`Cleanup` are idempotent and safe to `defer`.

## Supported values

| Go type | Written as |
| --- | --- |
| `string`, `[]byte`, named string and byte-slice types | Inline or shared string |
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
which is how Excel stores date-times. Both 1900 and 1904 date systems are read.

## Limits and safety

- 1,048,576 rows and 16,384 columns per worksheet.
- 32,767 characters per text cell and 8,192 characters per formula.
- 2,079 characters per hyperlink.
- Configurable limits for temp space, ZIP entries, uncompressed size, rows,
  columns and cell length on the reader.
- Reader rejects path traversal, duplicate ZIP entries, encrypted workbooks and
  XLSB packages with clear sentinel errors.
- `Workbook`, `Worksheet`, `Row` and the reader are not safe for concurrent use;
  use one workbook/reader per goroutine.

## Protocol conformance

The writer targets **ECMA-376 5th edition / ISO/IEC 29500 Transitional**
SpreadsheetML, which is what Excel writes by default. The test suite includes:

- structural validation of every generated part,
- round-trip reader/writer tests,
- an openpyxl interoperability job that verifies values, dates, styles,
  hyperlinks, frozen panes, auto-filters, merged cells and formulas,
- fuzzing of the reader and column conversion.

The OOXML standard continues to evolve through ISO/IEC and Microsoft
implementation notes. When changing OOXML output, compare generated files with
Excel, LibreOffice, openpyxl and Excelize and update the compatibility tests.

## API overview

```go
func NewWorkbook(filePath string) *Workbook
func NewWorkbookWithOptions(filePath string, opts Options) (*Workbook, error)
func NewWriter(w io.Writer, opts Options) (*Workbook, error)

func (wb *Workbook) AddWorksheet(opts *WorksheetOptions) *Worksheet
func (wb *Workbook) AddStyle(style Style) (int, error)
func (wb *Workbook) Commit() error
func (wb *Workbook) Close() error
func (wb *Workbook) Cleanup() error
func (wb *Workbook) TempDir() string
func (wb *Workbook) Worksheets() []*Worksheet

func (ws *Worksheet) DefineColumns(columns []*WorksheetColumn) error
func (ws *Worksheet) AddRow() (*Row, error)
func (ws *Worksheet) CommitRows() error
func (ws *Worksheet) Commit() error
func (ws *Worksheet) MergeCells(ranges ...string) error
func (ws *Worksheet) FreezePanes(rows, columns int) error
func (ws *Worksheet) SetAutoFilter(ref string) error

func (r *Row) AddCell() (*Cell, error)
func (r *Row) AddCellWithKey(key string) (*Cell, error)
func (r *Row) Append(values ...interface{}) error
func (r *Row) WriteString(value string) (*Cell, error)
func (r *Row) WriteInt(value int64) (*Cell, error)
func (r *Row) WriteFloat(value float64) (*Cell, error)
func (r *Row) WriteBool(value bool) (*Cell, error)
func (r *Row) WriteTime(value time.Time) (*Cell, error)

func Open(path string, opts ...ReadOption) (*Reader, error)
func OpenReaderAt(r io.ReaderAt, size int64, opts ...ReadOption) (*Reader, error)
func (r *Reader) Sheets() []SheetInfo
func (r *Reader) OpenSheet(name string) (*SheetReader, error)
func (r *Reader) Close() error
func (s *SheetReader) Next() bool
func (s *SheetReader) Row() *RowData
func (s *SheetReader) Err() error
func (s *SheetReader) Close() error
func Validate(path string, opts ...ReadOption) error
func ValidateReaderAt(r io.ReaderAt, size int64, opts ...ReadOption) error
```

Errors intended for callers are exported as sentinels and can be checked with
`errors.Is`, for example `ErrPendingRows`, `ErrReadLimit`, `ErrCorruptFile`,
`ErrEncrypted`, `ErrUnsupportedFormat`, `ErrCellTooLong` and `ErrInvalidStyle`.

See the [package documentation](https://pkg.go.dev/github.com/zignd/xlsx) for
the complete API and runnable examples.

## Benchmarks

```sh
make bench
# or
go test -run '^$' -bench . -benchmem ./...
```

The suite covers typed writes, shared strings, styles, row buffering,
`CommitRows`, full workbook commits, reading, validation and XML escaping.
Results depend on the machine and storage; CI runs a smoke test for every
benchmark.

## Performance guide

- Batch `CommitRows` calls. Flushing once per 1,000–50,000 rows is usually
  faster than flushing every row and still keeps memory bounded.
- Use `CompressionBestSpeed` when CPU is the bottleneck and
  `CompressionBestCompression` when file size matters more.
- Enable `SharedStrings` when many cells repeat the same text. It reduces file
  size but keeps a string table in memory.
- For reading large files, leave shared strings on `SharedStringsAuto`; disk
  mode keeps memory flat at the cost of extra temporary I/O.
- Keep `TempDir` on a fast local disk. Network filesystems can dominate commit
  time.
- Reuse one workbook/reader per goroutine and parallelize across independent
  files.

## Migration from the writer-only release

The `v0.1.x` API remains available. New code can opt into:

- `NewWorkbookWithOptions` for temp-dir, context and compression control.
- `Row.Append`/typed `Write*` methods and `Cell.Set*` helpers.
- `Workbook.AddStyle`, `Cell.SetFormula`, `Cell.SetHyperlink`,
  `Worksheet.MergeCells`, `Worksheet.FreezePanes` and
  `Worksheet.SetAutoFilter`.
- `Open`, `Validate` and `NewWriter`.

Existing `NewWorkbook`, `AddWorksheet`, `AddCell`, `AddCellWithKey`,
`CommitRows`, `Commit` and `Cleanup` calls continue to work.

## Development

```sh
make test        # go test ./...
make test-race   # race detector
make cover       # coverage profile + function report
make cover-check # coverage gate
make interop     # openpyxl interoperability check
make examples    # run all runnable examples
make bench       # benchmarks
make lint        # golangci-lint
make check       # tidy + format + vet + race tests + lint
```

CI runs the test matrix, lint, module tidiness, coverage gate, benchmark,
fuzz, gosec, CodeQL, govulncheck and openpyxl interoperability jobs. Dependabot
keeps the Go module and GitHub Actions up to date.

## Versioning and releases

The project follows [Semantic Versioning](https://semver.org/). Versions are
regular Go module tags, so they are published to
[pkg.go.dev](https://pkg.go.dev/github.com/zignd/xlsx) automatically.

To release from the GitHub UI, run the **Tag release** workflow and choose
`patch`, `minor`, or `major` (optionally with a prerelease suffix), or provide an
exact tag such as `v1.2.3`. The workflow runs the tests, computes or validates
the tag, creates an annotated tag, and pushes it. The same workflow calls the
reusable **Release** workflow, which validates the SemVer tag, builds, tests,
scans for vulnerabilities, and creates the GitHub release with generated notes.

The same flow can be driven locally:

```sh
git tag -a v1.2.3 -m "Release v1.2.3"
git push origin v1.2.3
```

## License

MIT. See [LICENSE](LICENSE).
