# Complex example

Generates a feature-rich workbook that exercises most of the writer API. It is
the best starting point if you want to adapt the package for a real report.

## Run

From the repository root:

```sh
go run ./examples/complex -out /tmp/dashboard.xlsx -rows 10000
```

`-out` defaults to `complex.xlsx`. `-rows` defaults to `10000`; increase it to
exercise the streaming writer further:

```sh
go run ./examples/complex -out /tmp/dashboard-large.xlsx -rows 250000
```

The writer flushes rows every 500 records, so memory usage stays bounded even
for large outputs.

## What it demonstrates

### Sales worksheet

- keyed columns with custom widths
- frozen header row
- auto-filter across the full data range
- batched `CommitRows` calls while streaming rows
- row heights
- date, integer, currency and percentage styles
- booleans, strings and large generated data sets

### Dashboard worksheet

- a merged title row with a custom style
- formulas: `SUM`, `AVERAGE`, `COUNTIF`
- formula results recalculated by Excel on open
- an internal hyperlink to another sheet
- generated timestamp and row count

### Notes worksheet

- merged cells
- wrapped text with a highlighted note style
- row heights and multi-line content

### Package-level features

- `xlsx.Options` with shared strings and fast compression
- `xlsx.Validate` after writing
- reading sheet metadata back with `xlsx.Open`

## Output

The program prints the output path, worksheet names and row count, then exits.
Open the file in Excel, LibreOffice or another XLSX reader to see the styled
dashboard, formulas and filters.
