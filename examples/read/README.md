# Read example

Validates an existing workbook and streams every worksheet row by row. Only one
row is kept in memory at a time, so the same code works for very large files.

## Run

Generate a sample workbook first:

```sh
go run ./examples/basic -out /tmp/people.xlsx
```

Then read it:

```sh
go run ./examples/read -in /tmp/people.xlsx
```

Limit the output while exploring a large file:

```sh
go run ./examples/read -in /tmp/dashboard.xlsx -max-rows 20
```

## Flags

| Flag | Default | Description |
| --- | --- | --- |
| `-in` | required | Workbook path. |
| `-max-rows` | `0` | Stop after this many data rows per sheet; `0` means all. |

## What it demonstrates

- `xlsx.Validate`
- `xlsx.Open`
- `Reader.Sheets`
- `Reader.OpenSheet`
- `SheetReader.Next`, `Row`, `Err` and `Close`
- `RowData.Cells`, `RowData.CellAt`
- `CellData.String` and typed accessors

This example works with any structurally valid `.xlsx` file, including files
created by Excel, LibreOffice, Google Sheets, openpyxl or this package. It does
not modify the input file.
