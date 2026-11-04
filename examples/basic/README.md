# Basic example

Writes a small keyed workbook, validates it, then streams it back and prints
the values.

## Run

From the repository root:

```sh
go run ./examples/basic -out /tmp/people.xlsx
```

`-out` defaults to `basic.xlsx` in the current working directory. Use an
absolute path if you run the command from another directory.

## What it demonstrates

- `xlsx.NewWorkbookWithOptions` with shared strings enabled
- `Worksheet.DefineColumns` with column widths
- `Row.Append` with strings, integers, booleans, floats, dates and text
- `Worksheet.CommitRows` and `Workbook.Commit`
- `xlsx.Validate`
- `xlsx.Open`, `Reader.OpenSheet`, `SheetReader.Next`
- typed cell accessors (`String`, `Int64`, `Bool`)

## Expected output

The exact timestamps are stable because the example uses fixed dates. The
values printed at the end look like:

```text
wrote /tmp/people.xlsx
sheets: [People]
Ada Lovelace | age 36 | active true
Grace Hopper | age 85 | active true
Alan Turing | age 41 | active false
```

Open the generated file in Excel, LibreOffice or another XLSX reader to inspect
the header, column widths and typed values.
