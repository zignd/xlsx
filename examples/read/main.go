// Command read demonstrates validation and memory-flat reading.
//
// It validates the package structure, lists every worksheet and streams each
// sheet row by row. Only one row is kept in memory at a time.
package main

import (
	"flag"
	"fmt"
	"log"
	"strings"

	"github.com/zignd/xlsx"
)

func main() {
	input := flag.String("in", "", "path of the workbook to read")
	maxRows := flag.Int("max-rows", 0, "stop after this many data rows per sheet (0 means all)")
	flag.Parse()

	if *input == "" {
		log.Fatal("-in is required; try: go run ./examples/read -in /tmp/people.xlsx")
	}
	if err := run(*input, *maxRows); err != nil {
		log.Fatal(err)
	}
}

func run(input string, maxRows int) error {
	if err := xlsx.Validate(input); err != nil {
		return fmt.Errorf("invalid workbook: %w", err)
	}
	fmt.Println("validation: ok")

	reader, err := xlsx.Open(input)
	if err != nil {
		return err
	}
	defer reader.Close()

	for _, info := range reader.Sheets() {
		fmt.Printf("\n=== %s (hidden=%t) ===\n", info.Name, info.Hidden)

		sheet, err := reader.OpenSheet(info.Name)
		if err != nil {
			return err
		}

		rows := 0
		for sheet.Next() {
			row := sheet.Row()
			if row.Index == 1 {
				continue // header
			}
			if maxRows > 0 && rows >= maxRows {
				fmt.Printf("... stopping after %d data rows\n", maxRows)
				break
			}
			rows++

			values := make([]string, 0, len(row.Cells))
			for i := range row.Cells {
				cell := &row.Cells[i]
				value, err := cell.String()
				if err != nil {
					value = fmt.Sprintf("<%v>", err)
				}
				values = append(values, fmt.Sprintf("%s=%q", cell.Ref, value))
			}
			fmt.Printf("row %d: %s\n", row.Index, strings.Join(values, " "))
		}
		if err := sheet.Err(); err != nil {
			_ = sheet.Close()
			return err
		}
		if err := sheet.Close(); err != nil {
			return err
		}
		fmt.Printf("read %d data rows from %s\n", rows, info.Name)
	}
	return nil
}
