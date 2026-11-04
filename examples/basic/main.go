// Command basic demonstrates the everyday writer and reader API.
//
// It writes a keyed worksheet, validates the package, then streams it back and
// prints the values.
package main

import (
	"flag"
	"fmt"
	"log"
	"time"

	"github.com/zignd/xlsx"
)

type person struct {
	Name   string
	Age    int
	Active bool
	Joined time.Time
	Hourly float64
	Tags   string
	Notes  string
}

func main() {
	output := flag.String("out", "basic.xlsx", "path of the workbook to create")
	flag.Parse()

	if err := run(*output); err != nil {
		log.Fatal(err)
	}
}

func run(output string) error {
	workbook, err := xlsx.NewWorkbookWithOptions(output, xlsx.Options{
		SharedStrings: true,
	})
	if err != nil {
		return err
	}
	defer workbook.Close()

	worksheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "People"})
	if err := worksheet.DefineColumns([]*xlsx.WorksheetColumn{
		{Key: "name", Value: "Name", Width: 22},
		{Key: "age", Value: "Age"},
		{Key: "active", Value: "Active"},
		{Key: "joined", Value: "Joined"},
		{Key: "hourly", Value: "Hourly rate"},
		{Key: "tags", Value: "Tags"},
		{Key: "notes", Value: "Notes", Width: 30},
	}); err != nil {
		return err
	}

	people := []person{
		{
			Name:   "Ada Lovelace",
			Age:    36,
			Active: true,
			Joined: time.Date(2024, 1, 15, 9, 30, 0, 0, time.UTC),
			Hourly: 95.5,
			Tags:   "engineering, math",
			Notes:  "First programmer; likes analytical engines.",
		},
		{
			Name:   "Grace Hopper",
			Age:    85,
			Active: true,
			Joined: time.Date(2023, 11, 2, 14, 0, 0, 0, time.UTC),
			Hourly: 120,
			Tags:   "compilers, navy",
			Notes:  "Coined the term debugging.",
		},
		{
			Name:   "Alan Turing",
			Age:    41,
			Active: false,
			Joined: time.Date(2022, 6, 23, 8, 15, 0, 0, time.UTC),
			Hourly: 88.25,
			Tags:   "cryptography, ai",
			Notes:  "Foundational work on computability.",
		},
	}

	for _, p := range people {
		row, err := worksheet.AddRow()
		if err != nil {
			return err
		}
		if err := row.Append(p.Name, p.Age, p.Active, p.Joined, p.Hourly, p.Tags, p.Notes); err != nil {
			return err
		}
		if err := worksheet.CommitRows(); err != nil {
			return err
		}
	}

	if err := worksheet.Commit(); err != nil {
		return err
	}
	if err := workbook.Commit(); err != nil {
		return err
	}

	if err := xlsx.Validate(output); err != nil {
		return fmt.Errorf("generated workbook is invalid: %w", err)
	}
	fmt.Printf("wrote %s\n", output)

	reader, err := xlsx.Open(output)
	if err != nil {
		return err
	}
	defer reader.Close()

	fmt.Println("sheets:", sheetNames(reader))

	sheet, err := reader.OpenSheet("People")
	if err != nil {
		return err
	}
	defer sheet.Close()

	for sheet.Next() {
		row := sheet.Row()
		if row.Index == 1 {
			continue // header
		}
		name, _ := row.CellAt(0)
		age, _ := row.CellAt(1)
		active, _ := row.CellAt(2)

		nameValue, _ := name.String()
		ageValue, _ := age.Int64()
		activeValue, _ := active.Bool()
		fmt.Printf("%s | age %d | active %t\n", nameValue, ageValue, activeValue)
	}
	if err := sheet.Err(); err != nil {
		return err
	}
	return nil
}

func sheetNames(reader *xlsx.Reader) []string {
	infos := reader.Sheets()
	names := make([]string, 0, len(infos))
	for _, info := range infos {
		names = append(names, info.Name)
	}
	return names
}
