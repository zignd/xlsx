// Command testgen writes a feature-rich workbook used by the interoperability
// CI job. It is not part of the public package API.
package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"time"

	"github.com/zignd/xlsx"
)

func main() {
	output := flag.String("out", "", "output .xlsx path")
	stream := flag.Bool("stream", false, "use the zero-temp streaming writer")
	flag.Parse()
	if *output == "" {
		log.Fatal("-out is required")
	}

	options := xlsx.Options{SharedStrings: true}
	var (
		workbook *xlsx.Workbook
		file     *os.File
		err      error
	)
	if *stream {
		file, err = os.Create(*output)
		if err != nil {
			log.Fatal(err)
		}
		workbook, err = xlsx.NewStreamWriter(file, options)
	} else {
		workbook, err = xlsx.NewWorkbookWithOptions(*output, options)
	}
	if err != nil {
		log.Fatal(err)
	}
	styleID, err := workbook.AddStyle(xlsx.Style{
		Bold:         true,
		NumberFormat: "#,##0.00",
		FillColor:    "FFF2CC",
		Horizontal:   "center",
	})
	if err != nil {
		log.Fatal(err)
	}

	data := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Data"})
	if err := data.FreezePanes(1, 0); err != nil {
		log.Fatal(err)
	}
	if err := data.DefineColumns([]*xlsx.WorksheetColumn{
		{Key: "name", Value: "Name", Width: 24},
		{Key: "amount", Value: "Amount"},
		{Key: "when", Value: "When"},
		{Key: "active", Value: "Active"},
	}); err != nil {
		log.Fatal(err)
	}
	if err := data.SetAutoFilter("A1:D3"); err != nil {
		log.Fatal(err)
	}

	when := time.Date(2024, 1, 2, 3, 4, 5, 0, time.UTC)
	for i := 0; i < 2; i++ {
		row, err := data.AddRow()
		if err != nil {
			log.Fatal(err)
		}
		name, _ := row.AddCellWithKey("name")
		name.Value = fmt.Sprintf("Widget & <%d>", i+1)
		name.SetHyperlink("https://example.com")
		amount, _ := row.AddCellWithKey("amount")
		amount.Value = 1.25 + float64(i)
		amount.SetStyle(styleID)
		dateCell, _ := row.AddCellWithKey("when")
		dateCell.Value = when
		active, _ := row.AddCellWithKey("active")
		active.Value = i%2 == 0
	}
	if err := data.CommitRows(); err != nil {
		log.Fatal(err)
	}
	if err := data.MergeCells("A4:B4"); err != nil {
		log.Fatal(err)
	}
	if err := data.Commit(); err != nil {
		log.Fatal(err)
	}

	summary := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Summary"})
	row, err := summary.AddRow()
	if err != nil {
		log.Fatal(err)
	}
	if err := row.Append("total", 42, true, nil); err != nil {
		log.Fatal(err)
	}
	formula, _ := row.AddCell()
	formula.SetFormula("=SUM(B1:B2)")
	if err := summary.CommitRows(); err != nil {
		log.Fatal(err)
	}

	if err := workbook.Commit(); err != nil {
		log.Fatal(err)
	}
	if file != nil {
		if err := file.Close(); err != nil {
			log.Fatal(err)
		}
	}
	if err := xlsx.Validate(*output); err != nil {
		log.Fatal(err)
	}
	fmt.Println(*output)
}
