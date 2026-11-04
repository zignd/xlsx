// Command complex generates a feature-rich workbook.
//
// It demonstrates multiple worksheets, keyed columns, styles, formulas,
// hyperlinks, merged cells, freeze panes, auto-filters, row heights, shared
// strings, batched streaming writes, validation and reading the result back.
package main

import (
	"flag"
	"fmt"
	"log"
	"math"
	"time"

	"github.com/zignd/xlsx"
)

func main() {
	output := flag.String("out", "complex.xlsx", "path of the workbook to create")
	rowCount := flag.Int("rows", 10_000, "number of sales rows to generate")
	flag.Parse()

	if *rowCount < 1 {
		log.Fatal("-rows must be at least 1")
	}
	if err := run(*output, *rowCount); err != nil {
		log.Fatal(err)
	}
}

func run(output string, rowCount int) error {
	workbook, err := xlsx.NewWorkbookWithOptions(output, xlsx.Options{
		SharedStrings: true,
		Compression:   xlsx.CompressionBestSpeed,
	})
	if err != nil {
		return err
	}
	defer workbook.Close()

	styles, err := registerStyles(workbook)
	if err != nil {
		return err
	}

	if err := writeSalesSheet(workbook, styles, rowCount); err != nil {
		return err
	}
	if err := writeDashboard(workbook, styles, rowCount); err != nil {
		return err
	}
	if err := writeNotes(workbook, styles); err != nil {
		return err
	}

	if err := workbook.Commit(); err != nil {
		return err
	}
	if err := xlsx.Validate(output); err != nil {
		return fmt.Errorf("generated workbook is invalid: %w", err)
	}

	names, err := sheetNames(output)
	if err != nil {
		return err
	}
	fmt.Printf("wrote %s\n", output)
	fmt.Printf("worksheets: %v\n", names)
	fmt.Printf("sales data rows: %d\n", rowCount)
	return nil
}

type styleSet struct {
	currency int
	date     int
	percent  int
	integer  int
	title    int
	note     int
	link     int
}

func registerStyles(workbook *xlsx.Workbook) (styleSet, error) {
	styles := styleSet{}

	var err error
	if styles.currency, err = workbook.AddStyle(xlsx.Style{
		NumberFormat: "$#,##0.00",
		Horizontal:   "right",
	}); err != nil {
		return styles, err
	}
	if styles.date, err = workbook.AddStyle(xlsx.Style{
		NumberFormat: "yyyy-mm-dd",
		Horizontal:   "center",
	}); err != nil {
		return styles, err
	}
	if styles.percent, err = workbook.AddStyle(xlsx.Style{
		NumberFormat: "0.00%",
		Horizontal:   "right",
	}); err != nil {
		return styles, err
	}
	if styles.integer, err = workbook.AddStyle(xlsx.Style{
		NumberFormat: "#,##0",
		Horizontal:   "right",
	}); err != nil {
		return styles, err
	}
	if styles.title, err = workbook.AddStyle(xlsx.Style{
		Bold:       true,
		FontSize:   18,
		FontColor:  "FFFFFF",
		FillColor:  "1F4E78",
		Horizontal: "center",
		Vertical:   "center",
	}); err != nil {
		return styles, err
	}
	if styles.note, err = workbook.AddStyle(xlsx.Style{
		Italic:    true,
		FillColor: "FFF2CC",
		WrapText:  true,
		Vertical:  "top",
	}); err != nil {
		return styles, err
	}
	if styles.link, err = workbook.AddStyle(xlsx.Style{
		FontColor: "0563C1",
		Underline: true,
	}); err != nil {
		return styles, err
	}

	return styles, nil
}

func writeSalesSheet(workbook *xlsx.Workbook, styles styleSet, rows int) error {
	sheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Sales"})
	if err := sheet.FreezePanes(1, 0); err != nil {
		return err
	}

	if err := sheet.DefineColumns([]*xlsx.WorksheetColumn{
		{Key: "date", Value: "Date", Width: 14},
		{Key: "region", Value: "Region", Width: 14},
		{Key: "rep", Value: "Sales rep", Width: 20},
		{Key: "units", Value: "Units", Width: 10},
		{Key: "unit_price", Value: "Unit price", Width: 14},
		{Key: "revenue", Value: "Revenue", Width: 16},
		{Key: "margin", Value: "Margin %", Width: 12},
		{Key: "active", Value: "Active", Width: 10},
		{Key: "notes", Value: "Notes", Width: 34},
	}); err != nil {
		return err
	}

	lastRow := rows + 1
	if err := sheet.SetAutoFilter(fmt.Sprintf("A1:I%d", lastRow)); err != nil {
		return err
	}

	regions := []string{"North", "South", "East", "West"}
	reps := []string{"Ada", "Grace", "Alan", "Katherine", "Linus"}
	start := time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC)

	for i := 0; i < rows; i++ {
		row, err := sheet.AddRow()
		if err != nil {
			return err
		}
		if i%250 == 0 {
			if err := row.SetHeight(18); err != nil {
				return err
			}
		}

		dateCell, _ := row.AddCellWithKey("date")
		dateCell.SetTime(start.AddDate(0, 0, i%365)).SetStyle(styles.date)

		regionCell, _ := row.AddCellWithKey("region")
		regionCell.SetString(regions[i%len(regions)])

		repCell, _ := row.AddCellWithKey("rep")
		repCell.SetString(reps[i%len(reps)])

		units := int64(1 + (i*37)%250)
		unitPrice := 5 + float64((i*13)%1000)/17
		revenue := float64(units) * unitPrice
		margin := math.Max(0.03, math.Min(0.42, 0.12+float64(i%17)/100))

		unitsCell, _ := row.AddCellWithKey("units")
		unitsCell.SetInt(units).SetStyle(styles.integer)

		priceCell, _ := row.AddCellWithKey("unit_price")
		priceCell.SetFloat(unitPrice).SetStyle(styles.currency)

		revenueCell, _ := row.AddCellWithKey("revenue")
		revenueCell.SetFloat(revenue).SetStyle(styles.currency)

		marginCell, _ := row.AddCellWithKey("margin")
		marginCell.SetFloat(margin).SetStyle(styles.percent)

		activeCell, _ := row.AddCellWithKey("active")
		activeCell.SetBool(i%3 != 0)

		notesCell, _ := row.AddCellWithKey("notes")
		notesCell.SetString(fmt.Sprintf("Order %06d; region cycle %d", i+1, i%len(regions)))

		if (i+1)%500 == 0 {
			if err := sheet.CommitRows(); err != nil {
				return err
			}
		}
	}
	if rows%500 != 0 {
		if err := sheet.CommitRows(); err != nil {
			return err
		}
	}
	return nil
}

func writeDashboard(workbook *xlsx.Workbook, styles styleSet, rows int) error {
	sheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Dashboard"})
	lastRow := rows + 1

	titleRow, err := sheet.AddRow()
	if err != nil {
		return err
	}
	if err := titleRow.SetHeight(30); err != nil {
		return err
	}
	title, err := titleRow.WriteString("Sales Dashboard")
	if err != nil {
		return err
	}
	title.SetStyle(styles.title)
	if err := sheet.MergeCells("A1:D1"); err != nil {
		return err
	}

	addSummary := func(label, formula string, style int) error {
		row, err := sheet.AddRow()
		if err != nil {
			return err
		}
		if _, err := row.WriteString(label); err != nil {
			return err
		}
		cell, err := row.AddCell()
		if err != nil {
			return err
		}
		cell.SetFormula(formula).SetStyle(style)
		return nil
	}

	if err := addSummary("Total revenue", fmt.Sprintf("=SUM(Sales!F2:F%d)", lastRow), styles.currency); err != nil {
		return err
	}
	if err := addSummary("Average revenue", fmt.Sprintf("=AVERAGE(Sales!F2:F%d)", lastRow), styles.currency); err != nil {
		return err
	}
	if err := addSummary("Units sold", fmt.Sprintf("=SUM(Sales!D2:D%d)", lastRow), styles.integer); err != nil {
		return err
	}
	if err := addSummary("Active orders", fmt.Sprintf("=COUNTIF(Sales!H2:H%d,TRUE)", lastRow), styles.integer); err != nil {
		return err
	}
	if err := addSummary("Average margin", fmt.Sprintf("=AVERAGE(Sales!G2:G%d)", lastRow), styles.percent); err != nil {
		return err
	}

	linkRow, err := sheet.AddRow()
	if err != nil {
		return err
	}
	link, err := linkRow.WriteString("Jump to the Sales sheet")
	if err != nil {
		return err
	}
	link.SetHyperlink("#Sales!A1").SetStyle(styles.link)

	infoRow, err := sheet.AddRow()
	if err != nil {
		return err
	}
	if _, err := infoRow.WriteString("Generated at"); err != nil {
		return err
	}
	generated, err := infoRow.AddCell()
	if err != nil {
		return err
	}
	generated.SetTime(time.Now()).SetStyle(styles.date)

	infoRow2, err := sheet.AddRow()
	if err != nil {
		return err
	}
	if _, err := infoRow2.WriteString("Rows in Sales"); err != nil {
		return err
	}
	rowsCell, err := infoRow2.AddCell()
	if err != nil {
		return err
	}
	rowsCell.SetInt(int64(rows)).SetStyle(styles.integer)

	return sheet.CommitRows()
}

func writeNotes(workbook *xlsx.Workbook, styles styleSet) error {
	sheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Notes"})

	titleRow, err := sheet.AddRow()
	if err != nil {
		return err
	}
	if err := titleRow.SetHeight(28); err != nil {
		return err
	}
	title, err := titleRow.WriteString("About this workbook")
	if err != nil {
		return err
	}
	title.SetStyle(styles.title)
	if err := sheet.MergeCells("A1:D1"); err != nil {
		return err
	}

	noteRow, err := sheet.AddRow()
	if err != nil {
		return err
	}
	if err := noteRow.SetHeight(54); err != nil {
		return err
	}
	note, err := noteRow.WriteString("This workbook is generated by github.com/zignd/xlsx. " +
		"It uses shared strings, styles, formulas, merged cells, hyperlinks, freeze panes, " +
		"auto-filters, row heights, column widths and batched streaming writes.")
	if err != nil {
		return err
	}
	note.SetStyle(styles.note)
	if err := sheet.MergeCells("A2:D2"); err != nil {
		return err
	}

	for _, line := range []string{
		"Open the Dashboard sheet to see formulas recalculate.",
		"Open the Sales sheet to inspect the auto-filter and frozen header.",
		"Use the complex example as a starting point for your own reports.",
	} {
		row, err := sheet.AddRow()
		if err != nil {
			return err
		}
		if err := row.SetHeight(20); err != nil {
			return err
		}
		cell, err := row.WriteString("• " + line)
		if err != nil {
			return err
		}
		cell.SetStyle(styles.note)
	}

	return sheet.CommitRows()
}

func sheetNames(path string) ([]string, error) {
	reader, err := xlsx.Open(path)
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	infos := reader.Sheets()
	names := make([]string, 0, len(infos))
	for _, info := range infos {
		names = append(names, info.Name)
	}
	return names, nil
}
