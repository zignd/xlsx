// Command http demonstrates both ways of serving an XLSX download over HTTP:
// a temp-file-backed writer for smaller reports and a zero-temp streaming
// writer for large, long-running exports.
//
//	go run ./examples/http
//	curl -o buffered.xlsx http://127.0.0.1:8080/buffered.xlsx
//	curl -o streamed.xlsx http://127.0.0.1:8080/streamed.xlsx
package main

import (
	"flag"
	"fmt"
	"log"
	"net/http"
	"time"

	"github.com/zignd/xlsx"
)

const rowsPerFlush = 500

func main() {
	addr := flag.String("addr", "127.0.0.1:8080", "HTTP listen address")
	rows := flag.Int("rows", 10000, "number of data rows per workbook")
	flag.Parse()

	mux := http.NewServeMux()
	mux.HandleFunc("/", func(w http.ResponseWriter, _ *http.Request) {
		w.Header().Set("Content-Type", "text/plain; charset=utf-8")
		fmt.Fprintln(w, "xlsx examples")
		fmt.Fprintln(w, "  /buffered.xlsx  temp-file-backed writer (materializes the archive)")
		fmt.Fprintln(w, "  /streamed.xlsx  zero-temp streaming writer (writes ZIP entries on the fly)")
	})
	mux.HandleFunc("/buffered.xlsx", func(w http.ResponseWriter, r *http.Request) {
		serveWorkbook(w, r, false, *rows)
	})
	mux.HandleFunc("/streamed.xlsx", func(w http.ResponseWriter, r *http.Request) {
		serveWorkbook(w, r, true, *rows)
	})

	server := &http.Server{
		Addr:              *addr,
		Handler:           mux,
		ReadHeaderTimeout: 5 * time.Second,
		ReadTimeout:       30 * time.Second,
		WriteTimeout:      5 * time.Minute,
		IdleTimeout:       60 * time.Second,
	}
	log.Printf("listening on http://%s", *addr)
	log.Fatal(server.ListenAndServe())
}

func serveWorkbook(w http.ResponseWriter, r *http.Request, stream bool, rows int) {
	filename := "buffered.xlsx"
	if stream {
		filename = "streamed.xlsx"
	}
	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", `attachment; filename="`+filename+`"`)
	w.Header().Set("X-Content-Type-Options", "nosniff")
	// Ask reverse proxies not to buffer the response. The stream writer also
	// flushes the http.ResponseWriter when the underlying type supports it.
	w.Header().Set("X-Accel-Buffering", "no")

	options := xlsx.Options{
		Context:        r.Context(),
		SharedStrings:  true,
		FlushEveryRows: rowsPerFlush,
	}

	var (
		workbook *xlsx.Workbook
		err      error
	)
	if stream {
		workbook, err = xlsx.NewStreamWriter(w, options)
	} else {
		workbook, err = xlsx.NewWriter(w, options)
	}
	if err != nil {
		http.Error(w, err.Error(), http.StatusInternalServerError)
		return
	}

	sheet := workbook.AddWorksheet(&xlsx.WorksheetOptions{Name: "Data"})
	if err := writeRows(sheet, rows); err != nil {
		log.Printf("failed to write %s: %v", filename, err)
		return
	}
	if err := workbook.Commit(); err != nil {
		log.Printf("failed to commit %s: %v", filename, err)
		return
	}
	log.Printf("served %s with %d rows", filename, rows)
}

func writeRows(sheet *xlsx.Worksheet, rows int) error {
	header, err := sheet.AddRow()
	if err != nil {
		return err
	}
	if err := header.Append("ID", "Name", "Amount", "Created"); err != nil {
		return err
	}
	if err := sheet.CommitRows(); err != nil {
		return err
	}

	created := time.Now()
	for i := 0; i < rows; i++ {
		row, err := sheet.AddRow()
		if err != nil {
			return err
		}
		if err := row.Append(i, fmt.Sprintf("Item %d", i), float64(i)*1.25, created.Add(time.Duration(i)*time.Second)); err != nil {
			return err
		}
		if (i+1)%rowsPerFlush == 0 {
			if err := sheet.CommitRows(); err != nil {
				return err
			}
		}
	}
	if rows%rowsPerFlush != 0 {
		return sheet.CommitRows()
	}
	return nil
}
