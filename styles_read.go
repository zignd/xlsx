package xlsx

import (
	"context"
	"encoding/xml"
	"fmt"
	"math"
	"strings"
	"time"
)

type cellFormat struct {
	NumFmtID int `xml:"numFmtId,attr"`
}

type stylesDoc struct {
	NumFmts []struct {
		NumFmtID   int    `xml:"numFmtId,attr"`
		FormatCode string `xml:"formatCode,attr"`
	} `xml:"numFmts>numFmt"`
	CellXfs []cellFormat `xml:"cellXfs>xf"`
}

type styleTable struct {
	cellXfs []cellFormat
	numFmts map[int]string
}

func parseStyles(data []byte) (*styleTable, error) {
	var doc stylesDoc
	if err := xml.Unmarshal(data, &doc); err != nil {
		return nil, err
	}
	table := &styleTable{
		cellXfs: doc.CellXfs,
		numFmts: make(map[int]string, len(doc.NumFmts)),
	}
	for _, format := range doc.NumFmts {
		table.numFmts[format.NumFmtID] = format.FormatCode
	}
	return table, nil
}

func (t *styleTable) isDate(styleID int) bool {
	if t == nil || styleID < 0 || styleID >= len(t.cellXfs) {
		return false
	}
	numFmtID := t.cellXfs[styleID].NumFmtID
	switch {
	case numFmtID >= 14 && numFmtID <= 22:
		return true
	case numFmtID >= 27 && numFmtID <= 36:
		return true
	case numFmtID >= 45 && numFmtID <= 47:
		return true
	case numFmtID >= 50 && numFmtID <= 58:
		return true
	}
	return formatLooksLikeDate(t.numFmts[numFmtID])
}

func formatLooksLikeDate(format string) bool {
	if format == "" {
		return false
	}
	var b strings.Builder
	inQuote := false
	for i := 0; i < len(format); i++ {
		c := format[i]
		if inQuote {
			if c == '"' {
				inQuote = false
			}
			continue
		}
		switch c {
		case '"':
			inQuote = true
		case '\\':
			i++
		case '_', '*':
			i++
		case '[':
			for i < len(format) && format[i] != ']' {
				i++
			}
		default:
			b.WriteByte(c)
		}
	}
	return strings.ContainsAny(strings.ToLower(b.String()), "ymdhs")
}

func excelSerialToTime(serial float64, date1904 bool) time.Time {
	days := math.Floor(serial)
	fraction := serial - days

	var base time.Time
	switch {
	case date1904:
		base = time.Date(1904, 1, 1, 0, 0, 0, 0, time.UTC)
	case days < 60:
		// Excel's fictional 1900-02-29 means serials below 60 need a
		// different offset.
		base = time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)
	default:
		base = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	}

	nanos := int64(math.Round(fraction*86400*1e6)) * 1000
	return base.AddDate(0, 0, int(days)).Add(time.Duration(nanos))
}

func contextError(ctx context.Context) error {
	if ctx == nil {
		return nil
	}
	if err := ctx.Err(); err != nil {
		return fmt.Errorf("operation canceled: %w", err)
	}
	return nil
}
