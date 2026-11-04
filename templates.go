package xlsx

import (
	"fmt"
	"strconv"
	"strings"
	"time"
)

const (
	xmlDeclaration = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`

	contentTypesStart = xmlDeclaration + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
		`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
		`<Default Extension="xml" ContentType="application/xml"/>`
	contentTypesEnd = `</Types>`

	contentTypeWorkbook   = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
	contentTypeWorksheet  = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
	contentTypeStyles     = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
	contentTypeOfficeDoc  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	contentTypeWorksheetR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
	contentTypeStylesR    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
	contentTypeHyperlinkR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

	rootRelationships = xmlDeclaration + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="` + contentTypeOfficeDoc + `" Target="xl/workbook.xml"/>` +
		`</Relationships>`

	workbookStart = xmlDeclaration + `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27425"/>` +
		`<workbookPr defaultThemeVersion="164011"/>` +
		`<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="16384" windowHeight="8192"/></bookViews><sheets>`
	workbookEnd = `</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`

	workbookRelsStart = xmlDeclaration + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
	workbookRelsEnd   = `</Relationships>`

	worksheetStartPrefix = xmlDeclaration + `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`
	worksheetDataStart = `<sheetData>`
	worksheetDataEnd   = `</sheetData>`
	worksheetEnd       = `</worksheet>`

	worksheetRelsStart = workbookRelsStart
	worksheetRelsEnd   = workbookRelsEnd
)

func contentTypeOverride(partName, contentType string) string {
	return `<Override PartName="` + xmlEscape(partName) + `" ContentType="` + xmlEscape(contentType) + `"/>`
}

func sheetXML(name string, id int) string {
	return fmt.Sprintf(`<sheet name="%s" sheetId="%d" r:id="rId%d"/>`, xmlEscape(name), id, id)
}

func relationshipXML(id int, relType, target string) string {
	return fmt.Sprintf(`<Relationship Id="rId%d" Type="%s" Target="%s"/>`, id, xmlEscape(relType), xmlEscape(target))
}

func externalRelationshipXML(id int, relType, target string) string {
	return fmt.Sprintf(`<Relationship Id="rId%d" Type="%s" Target="%s" TargetMode="External"/>`, id, xmlEscape(relType), xmlEscape(target))
}

func columnXML(index int, width float64) string {
	return fmt.Sprintf(`<col min="%d" max="%d" width="%s" customWidth="1"/>`, index, index, formatFloat(width))
}

func worksheetViewsXML(freezeRows, freezeCols int) string {
	if freezeRows == 0 && freezeCols == 0 {
		return `<sheetViews><sheetView workbookViewId="0"/></sheetViews>`
	}

	activePane := "topRight"
	if freezeRows > 0 && freezeCols > 0 {
		activePane = "bottomRight"
	} else if freezeRows > 0 {
		activePane = "bottomLeft"
	}

	topLeft := createIdentifierFromCoords(freezeCols, freezeRows)
	var pane strings.Builder
	pane.WriteString(`<pane`)
	if freezeCols > 0 {
		fmt.Fprintf(&pane, ` xSplit="%d"`, freezeCols)
	}
	if freezeRows > 0 {
		fmt.Fprintf(&pane, ` ySplit="%d"`, freezeRows)
	}
	fmt.Fprintf(&pane, ` topLeftCell="%s" activePane="%s" state="frozen"/>`, topLeft, activePane)
	fmt.Fprintf(&pane, `<selection pane="%s" activeCell="%s" sqref="%s"/>`, activePane, topLeft, topLeft)
	return `<sheetViews><sheetView workbookViewId="0">` + pane.String() + `</sheetView></sheetViews>`
}

func startRowXML(index int, height float64) string {
	if height > 0 {
		return fmt.Sprintf(`<row r="%d" ht="%s" customHeight="1">`, index+1, formatFloat(height))
	}
	return fmt.Sprintf(`<row r="%d">`, index+1)
}

func endRowXML() string {
	return `</row>`
}

func styleAttribute(styleID int) string {
	if styleID == 0 {
		return ""
	}
	return ` s="` + strconv.Itoa(styleID) + `"`
}

func textCellXML(identifier, value string, styleID int) string {
	return `<c r="` + xmlEscape(identifier) + `"` + styleAttribute(styleID) + ` t="inlineStr"><is><t xml:space="preserve">` +
		xmlEscape(value) + `</t></is></c>`
}

func sharedStringCellXML(identifier string, index, styleID int) string {
	return `<c r="` + xmlEscape(identifier) + `"` + styleAttribute(styleID) + ` t="s"><v>` +
		strconv.Itoa(index) + `</v></c>`
}

func numberCellXML(identifier, value string, styleID int) string {
	return `<c r="` + xmlEscape(identifier) + `"` + styleAttribute(styleID) + ` t="n"><v>` + value + `</v></c>`
}

func boolCellXML(identifier string, value bool, styleID int) string {
	if value {
		return `<c r="` + xmlEscape(identifier) + `"` + styleAttribute(styleID) + ` t="b"><v>1</v></c>`
	}
	return `<c r="` + xmlEscape(identifier) + `"` + styleAttribute(styleID) + ` t="b"><v>0</v></c>`
}

func dateCellXML(identifier string, value time.Time, styleID int) string {
	if styleID == 0 {
		styleID = dateStyleID
	}
	return `<c r="` + xmlEscape(identifier) + `" s="` + strconv.Itoa(styleID) + `" t="n"><v>` +
		formatFloat(timeToExcelTime(value)) + `</v></c>`
}

func formulaCellXML(identifier, formula string, styleID int) string {
	return `<c r="` + xmlEscape(identifier) + `"` + styleAttribute(styleID) + `><f>` +
		xmlEscape(formula) + `</f></c>`
}

func blankCellXML(identifier string, styleID int) string {
	return `<c r="` + xmlEscape(identifier) + `"` + styleAttribute(styleID) + `/>`
}

func mergeCellsXML(ranges []string) string {
	if len(ranges) == 0 {
		return ""
	}
	var b strings.Builder
	fmt.Fprintf(&b, `<mergeCells count="%d">`, len(ranges))
	for _, ref := range ranges {
		b.WriteString(`<mergeCell ref="` + xmlEscape(ref) + `"/>`)
	}
	b.WriteString(`</mergeCells>`)
	return b.String()
}

func hyperlinksXML(links []hyperlink) string {
	if len(links) == 0 {
		return ""
	}
	var b strings.Builder
	b.WriteString(`<hyperlinks>`)
	for i, link := range links {
		fmt.Fprintf(&b, `<hyperlink ref="%s" r:id="rId%d"/>`, xmlEscape(link.ref), i+1)
	}
	b.WriteString(`</hyperlinks>`)
	return b.String()
}

func autoFilterXML(ref string) string {
	if ref == "" {
		return ""
	}
	return `<autoFilter ref="` + xmlEscape(ref) + `"/>`
}

func formatFloat(value float64) string {
	return strconv.FormatFloat(value, 'f', -1, 64)
}

func formatFloat32(value float32) string {
	return strconv.FormatFloat(float64(value), 'f', -1, 32)
}

var timeLocationUTC = time.UTC

// timeToUTCTime converts a time to a UTC time while preserving its wall-clock
// components. Excel stores naive local date-times, therefore the location is
// intentionally discarded.
func timeToUTCTime(t time.Time) time.Time {
	return time.Date(t.Year(), t.Month(), t.Day(), t.Hour(), t.Minute(), t.Second(), t.Nanosecond(), timeLocationUTC)
}

// timeToExcelTime converts a time to an Excel serial date. The epoch accounts
// for Excel's 1900 date system (including its historical leap-year bug for
// dates before 1900-03-01, which is outside the supported range of this
// package).
func timeToExcelTime(t time.Time) float64 {
	t = timeToUTCTime(t)
	seconds := float64(t.Unix()) + float64(t.Nanosecond())/1e9
	return seconds/86400.0 + 25569.0
}
