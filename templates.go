package xlsx

import (
	"fmt"
	"strconv"
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

	rootRelationships = xmlDeclaration + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="` + contentTypeOfficeDoc + `" Target="xl/workbook.xml"/>` +
		`</Relationships>`

	workbookStart = xmlDeclaration + `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27425"/>` +
		`<workbookPr defaultThemeVersion="164011"/>` +
		`<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="16384" windowHeight="8192"/></bookViews><sheets>`
	workbookEnd = `</sheets><calcPr calcId="191029"/></workbook>`

	workbookRelsStart = xmlDeclaration + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
	workbookRelsEnd   = `</Relationships>`

	stylesXML = xmlDeclaration + `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
		`<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts>` +
		`<fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>` +
		`<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
		`<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
		`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
		`<cellXfs count="2">` +
		`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
		`<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>` +
		`</cellXfs>` +
		`<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
		`<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>` +
		`</styleSheet>`

	worksheetStart = xmlDeclaration + `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<sheetViews><sheetView workbookViewId="0"/></sheetViews>` +
		`<sheetFormatPr defaultRowHeight="15"/>`
	worksheetDataStart = `<sheetData>`
	worksheetDataEnd   = `</sheetData></worksheet>`
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

func columnXML(index int, width float64) string {
	return fmt.Sprintf(`<col min="%d" max="%d" width="%s" customWidth="1"/>`, index, index, formatFloat(width))
}

func startRowXML(index int) string {
	return fmt.Sprintf(`<row r="%d">`, index+1)
}

func endRowXML() string {
	return `</row>`
}

func textCellXML(identifier, value string) string {
	return `<c r="` + xmlEscape(identifier) + `" t="inlineStr"><is><t xml:space="preserve">` +
		xmlEscape(value) + `</t></is></c>`
}

func numberCellXML(identifier, value string) string {
	return `<c r="` + xmlEscape(identifier) + `" t="n"><v>` + value + `</v></c>`
}

func boolCellXML(identifier string, value bool) string {
	if value {
		return `<c r="` + xmlEscape(identifier) + `" t="b"><v>1</v></c>`
	}
	return `<c r="` + xmlEscape(identifier) + `" t="b"><v>0</v></c>`
}

func dateCellXML(identifier string, value time.Time) string {
	return `<c r="` + xmlEscape(identifier) + `" s="1" t="n"><v>` +
		formatFloat(timeToExcelTime(value)) + `</v></c>`
}

func blankCellXML(identifier string) string {
	return `<c r="` + xmlEscape(identifier) + `"/>`
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
