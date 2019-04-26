package xlsx

import (
	"fmt"
	"time"
)

var timeLocationUTC, _ = time.LoadLocation("UTC")

func timeToUTCTime(t time.Time) time.Time {
	return time.Date(t.Year(), t.Month(), t.Day(), t.Hour(), t.Minute(), t.Second(), t.Nanosecond(), timeLocationUTC)
}

func timeToExcelTime(t time.Time) float64 {
	return float64(t.Unix())/86400.0 + 25569.0
}

const (
	startContentTypes  = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`
	endContentTypes    = "</Types>"
	rels               = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
	startWorkbook      = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9303"/><workbookPr defaultThemeVersion="124226"/><bookViews><workbookView xWindow="480" yWindow="60" windowWidth="18195" windowHeight="8505"/></bookViews><sheets>`
	endWorkbook        = `</sheets><calcPr calcId="145621"/></workbook>`
	startWorkbookRels  = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
	endWorkbookRels    = "</Relationships>"
	styles             = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><fonts count="1" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst></styleSheet>`
	startWorksheet     = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>`
	startColumns       = "<cols>"
	endColumns         = "</cols>"
	startWorksheetData = "<sheetData>"
	endRow             = "</row>"
	endWorksheetData   = "</sheetData>"
	endWorksheet       = "</worksheet>"
)

func overrideWorksheetFormat(fileName string) string {
	return fmt.Sprintf(`<Override PartName="/xl/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`, fileName)
}

func sheetFormat(name string, id int) string {
	return fmt.Sprintf(`<sheet name="%s" sheetId="%d" r:id="rId%d"/>`, name, id, id)
}

func relationshipFormat(id int, relType, target string) string {
	return fmt.Sprintf(`<Relationship Id="rId%d" Type="%s" Target="%s"/>`, id, relType, target)
}

func columnFormat(width, index int) string {
	return fmt.Sprintf(`<col min="%d" max="%d" width="%d" customWidth="1"/>`, index, index, width)
}

func startRowFormat(index int) string {
	return fmt.Sprintf(`<row r="%d">`, index)
}

func cellFormat(identifier, value string) string {
	return fmt.Sprintf(`<c r="%s" t="str"><v>%s</v></c>`, identifier, value)
}

func dateCellFormat(identifier string, value time.Time) string {
	fmtValue := timeToExcelTime(timeToUTCTime(value))
	return fmt.Sprintf(`<c r="%s" s="1" t="n"><v>%f</v></c>`, identifier, fmtValue)
}

func numberCellFormat(identifier string, value interface{}) string {
	return fmt.Sprintf(`<c r="%s" s="0" t="n"><v>%v</v></c>`, identifier, value)
}
