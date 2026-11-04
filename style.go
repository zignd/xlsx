package xlsx

import (
	"fmt"
	"strings"
)

const (
	defaultStyleID   = 0
	dateStyleID      = 1
	firstUserStyleID = 2
	dateNumberFormat = 164
)

// Style is a cell style. The zero value is the default workbook style.
type Style struct {
	Bold         bool
	Italic       bool
	Underline    bool
	FontSize     float64
	FontColor    string
	FillColor    string
	NumberFormat string
	Horizontal   string
	Vertical     string
	WrapText     bool
}

var builtInNumberFormats = map[string]int{
	"0":             1,
	"0.00":          2,
	"#,##0":         3,
	"#,##0.00":      4,
	"0%":            9,
	"0.00%":         10,
	"0.00E+00":      11,
	"mm-dd-yy":      14,
	"d-mmm-yy":      15,
	"d-mmm":         16,
	"mmm-yy":        17,
	"h:mm AM/PM":    18,
	"h:mm:ss AM/PM": 19,
	"h:mm":          20,
	"h:mm:ss":       21,
	"m/d/yy h:mm":   22,
	"@":             49,
}

var validHorizontalAlignments = map[string]struct{}{
	"left":             {},
	"center":           {},
	"right":            {},
	"fill":             {},
	"justify":          {},
	"centerContinuous": {},
	"distributed":      {},
}

var validVerticalAlignments = map[string]struct{}{
	"top":         {},
	"center":      {},
	"bottom":      {},
	"justify":     {},
	"distributed": {},
}

// AddStyle registers a style and returns its style ID. Style ID 0 is the
// default style and style ID 1 is reserved for dates.
func (wb *Workbook) AddStyle(style Style) (int, error) {
	if wb.committed {
		return 0, fmt.Errorf("%w: can't add styles", ErrWorkbookCommitted)
	}
	if style == (Style{}) {
		return defaultStyleID, nil
	}
	if err := validateStyle(style); err != nil {
		return 0, err
	}
	id := len(wb.styles) + firstUserStyleID
	wb.styles = append(wb.styles, style)
	return id, nil
}

func validateStyle(style Style) error {
	if style.FontSize < 0 || style.FontSize > 409 {
		return fmt.Errorf("%w: font size must be between 1 and 409", ErrInvalidStyle)
	}
	for name, color := range map[string]string{"font color": style.FontColor, "fill color": style.FillColor} {
		if color == "" {
			continue
		}
		if _, err := normalizeColor(color); err != nil {
			return fmt.Errorf("%w: %s: %w", ErrInvalidStyle, name, err)
		}
	}
	if style.NumberFormat != "" {
		if len([]rune(style.NumberFormat)) > 255 {
			return fmt.Errorf("%w: number format exceeds 255 characters", ErrInvalidStyle)
		}
		for _, r := range style.NumberFormat {
			if r < 0x20 || r == 0x7f {
				return fmt.Errorf("%w: number format contains a control character", ErrInvalidStyle)
			}
		}
	}
	if style.Horizontal != "" {
		if _, ok := validHorizontalAlignments[style.Horizontal]; !ok {
			return fmt.Errorf("%w: unknown horizontal alignment %q", ErrInvalidStyle, style.Horizontal)
		}
	}
	if style.Vertical != "" {
		if _, ok := validVerticalAlignments[style.Vertical]; !ok {
			return fmt.Errorf("%w: unknown vertical alignment %q", ErrInvalidStyle, style.Vertical)
		}
	}
	return nil
}

func normalizeColor(value string) (string, error) {
	value = strings.TrimPrefix(value, "#")
	if len(value) == 6 {
		value = "FF" + value
	}
	if len(value) != 8 {
		return "", fmt.Errorf("expected a 6- or 8-digit hex color, got %q", value)
	}
	for _, r := range value {
		if (r < '0' || r > '9') && (r < 'a' || r > 'f') && (r < 'A' || r > 'F') {
			return "", fmt.Errorf("invalid hex color %q", value)
		}
	}
	return strings.ToUpper(value), nil
}

func (wb *Workbook) styleCount() int {
	return len(wb.styles) + firstUserStyleID
}

func (wb *Workbook) validateStyleID(id int) error {
	if id < 0 || id >= wb.styleCount() {
		return fmt.Errorf("%w: style ID %d is not registered", ErrInvalidStyle, id)
	}
	return nil
}

type styleSheetBuilder struct {
	numFmts     map[string]int
	numFmtOrder []string
	nextNumFmt  int
	fonts       map[string]int
	fontXML     []string
	fills       map[string]int
	fillXML     []string
	xfXML       []string
}

func (wb *Workbook) stylesXML() string {
	builder := &styleSheetBuilder{
		numFmts:    map[string]int{dateNumberFormatCode: dateNumberFormat},
		nextNumFmt: dateNumberFormat + 1,
		fonts:      map[string]int{defaultFontKey: 0},
		fontXML:    []string{defaultFontXML},
		fills:      map[string]int{"": 0, "gray125": 1},
		fillXML:    []string{defaultFillXML, gray125FillXML},
	}

	// Style 0: default.
	builder.xfXML = append(builder.xfXML, `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`)
	// Style 1: date/time.
	builder.xfXML = append(builder.xfXML, fmt.Sprintf(
		`<xf numFmtId="%d" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`,
		dateNumberFormat,
	))

	for _, style := range wb.styles {
		fontID := builder.addFont(style)
		fillID := builder.addFill(style)
		numFmtID := builder.addNumberFormat(style.NumberFormat)
		builder.xfXML = append(builder.xfXML, builder.xf(style, numFmtID, fontID, fillID))
	}

	var b strings.Builder
	b.WriteString(xmlDeclaration)
	b.WriteString(`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`)
	fmt.Fprintf(&b, `<numFmts count="%d">`, len(builder.numFmts))
	fmt.Fprintf(&b, `<numFmt numFmtId="%d" formatCode="%s"/>`, dateNumberFormat, dateFormatCode)
	for _, format := range builder.numFmtOrder {
		id := builder.numFmts[format]
		fmt.Fprintf(&b, `<numFmt numFmtId="%d" formatCode="%s"/>`, id, xmlEscape(format))
	}
	b.WriteString(`</numFmts>`)
	fmt.Fprintf(&b, `<fonts count="%d">`, len(builder.fontXML))
	b.WriteString(strings.Join(builder.fontXML, ""))
	b.WriteString(`</fonts>`)
	fmt.Fprintf(&b, `<fills count="%d">`, len(builder.fillXML))
	b.WriteString(strings.Join(builder.fillXML, ""))
	b.WriteString(`</fills>`)
	b.WriteString(`<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>`)
	b.WriteString(`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>`)
	fmt.Fprintf(&b, `<cellXfs count="%d">`, len(builder.xfXML))
	b.WriteString(strings.Join(builder.xfXML, ""))
	b.WriteString(`</cellXfs>`)
	b.WriteString(`<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>`)
	b.WriteString(`<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>`)
	b.WriteString(`</styleSheet>`)
	return b.String()
}

const (
	dateNumberFormatCode = "yyyy-mm-dd hh:mm:ss"
	dateFormatCode       = dateNumberFormatCode
	defaultFontKey       = "default"
	defaultFontXML       = `<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>`
	defaultFillXML       = `<fill><patternFill patternType="none"/></fill>`
	gray125FillXML       = `<fill><patternFill patternType="gray125"/></fill>`
)

func (b *styleSheetBuilder) addFont(style Style) int {
	if !style.Bold && !style.Italic && !style.Underline && style.FontSize == 0 && style.FontColor == "" {
		return 0
	}

	size := style.FontSize
	if size == 0 {
		size = 11
	}
	color := style.FontColor
	if color != "" {
		normalized, _ := normalizeColor(color)
		color = normalized
	}
	key := fmt.Sprintf("%t|%t|%t|%s|%s", style.Bold, style.Italic, style.Underline, formatFloat(size), color)
	if id, ok := b.fonts[key]; ok {
		return id
	}

	var font strings.Builder
	font.WriteString(`<font>`)
	if style.Bold {
		font.WriteString(`<b/>`)
	}
	if style.Italic {
		font.WriteString(`<i/>`)
	}
	if style.Underline {
		font.WriteString(`<u/>`)
	}
	fmt.Fprintf(&font, `<sz val="%s"/>`, formatFloat(size))
	if color != "" {
		fmt.Fprintf(&font, `<color rgb="%s"/>`, color)
	} else {
		font.WriteString(`<color theme="1"/>`)
	}
	font.WriteString(`<name val="Calibri"/><family val="2"/></font>`)

	id := len(b.fontXML)
	b.fonts[key] = id
	b.fontXML = append(b.fontXML, font.String())
	return id
}

func (b *styleSheetBuilder) addFill(style Style) int {
	if style.FillColor == "" {
		return 0
	}
	normalized, _ := normalizeColor(style.FillColor)
	if id, ok := b.fills[normalized]; ok {
		return id
	}
	fill := fmt.Sprintf(`<fill><patternFill patternType="solid"><fgColor rgb="%s"/><bgColor indexed="64"/></patternFill></fill>`, normalized)
	id := len(b.fillXML)
	b.fills[normalized] = id
	b.fillXML = append(b.fillXML, fill)
	return id
}

func (b *styleSheetBuilder) addNumberFormat(format string) int {
	if format == "" {
		return 0
	}
	if id, ok := builtInNumberFormats[format]; ok {
		return id
	}
	if id, ok := b.numFmts[format]; ok {
		return id
	}
	id := b.nextNumFmt
	b.nextNumFmt++
	b.numFmts[format] = id
	b.numFmtOrder = append(b.numFmtOrder, format)
	return id
}

func (b *styleSheetBuilder) xf(style Style, numFmtID, fontID, fillID int) string {
	var attrs strings.Builder
	fmt.Fprintf(&attrs, ` numFmtId="%d" fontId="%d" fillId="%d" borderId="0" xfId="0"`, numFmtID, fontID, fillID)
	if numFmtID != 0 {
		attrs.WriteString(` applyNumberFormat="1"`)
	}
	if fontID != 0 {
		attrs.WriteString(` applyFont="1"`)
	}
	if fillID != 0 {
		attrs.WriteString(` applyFill="1"`)
	}

	alignment := b.alignment(style)
	if alignment != "" {
		attrs.WriteString(` applyAlignment="1"`)
		return `<xf` + attrs.String() + `>` + alignment + `</xf>`
	}
	return `<xf` + attrs.String() + `/>`
}

func (b *styleSheetBuilder) alignment(style Style) string {
	if style.Horizontal == "" && style.Vertical == "" && !style.WrapText {
		return ""
	}
	var attrs strings.Builder
	if style.Horizontal != "" {
		fmt.Fprintf(&attrs, ` horizontal="%s"`, style.Horizontal)
	}
	if style.Vertical != "" {
		fmt.Fprintf(&attrs, ` vertical="%s"`, style.Vertical)
	}
	if style.WrapText {
		attrs.WriteString(` wrapText="1"`)
	}
	return `<alignment` + attrs.String() + `/>`
}
