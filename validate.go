package xlsx

import (
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"path"
	"strings"
	"unicode/utf8"
)

// ValidateOption customizes Validate and ValidateReaderAt.
type ValidateOption = ReadOption

// Validate checks that a file is a structurally valid XLSX package and that
// its worksheets can be streamed by Reader. It does not recalculate formulas.
func Validate(filePath string, opts ...ValidateOption) error {
	reader, err := Open(filePath, opts...)
	if err != nil {
		return err
	}
	defer reader.Close()
	return validateReader(reader)
}

// ValidateReaderAt validates an XLSX package from an io.ReaderAt.
func ValidateReaderAt(r io.ReaderAt, size int64, opts ...ValidateOption) error {
	reader, err := OpenReaderAt(r, size, opts...)
	if err != nil {
		return err
	}
	defer reader.Close()
	return validateReader(reader)
}

func validateReader(reader *Reader) error {
	var problems []error
	add := func(err error) {
		if err != nil {
			problems = append(problems, err)
		}
	}

	add(validateContentTypes(reader))
	add(validateRootRelationships(reader))
	add(validateWorkbookStructure(reader))
	add(validateWorkbookRelationships(reader))
	add(validateWorksheetRelationships(reader))

	for _, sheet := range reader.sheets {
		sheetReader, err := reader.OpenSheet(sheet.Name)
		if err != nil {
			add(fmt.Errorf("worksheet %q: %w", sheet.Name, err))
			continue
		}
		previousRow := 0
		rows := 0
		for sheetReader.Next() {
			rows++
			row := sheetReader.Row()
			if row.Index <= previousRow {
				add(fmt.Errorf("worksheet %q row %d is not in ascending order", sheet.Name, row.Index))
			}
			previousRow = row.Index
			previousColumn := -1
			for i := range row.Cells {
				cell := &row.Cells[i]
				column, cellRow, err := parseCellReference(cell.Ref)
				if err != nil {
					add(fmt.Errorf("worksheet %q cell %q: %w", sheet.Name, cell.Ref, err))
					continue
				}
				if column <= previousColumn {
					add(fmt.Errorf("worksheet %q row %d cell %q is not in ascending column order", sheet.Name, row.Index, cell.Ref))
				}
				previousColumn = column
				if cellRow != row.Index {
					add(fmt.Errorf("worksheet %q cell %q belongs to row %d, not %d", sheet.Name, cell.Ref, cellRow, row.Index))
				}
				if cell.Column != column || cell.Row != cellRow {
					add(fmt.Errorf("worksheet %q cell %q has inconsistent coordinates", sheet.Name, cell.Ref))
				}
				if reader.styles == nil || cell.StyleID < 0 || cell.StyleID >= len(reader.styles.cellXfs) {
					if cell.StyleID != 0 {
						add(fmt.Errorf("worksheet %q cell %q references missing style %d", sheet.Name, cell.Ref, cell.StyleID))
					}
				}
				if utf8.RuneCountInString(cell.Raw) > reader.options.MaxCellLength {
					add(fmt.Errorf("worksheet %q cell %q exceeds the cell length limit", sheet.Name, cell.Ref))
				}
				if len([]rune(cell.Formula)) > maxFormulaLength {
					add(fmt.Errorf("worksheet %q cell %q has an overlong formula", sheet.Name, cell.Ref))
				}
			}
		}
		add(sheetReader.Err())
		add(sheetReader.Close())
		if rows > reader.options.MaxRows {
			add(fmt.Errorf("worksheet %q exceeds %d rows", sheet.Name, reader.options.MaxRows))
		}
	}

	if len(problems) == 0 {
		return nil
	}
	return errors.Join(append([]error{ErrCorruptFile}, problems...)...)
}

type contentTypesDoc struct {
	Defaults []struct {
		Extension   string `xml:"Extension,attr"`
		ContentType string `xml:"ContentType,attr"`
	} `xml:"Default"`
	Overrides []struct {
		PartName    string `xml:"PartName,attr"`
		ContentType string `xml:"ContentType,attr"`
	} `xml:"Override"`
}

func validateContentTypes(reader *Reader) error {
	file, ok := reader.entries["[Content_Types].xml"]
	if !ok {
		return errors.New("[Content_Types].xml is missing")
	}
	data, err := readZipFile(file)
	if err != nil {
		return fmt.Errorf("failed to read [Content_Types].xml: %w", err)
	}
	var doc contentTypesDoc
	if err := xml.Unmarshal(data, &doc); err != nil {
		return fmt.Errorf("[Content_Types].xml is not well-formed: %w", err)
	}

	overrides := make(map[string]string, len(doc.Overrides))
	for _, override := range doc.Overrides {
		partName := strings.TrimPrefix(override.PartName, "/")
		overrides[partName] = override.ContentType
		if _, exists := reader.entries[partName]; !exists {
			return fmt.Errorf("[Content_Types].xml references missing part %q", partName)
		}
	}
	if _, ok := overrides[reader.workbookPath]; !ok {
		return fmt.Errorf("[Content_Types].xml has no override for %q", reader.workbookPath)
	}
	if contentType := overrides[reader.workbookPath]; contentType != contentTypeWorkbook {
		return fmt.Errorf("workbook content type is %q, want %q", contentType, contentTypeWorkbook)
	}

	hasXMLDefault := false
	for _, def := range doc.Defaults {
		if strings.EqualFold(def.Extension, "xml") {
			hasXMLDefault = true
		}
	}
	if !hasXMLDefault {
		return errors.New("[Content_Types].xml has no xml default type")
	}

	for _, sheet := range reader.sheets {
		contentType, ok := overrides[sheet.Path]
		if !ok {
			// The XML default covers worksheet parts. That is valid OPC, but
			// Excel normally writes an override. Accept it.
			continue
		}
		if contentType != contentTypeWorksheet {
			return fmt.Errorf("worksheet %q content type is %q, want %q", sheet.Path, contentType, contentTypeWorksheet)
		}
	}
	if _, ok := reader.entries[path.Join(path.Dir(reader.workbookPath), "styles.xml")]; ok {
		stylesPath := path.Join(path.Dir(reader.workbookPath), "styles.xml")
		contentType, ok := overrides[stylesPath]
		if !ok {
			return fmt.Errorf("[Content_Types].xml has no override for %q", stylesPath)
		}
		if contentType != contentTypeStyles {
			return fmt.Errorf("styles content type is %q, want %q", contentType, contentTypeStyles)
		}
	}
	sharedPath := path.Join(path.Dir(reader.workbookPath), "sharedStrings.xml")
	if _, exists := reader.entries[sharedPath]; exists {
		contentType, ok := overrides[sharedPath]
		if !ok {
			return fmt.Errorf("[Content_Types].xml has no override for %q", sharedPath)
		}
		if contentType != contentTypeSharedStrings {
			return fmt.Errorf("shared strings content type is %q, want %q", contentType, contentTypeSharedStrings)
		}
	}
	return nil
}

func validateRootRelationships(reader *Reader) error {
	file, ok := reader.entries["_rels/.rels"]
	if !ok {
		return errors.New("_rels/.rels is missing")
	}
	data, err := readZipFile(file)
	if err != nil {
		return fmt.Errorf("failed to read _rels/.rels: %w", err)
	}
	rels, err := parseRelationships(data)
	if err != nil {
		return fmt.Errorf("_rels/.rels is not well-formed: %w", err)
	}
	found := false
	for _, rel := range rels {
		if strings.HasSuffix(rel.Type, "/officeDocument") {
			found = true
			target := resolvePartPath("", rel.Target)
			if _, exists := reader.entries[target]; !exists {
				return fmt.Errorf("root relationship target %q is missing", target)
			}
		}
	}
	if !found {
		return errors.New("_rels/.rels has no officeDocument relationship")
	}
	return nil
}

func validateWorkbookStructure(reader *Reader) error {
	seen := make(map[string]struct{}, len(reader.sheets))
	seenIDs := make(map[int]struct{}, len(reader.sheets))
	for _, sheet := range reader.sheets {
		if err := validateSheetName(sheet.Name); err != nil {
			return err
		}
		key := strings.ToLower(sheet.Name)
		if _, exists := seen[key]; exists {
			return fmt.Errorf("%w: %q", ErrDuplicateSheetName, sheet.Name)
		}
		seen[key] = struct{}{}
		if _, exists := seenIDs[sheet.ID]; exists {
			return fmt.Errorf("duplicate sheet ID %d", sheet.ID)
		}
		seenIDs[sheet.ID] = struct{}{}
	}
	return nil
}

func validateWorkbookRelationships(reader *Reader) error {
	relsPath := path.Join(path.Dir(reader.workbookPath), "_rels", path.Base(reader.workbookPath)+".rels")
	data, err := reader.readPart(relsPath)
	if err != nil {
		return err
	}
	rels, err := parseRelationships(data)
	if err != nil {
		return fmt.Errorf("workbook relationships are not well-formed: %w", err)
	}
	for _, rel := range rels {
		if strings.EqualFold(rel.TargetMode, "External") {
			continue
		}
		target := resolvePartPath(reader.workbookPath, rel.Target)
		if _, exists := reader.entries[target]; !exists {
			return fmt.Errorf("workbook relationship target %q is missing", target)
		}
	}
	return nil
}

func validateWorksheetRelationships(reader *Reader) error {
	for _, sheet := range reader.sheets {
		relsPath := path.Join(path.Dir(sheet.Path), "_rels", path.Base(sheet.Path)+".rels")
		file, ok := reader.entries[relsPath]
		if !ok {
			continue
		}
		data, err := readZipFile(file)
		if err != nil {
			return fmt.Errorf("failed to read %s: %w", relsPath, err)
		}
		rels, err := parseRelationships(data)
		if err != nil {
			return fmt.Errorf("%s is not well-formed: %w", relsPath, err)
		}
		for _, rel := range rels {
			if strings.EqualFold(rel.TargetMode, "External") || strings.HasPrefix(rel.Target, "#") {
				continue
			}
			target := resolvePartPath(sheet.Path, rel.Target)
			if _, exists := reader.entries[target]; !exists {
				return fmt.Errorf("worksheet relationship target %q is missing", target)
			}
		}
	}
	return nil
}
