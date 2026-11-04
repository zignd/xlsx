package xlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"path"
	"strings"
)

var oleSignature = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

// SheetInfo describes a worksheet in a workbook.
type SheetInfo struct {
	Name   string
	ID     int
	RelID  string
	Path   string
	State  string
	Hidden bool
}

// Reader is a streaming XLSX reader.
type Reader struct {
	zr           *zip.Reader
	entries      map[string]*zip.File
	workbookPath string
	sheets       []SheetInfo
	shared       *sharedStrings
	styles       *styleTable
	date1904     bool
	options      ReadOptions
	closer       io.Closer
	closed       bool
}

type relationshipInfo struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr"`
}

type relationshipsDoc struct {
	Items []relationshipInfo `xml:"Relationship"`
}

func sniffZipFormat(r io.ReaderAt, size int64) error {
	if size < 4 {
		return fmt.Errorf("%w: file is too small", ErrCorruptFile)
	}
	var header [8]byte
	n, err := r.ReadAt(header[:], 0)
	if err != nil && !errors.Is(err, io.EOF) {
		return fmt.Errorf("%w: %w", ErrCorruptFile, err)
	}
	if n >= len(oleSignature) && bytes.Equal(header[:len(oleSignature)], oleSignature) {
		return fmt.Errorf("%w: %w: OLE compound file (encrypted workbook or legacy .xls)", ErrUnsupportedFormat, ErrEncrypted)
	}
	if n < 4 || header[0] != 'P' || header[1] != 'K' {
		return fmt.Errorf("%w: missing ZIP signature", ErrCorruptFile)
	}
	return nil
}

func newReader(r io.ReaderAt, size int64, closer io.Closer, options ReadOptions) (*Reader, error) {
	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, fmt.Errorf("%w: %w", ErrCorruptFile, err)
	}

	reader := &Reader{
		zr:      zr,
		entries: make(map[string]*zip.File, len(zr.File)),
		options: options,
		closer:  closer,
	}

	var total uint64
	for _, entry := range zr.File {
		if err := reader.checkContext(); err != nil {
			_ = reader.Close()
			return nil, err
		}
		if len(reader.entries) >= options.MaxEntries {
			_ = reader.Close()
			return nil, fmt.Errorf("%w: more than %d ZIP entries", ErrReadLimit, options.MaxEntries)
		}
		if !isSafeZipName(entry.Name) {
			_ = reader.Close()
			return nil, fmt.Errorf("%w: unsafe ZIP entry name %q", ErrCorruptFile, entry.Name)
		}
		if _, exists := reader.entries[entry.Name]; exists {
			_ = reader.Close()
			return nil, fmt.Errorf("%w: duplicate ZIP entry %q", ErrCorruptFile, entry.Name)
		}
		// #nosec G115 -- read options are validated as non-negative.
		maxSize := uint64(options.MaxUncompressedSize)
		if entry.UncompressedSize64 > maxSize || total > maxSize-entry.UncompressedSize64 {
			_ = reader.Close()
			return nil, fmt.Errorf("%w: uncompressed package exceeds %d bytes", ErrReadLimit, options.MaxUncompressedSize)
		}
		total += entry.UncompressedSize64
		reader.entries[entry.Name] = entry

		switch entry.Name {
		case "EncryptedPackage", "EncryptionInfo":
			_ = reader.Close()
			return nil, ErrEncrypted
		case "xl/workbook.bin":
			_ = reader.Close()
			return nil, fmt.Errorf("%w: XLSB workbook", ErrUnsupportedFormat)
		}
	}

	workbookPath, err := findWorkbookPath(reader.entries)
	if err != nil {
		_ = reader.Close()
		return nil, err
	}
	reader.workbookPath = workbookPath

	contentTypesData, err := reader.readPart("[Content_Types].xml")
	if err != nil {
		_ = reader.Close()
		return nil, fmt.Errorf("%w: [Content_Types].xml: %w", ErrCorruptFile, err)
	}
	var contentTypes contentTypesDoc
	if err := xml.Unmarshal(contentTypesData, &contentTypes); err != nil {
		_ = reader.Close()
		return nil, fmt.Errorf("%w: [Content_Types].xml: %w", ErrCorruptFile, err)
	}

	workbookData, err := reader.readPart(workbookPath)
	if err != nil {
		_ = reader.Close()
		return nil, err
	}
	var workbook workbookDoc
	if err := xml.Unmarshal(workbookData, &workbook); err != nil {
		_ = reader.Close()
		return nil, fmt.Errorf("%w: workbook.xml: %w", ErrCorruptFile, err)
	}
	reader.date1904 = workbook.WorkbookPr.Date1904

	relsPath := path.Join(path.Dir(workbookPath), "_rels", path.Base(workbookPath)+".rels")
	relsData, err := reader.readPart(relsPath)
	if err != nil {
		_ = reader.Close()
		return nil, fmt.Errorf("%w: workbook relationships: %w", ErrCorruptFile, err)
	}
	rels, err := parseRelationships(relsData)
	if err != nil {
		_ = reader.Close()
		return nil, fmt.Errorf("%w: workbook relationships: %w", ErrCorruptFile, err)
	}

	for _, sheet := range workbook.Sheets {
		rel, ok := rels[sheet.RelID]
		if !ok {
			_ = reader.Close()
			return nil, fmt.Errorf("%w: worksheet %q has no relationship %q", ErrCorruptFile, sheet.Name, sheet.RelID)
		}
		if strings.EqualFold(rel.TargetMode, "External") {
			_ = reader.Close()
			return nil, fmt.Errorf("%w: worksheet %q is external", ErrCorruptFile, sheet.Name)
		}
		sheetPath := resolvePartPath(workbookPath, rel.Target)
		if _, ok := reader.entries[sheetPath]; !ok {
			_ = reader.Close()
			return nil, fmt.Errorf("%w: worksheet part %q is missing", ErrCorruptFile, sheetPath)
		}
		state := sheet.State
		if state == "" {
			state = "visible"
		}
		reader.sheets = append(reader.sheets, SheetInfo{
			Name:   sheet.Name,
			ID:     sheet.SheetID,
			RelID:  sheet.RelID,
			Path:   sheetPath,
			State:  state,
			Hidden: state == "hidden" || state == "veryHidden",
		})
	}
	if len(reader.sheets) == 0 {
		_ = reader.Close()
		return nil, fmt.Errorf("%w: workbook has no worksheets", ErrCorruptFile)
	}

	stylesPath := resolvePartPath(workbookPath, "styles.xml")
	if _, ok := reader.entries[stylesPath]; ok {
		stylesData, err := reader.readPart(stylesPath)
		if err != nil {
			_ = reader.Close()
			return nil, err
		}
		reader.styles, err = parseStyles(stylesData)
		if err != nil {
			_ = reader.Close()
			return nil, fmt.Errorf("%w: styles.xml: %w", ErrCorruptFile, err)
		}
	} else {
		reader.styles = &styleTable{}
	}

	sharedPath := resolvePartPath(workbookPath, "sharedStrings.xml")
	reader.shared, err = openSharedStrings(reader.entries, sharedPath, options)
	if err != nil {
		_ = reader.Close()
		return nil, err
	}

	return reader, nil
}

type workbookDoc struct {
	WorkbookPr struct {
		Date1904 bool `xml:"date1904,attr"`
	} `xml:"workbookPr"`
	Sheets []struct {
		Name    string `xml:"name,attr"`
		SheetID int    `xml:"sheetId,attr"`
		RelID   string `xml:"id,attr"`
		State   string `xml:"state,attr"`
	} `xml:"sheets>sheet"`
}

func findWorkbookPath(entries map[string]*zip.File) (string, error) {
	rootRels, ok := entries["_rels/.rels"]
	if ok {
		data, err := readZipFile(rootRels)
		if err != nil {
			return "", fmt.Errorf("%w: root relationships: %w", ErrCorruptFile, err)
		}
		rels, err := parseRelationships(data)
		if err != nil {
			return "", fmt.Errorf("%w: root relationships: %w", ErrCorruptFile, err)
		}
		for _, rel := range rels {
			if strings.HasSuffix(rel.Type, "/officeDocument") && !strings.EqualFold(rel.TargetMode, "External") {
				target := path.Clean(strings.TrimPrefix(rel.Target, "/"))
				if _, exists := entries[target]; exists {
					return target, nil
				}
				return "", fmt.Errorf("%w: workbook part %q is missing", ErrCorruptFile, target)
			}
		}
	}
	if _, ok := entries["xl/workbook.xml"]; ok {
		return "xl/workbook.xml", nil
	}
	return "", fmt.Errorf("%w: xl/workbook.xml is missing", ErrCorruptFile)
}

func parseRelationships(data []byte) (map[string]relationshipInfo, error) {
	var doc relationshipsDoc
	if err := xml.Unmarshal(data, &doc); err != nil {
		return nil, err
	}
	rels := make(map[string]relationshipInfo, len(doc.Items))
	for _, item := range doc.Items {
		rels[item.ID] = item
	}
	return rels, nil
}

func resolvePartPath(basePart, target string) string {
	target = strings.TrimSpace(target)
	if strings.HasPrefix(target, "/") {
		return path.Clean(strings.TrimPrefix(target, "/"))
	}
	return path.Clean(path.Join(path.Dir(basePart), target))
}

func isSafeZipName(name string) bool {
	if name == "" || strings.HasPrefix(name, "/") || strings.Contains(name, "\\") {
		return false
	}
	if strings.ContainsRune(name, '\x00') {
		return false
	}
	for _, part := range strings.Split(name, "/") {
		if part == ".." {
			return false
		}
	}
	return true
}

func readZipFile(file *zip.File) ([]byte, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()
	return io.ReadAll(rc)
}

func (r *Reader) readPart(name string) ([]byte, error) {
	file, ok := r.entries[name]
	if !ok {
		return nil, fmt.Errorf("part %q is missing", name)
	}
	data, err := readZipFile(file)
	if err != nil {
		return nil, fmt.Errorf("failed to read part %q: %w", name, err)
	}
	return data, nil
}

func (r *Reader) checkContext() error {
	if r.options.Context == nil {
		return nil
	}
	return r.options.Context.Err()
}

// Sheets returns a copy of the workbook's sheet metadata.
func (r *Reader) Sheets() []SheetInfo {
	sheets := make([]SheetInfo, len(r.sheets))
	copy(sheets, r.sheets)
	return sheets
}

// OpenSheet opens a worksheet for streaming reads.
func (r *Reader) OpenSheet(name string) (*SheetReader, error) {
	if r.closed {
		return nil, fmt.Errorf("%w: reader is closed", ErrCorruptFile)
	}
	if err := r.checkContext(); err != nil {
		return nil, err
	}
	for _, sheet := range r.sheets {
		if sheet.Name != name {
			continue
		}
		entry, ok := r.entries[sheet.Path]
		if !ok {
			return nil, fmt.Errorf("%w: worksheet part %q is missing", ErrCorruptFile, sheet.Path)
		}
		rc, err := entry.Open()
		if err != nil {
			return nil, fmt.Errorf("%w: failed to open %q: %w", ErrCorruptFile, sheet.Path, err)
		}
		return &SheetReader{
			reader: r,
			info:   sheet,
			rc:     rc,
			dec:    xml.NewDecoder(rc),
		}, nil
	}
	return nil, fmt.Errorf("%w: %q", ErrSheetNotFound, name)
}

// Close releases the package and any temporary reader files.
func (r *Reader) Close() error {
	if r.closed {
		return nil
	}
	r.closed = true
	var errs []error
	if r.shared != nil {
		errs = append(errs, r.shared.Close())
	}
	if r.closer != nil {
		errs = append(errs, r.closer.Close())
	}
	return errors.Join(errs...)
}
