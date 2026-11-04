package xlsx

import (
	"archive/zip"
	"compress/flate"
	"context"
	"errors"
	"fmt"
	"io"
	"io/fs"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
)

// Workbook represents a spreadsheet workbook.
type Workbook struct {
	// FilePath is the destination of the XLSX file written by Commit.
	FilePath        string
	worksheets      []*Worksheet
	tempDirsCreated bool
	tempRootDir     string
	tempDirs        *workbookTempDirs
	relationships   []*relationship
	committed       bool
	cleaned         bool
	options         Options
	output          io.Writer
	sharedStrings   *sharedStringsWriter
	styles          []Style
	streaming       bool
	stream          *streamState
}

type streamState struct {
	zipWriter *zip.Writer
	openSheet *Worksheet
	closed    bool
}

// WorksheetOptions has options used when creating a new worksheet.
type WorksheetOptions struct {
	// Name is the worksheet name shown in Excel.
	Name string
}

type workbookTempDirs struct {
	rels            string
	xl              string
	xlRels          string
	xlWorksheets    string
	xlWorksheetRels string
}

type relationship struct {
	relType string
	target  string
}

// AddWorksheet adds a new worksheet to the workbook.
//
// Invalid or duplicate worksheet names are reported when the worksheet (or the
// workbook) is committed, because this method intentionally keeps the original
// error-free API.
func (wb *Workbook) AddWorksheet(opts *WorksheetOptions) *Worksheet {
	if opts == nil {
		opts = &WorksheetOptions{}
	}

	if err := wb.checkContext(); err != nil {
		return &Worksheet{workbook: wb, err: err}
	}
	if len(wb.worksheets) >= wb.options.MaxSheets {
		id := len(wb.worksheets) + 1
		ws := &Worksheet{
			workbook: wb,
			id:       id,
			name:     opts.Name,
			fileName: fmt.Sprintf("sheet%d.xml", id),
			err:      fmt.Errorf("%w: maximum is %d", ErrTooManySheets, wb.options.MaxSheets),
		}
		wb.worksheets = append(wb.worksheets, ws)
		return ws
	}

	id := len(wb.worksheets) + 1
	ws := &Worksheet{
		workbook: wb,
		id:       id,
		name:     opts.Name,
		fileName: fmt.Sprintf("sheet%d.xml", id),
	}

	if wb.committed {
		ws.err = fmt.Errorf("%w: can't add worksheets", ErrWorkbookCommitted)
	} else if err := validateSheetName(opts.Name); err != nil {
		ws.err = err
	} else {
		for _, existing := range wb.worksheets {
			if strings.EqualFold(existing.name, opts.Name) {
				ws.err = fmt.Errorf("%w: %q", ErrDuplicateSheetName, opts.Name)
				break
			}
		}
	}

	wb.worksheets = append(wb.worksheets, ws)
	wb.relationships = append(wb.relationships, &relationship{
		relType: contentTypeWorksheetR,
		target:  filepath.ToSlash(filepath.Join("worksheets", ws.fileName)),
	})

	return ws
}

// Worksheets returns a copy of the workbook's worksheet slice.
func (wb *Workbook) Worksheets() []*Worksheet {
	worksheets := make([]*Worksheet, len(wb.worksheets))
	copy(worksheets, wb.worksheets)
	return worksheets
}

// HasPendingWorksheets indicates whether the workbook has worksheets that have
// not been finalized yet.
func (wb *Workbook) HasPendingWorksheets() bool {
	for _, ws := range wb.worksheets {
		if !ws.committed || len(ws.pendingRows) > 0 {
			return true
		}
	}
	return false
}

// TempDir returns the current private temporary directory, or an empty string
// before temporary files have been created.
func (wb *Workbook) TempDir() string {
	return wb.tempRootDir
}

// Context returns the writer context.
func (wb *Workbook) Context() context.Context {
	return wb.options.Context
}

func (wb *Workbook) checkContext() error {
	if wb.options.Context == nil {
		return nil
	}
	return wb.options.Context.Err()
}

// Close releases temporary resources. It is an idempotent alias for Cleanup and
// is intended for defer use.
func (wb *Workbook) Close() error {
	return wb.Cleanup()
}

func (wb *Workbook) ensureTempDirs() error {
	if wb.tempDirsCreated {
		return nil
	}
	return wb.createTempDirs()
}

func (wb *Workbook) createTempDirs() error {
	if wb.tempDirsCreated {
		return errors.New("temporary directories already created")
	}
	if wb.tempDirs == nil {
		wb.tempDirs = &workbookTempDirs{}
	}

	tempRootDir, err := os.MkdirTemp(wb.options.TempDir, "xlsx-*")
	if err != nil {
		return fmt.Errorf("failed to create a temporary directory for the workbook files: %w", err)
	}
	if err := os.Chmod(tempRootDir, wb.options.TempDirMode); err != nil {
		_ = os.RemoveAll(tempRootDir)
		return fmt.Errorf("failed to secure temporary directory %s: %w", tempRootDir, err)
	}

	dirs := []struct {
		path   string
		target *string
	}{
		{filepath.Join(tempRootDir, "_rels"), &wb.tempDirs.rels},
		{filepath.Join(tempRootDir, "xl"), &wb.tempDirs.xl},
		{filepath.Join(tempRootDir, "xl", "_rels"), &wb.tempDirs.xlRels},
		{filepath.Join(tempRootDir, "xl", "worksheets"), &wb.tempDirs.xlWorksheets},
		{filepath.Join(tempRootDir, "xl", "worksheets", "_rels"), &wb.tempDirs.xlWorksheetRels},
	}

	for _, dir := range dirs {
		if err := os.MkdirAll(dir.path, wb.options.TempDirMode); err != nil {
			_ = os.RemoveAll(tempRootDir)
			return fmt.Errorf("failed to create temporary directory %q: %w", dir.path, err)
		}
		*dir.target = dir.path
	}

	wb.tempRootDir = tempRootDir
	wb.tempDirsCreated = true
	return nil
}

func (wb *Workbook) contentTypes() string {
	var b strings.Builder
	b.WriteString(contentTypesStart)
	b.WriteString(contentTypeOverride("/xl/workbook.xml", contentTypeWorkbook))
	for _, ws := range wb.worksheets {
		b.WriteString(contentTypeOverride("/xl/worksheets/"+ws.fileName, contentTypeWorksheet))
	}
	b.WriteString(contentTypeOverride("/xl/styles.xml", contentTypeStyles))
	if wb.options.SharedStrings && wb.sharedStrings.uniqueCount() > 0 {
		b.WriteString(contentTypeOverride("/xl/sharedStrings.xml", contentTypeSharedStrings))
	}
	b.WriteString(contentTypesEnd)
	return b.String()
}

func (wb *Workbook) workbookXML() string {
	var b strings.Builder
	b.WriteString(workbookStart)
	for _, ws := range wb.worksheets {
		b.WriteString(sheetXML(ws.name, ws.id))
	}
	b.WriteString(workbookEnd)
	return b.String()
}

func (wb *Workbook) workbookRelationships() string {
	var b strings.Builder
	b.WriteString(workbookRelsStart)
	for i, rel := range wb.relationships {
		b.WriteString(relationshipXML(i+1, rel.relType, rel.target))
	}
	nextID := len(wb.relationships) + 1
	b.WriteString(relationshipXML(nextID, contentTypeStylesR, "styles.xml"))
	nextID++
	if wb.options.SharedStrings && wb.sharedStrings.uniqueCount() > 0 {
		b.WriteString(relationshipXML(nextID, relationshipSharedStrings, "sharedStrings.xml"))
	}
	b.WriteString(workbookRelsEnd)
	return b.String()
}

func (wb *Workbook) createMetadataFiles() error {
	files := []struct {
		path    string
		content string
	}{
		{filepath.Join(wb.tempRootDir, "[Content_Types].xml"), wb.contentTypes()},
		{filepath.Join(wb.tempDirs.rels, ".rels"), rootRelationships},
		{filepath.Join(wb.tempDirs.xl, "workbook.xml"), wb.workbookXML()},
		{filepath.Join(wb.tempDirs.xlRels, "workbook.xml.rels"), wb.workbookRelationships()},
		{filepath.Join(wb.tempDirs.xl, "styles.xml"), wb.stylesXML()},
	}
	if wb.options.SharedStrings && wb.sharedStrings.uniqueCount() > 0 {
		files = append(files, struct {
			path    string
			content string
		}{filepath.Join(wb.tempDirs.xl, "sharedStrings.xml"), wb.sharedStrings.xml()})
	}

	for _, file := range files {
		if err := os.WriteFile(file.path, []byte(file.content), wb.options.FileMode); err != nil {
			return fmt.Errorf("failed to write %s: %w", file.path, err)
		}
	}
	return nil
}

// Commit commits the workbook, persisting the data to the file path supplied to
// NewWorkbook. Worksheets with no pending rows are finalized automatically.
func (wb *Workbook) Commit() error {
	if err := wb.checkContext(); err != nil {
		return err
	}
	if wb.committed {
		return fmt.Errorf("%w: can't commit it twice", ErrWorkbookCommitted)
	}
	if wb.cleaned {
		return errors.New("workbook was cleaned up before it was committed")
	}
	if len(wb.worksheets) == 0 {
		return ErrNoWorksheets
	}
	if wb.output == nil && wb.FilePath == "" {
		return errors.New("the workbook file path cannot be empty")
	}

	for _, ws := range wb.worksheets {
		if ws.err != nil {
			return fmt.Errorf("worksheet %q is invalid: %w", ws.name, ws.err)
		}
		if len(ws.pendingRows) > 0 {
			return fmt.Errorf("%w: worksheet %q has %d pending row(s), call CommitRows first",
				ErrPendingRows, ws.name, len(ws.pendingRows))
		}
		if !ws.committed {
			if err := ws.Commit(); err != nil {
				return fmt.Errorf("failed to commit worksheet %q: %w", ws.name, err)
			}
		}
	}

	if err := wb.checkContext(); err != nil {
		return err
	}
	if wb.streaming {
		return wb.finishStreaming()
	}
	if err := wb.ensureTempDirs(); err != nil {
		return fmt.Errorf("failed to create the temporary directories: %w", err)
	}
	if err := wb.createMetadataFiles(); err != nil {
		return err
	}
	if wb.output != nil {
		if err := wb.writeTo(wb.output); err != nil {
			return err
		}
	} else if err := wb.writeFile(); err != nil {
		return err
	}

	wb.committed = true
	_ = wb.Cleanup()
	return nil
}

func (wb *Workbook) finishStreaming() error {
	if wb.stream == nil {
		return errors.New("streaming state is not initialized")
	}
	if wb.stream.closed {
		return errors.New("streaming writer is already closed")
	}
	if wb.stream.openSheet != nil {
		return fmt.Errorf("%w: worksheet %q is still open", ErrStreamingOrder, wb.stream.openSheet.name)
	}

	parts := []struct {
		name    string
		content string
	}{
		{"[Content_Types].xml", wb.contentTypes()},
		{"_rels/.rels", rootRelationships},
		{"xl/workbook.xml", wb.workbookXML()},
		{"xl/_rels/workbook.xml.rels", wb.workbookRelationships()},
		{"xl/styles.xml", wb.stylesXML()},
	}
	if wb.options.SharedStrings && wb.sharedStrings.uniqueCount() > 0 {
		parts = append(parts, struct {
			name    string
			content string
		}{"xl/sharedStrings.xml", wb.sharedStrings.xml()})
	}

	for _, part := range parts {
		if err := wb.writeStreamPart(part.name, part.content); err != nil {
			return err
		}
	}
	if err := wb.stream.zipWriter.Close(); err != nil {
		return fmt.Errorf("failed to close the streaming zip writer: %w", err)
	}
	wb.stream.closed = true
	wb.committed = true
	wb.cleaned = true
	return nil
}

func (wb *Workbook) writeStreamPart(name, content string) error {
	if wb.stream == nil || wb.stream.closed {
		return errors.New("streaming writer is closed")
	}
	if wb.stream.openSheet != nil {
		return fmt.Errorf("%w: worksheet %q is still open", ErrStreamingOrder, wb.stream.openSheet.name)
	}
	header := &zip.FileHeader{Name: name, Method: zip.Deflate}
	if wb.options.Compression == CompressionStore {
		header.Method = zip.Store
	}
	header.SetMode(wb.options.FileMode)
	writer, err := wb.stream.zipWriter.CreateHeader(header)
	if err != nil {
		return fmt.Errorf("failed to add %s to the streaming zip: %w", name, err)
	}
	if _, err := io.WriteString(writer, content); err != nil {
		return fmt.Errorf("failed to write %s to the streaming zip: %w", name, err)
	}
	return nil
}

// Flush pushes buffered data to the underlying writer. For a streaming writer
// it also calls Flush on the underlying writer when it implements
// interface{ Flush() }, which makes it useful with http.Flusher.
func (wb *Workbook) Flush() error {
	if err := wb.checkContext(); err != nil {
		return err
	}
	if wb.streaming {
		if wb.stream == nil || wb.stream.closed {
			return nil
		}
		if wb.stream.openSheet != nil && wb.stream.openSheet.writer != nil {
			if err := wb.stream.openSheet.writer.Flush(); err != nil {
				return fmt.Errorf("failed to flush the worksheet: %w", err)
			}
		}
		if err := wb.stream.zipWriter.Flush(); err != nil {
			return fmt.Errorf("failed to flush the streaming zip: %w", err)
		}
		if wb.stream.openSheet != nil {
			wb.stream.openSheet.unflushedRows = 0
		}
		if flusher, ok := wb.output.(interface{ Flush() }); ok {
			flusher.Flush()
		}
		return nil
	}
	for _, ws := range wb.worksheets {
		if ws.writer != nil {
			if err := ws.writer.Flush(); err != nil {
				return fmt.Errorf("failed to flush worksheet %q: %w", ws.name, err)
			}
		}
	}
	return nil
}

func (wb *Workbook) writeTo(w io.Writer) error {
	zw := zip.NewWriter(w)
	wb.registerCompressor(zw)
	if err := wb.addFilesToZip(zw); err != nil {
		_ = zw.Close()
		return err
	}
	if err := zw.Close(); err != nil {
		return fmt.Errorf("failed to close the zip writer: %w", err)
	}
	return nil
}

func (wb *Workbook) registerCompressor(zw *zip.Writer) {
	switch wb.options.Compression {
	case CompressionBestSpeed:
		zw.RegisterCompressor(zip.Deflate, func(out io.Writer) (io.WriteCloser, error) {
			return flate.NewWriter(out, flate.BestSpeed)
		})
	case CompressionBestCompression:
		zw.RegisterCompressor(zip.Deflate, func(out io.Writer) (io.WriteCloser, error) {
			return flate.NewWriter(out, flate.BestCompression)
		})
	}
}

func (wb *Workbook) writeFile() error {
	dir := filepath.Dir(wb.FilePath)
	if err := os.MkdirAll(dir, 0o750); err != nil {
		return fmt.Errorf("failed to create the output directory %s: %w", dir, err)
	}
	if info, statErr := os.Stat(wb.FilePath); statErr == nil && info.IsDir() {
		return fmt.Errorf("the output path %s is a directory", wb.FilePath)
	}

	tmp, err := os.CreateTemp(dir, ".xlsx-*.tmp")
	if err != nil {
		return fmt.Errorf("failed to create a temporary output file: %w", err)
	}
	tmpPath := tmp.Name()
	removeTmp := true
	defer func() {
		if removeTmp {
			_ = os.Remove(tmpPath)
		}
	}()

	if err := wb.writeTo(tmp); err != nil {
		_ = tmp.Close()
		return err
	}
	if err := tmp.Sync(); err != nil {
		_ = tmp.Close()
		return fmt.Errorf("failed to sync the output file: %w", err)
	}
	if err := tmp.Close(); err != nil {
		return fmt.Errorf("failed to close the output file: %w", err)
	}
	if err := os.Chmod(tmpPath, wb.options.FileMode); err != nil {
		return fmt.Errorf("failed to set permissions on the output file: %w", err)
	}

	if err := os.Rename(tmpPath, wb.FilePath); err != nil {
		if runtime.GOOS != "windows" {
			return fmt.Errorf("failed to move the output file to %s: %w", wb.FilePath, err)
		}
		// Windows does not replace an existing destination atomically.
		if removeErr := os.Remove(wb.FilePath); removeErr == nil || os.IsNotExist(removeErr) {
			if err = os.Rename(tmpPath, wb.FilePath); err == nil {
				removeTmp = false
				return nil
			}
		}
		return fmt.Errorf("failed to move the output file to %s: %w", wb.FilePath, err)
	}

	removeTmp = false
	return nil
}

func (wb *Workbook) addFilesToZip(zw *zip.Writer) error {
	var files []string
	err := filepath.WalkDir(wb.tempRootDir, func(path string, entry fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if entry.IsDir() {
			return nil
		}
		files = append(files, path)
		return nil
	})
	if err != nil {
		return fmt.Errorf("failed to enumerate the temporary files: %w", err)
	}
	sort.Strings(files)

	for _, path := range files {
		relativePath, err := filepath.Rel(wb.tempRootDir, path)
		if err != nil {
			return fmt.Errorf("failed to determine the relative path of %s: %w", path, err)
		}

		info, err := os.Stat(path)
		if err != nil {
			return fmt.Errorf("failed to stat %s: %w", path, err)
		}
		header, err := zip.FileInfoHeader(info)
		if err != nil {
			return fmt.Errorf("failed to create a zip header for %s: %w", path, err)
		}
		header.Name = filepath.ToSlash(relativePath)
		if wb.options.Compression == CompressionStore {
			header.Method = zip.Store
		} else {
			header.Method = zip.Deflate
		}

		writer, err := zw.CreateHeader(header)
		if err != nil {
			return fmt.Errorf("failed to add %s to the zip file: %w", relativePath, err)
		}
		f, err := os.Open(path) // #nosec G304 -- path is inside the managed temporary directory.
		if err != nil {
			return fmt.Errorf("failed to open %s: %w", path, err)
		}
		if _, err := io.Copy(writer, f); err != nil {
			_ = f.Close()
			return fmt.Errorf("failed to copy %s into the zip file: %w", path, err)
		}
		if err := f.Close(); err != nil {
			return fmt.Errorf("failed to close %s: %w", path, err)
		}
	}

	return nil
}

// Cleanup removes the temporary files used to build the workbook. It is safe
// to call multiple times and is called automatically after a successful
// Commit.
func (wb *Workbook) Cleanup() error {
	for _, ws := range wb.worksheets {
		ws.closeFile()
		if wb.tempDirsCreated && !ws.committed {
			ws.err = errors.New("worksheet was cleaned up before it was committed")
		}
	}
	if wb.streaming {
		if wb.stream != nil && !wb.stream.closed {
			_ = wb.stream.zipWriter.Close()
			wb.stream.closed = true
		}
		wb.cleaned = true
		return nil
	}
	if !wb.tempDirsCreated {
		return nil
	}

	if wb.options.KeepTempFiles {
		wb.tempDirsCreated = false
		wb.cleaned = true
		return nil
	}

	tempRootDir := wb.tempRootDir
	err := os.RemoveAll(tempRootDir)
	wb.tempDirsCreated = false
	wb.tempRootDir = ""
	wb.tempDirs = &workbookTempDirs{}
	wb.cleaned = true
	if err != nil {
		return fmt.Errorf("failed to remove the temporary directory %s: %w", tempRootDir, err)
	}
	return nil
}
