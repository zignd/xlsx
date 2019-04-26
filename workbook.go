package xlsx

import (
	"bytes"
	"fmt"
	"io/ioutil"
	"os"
	"path"
	"strings"

	"github.com/jhoonb/archivex"
	"github.com/pkg/errors"
)

// Workbook represents a spreadsheet workbook.
type Workbook struct {
	FilePath        string
	worksheets      []*Worksheet
	tempDirsCreated bool
	tempRootDir     string
	tempDirs        *workbookTempDirs
	relationships   []*relationship
	committed       bool
}

// WorksheetOptions has options used when creating a new worksheet.
type WorksheetOptions struct {
	Name string
}

type workbookTempDirs struct {
	rels         string
	xl           string
	xlRels       string
	xlWorksheets string
}

type relationship struct {
	relType string
	target  string
}

// AddWorksheet adds a new worksheet to the workbook.
func (wb *Workbook) AddWorksheet(opts *WorksheetOptions) *Worksheet {
	id := len(wb.worksheets) + 1

	var fileName bytes.Buffer
	fileName.WriteString("sheet")
	fileName.WriteString(fmt.Sprint(id))
	fileName.WriteString(".xml")

	ws := &Worksheet{
		workbook: wb,
		id:       id,
		name:     opts.Name,
		fileName: fileName.String(),
	}
	wb.worksheets = append(wb.worksheets, ws)

	wb.relationships = append(wb.relationships, &relationship{
		relType: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
		target:  path.Join("worksheets", ws.fileName),
	})

	return ws
}

// HasPendingWorksheets indicates whether or not the workbook has worksheets yet to be committed.
func (wb *Workbook) HasPendingWorksheets() bool {
	for i := 0; i < len(wb.worksheets); i++ {
		if !wb.worksheets[i].committed {
			return true
		}
	}
	return false
}

func (wb *Workbook) createContentTypes() error {
	ctFilePath := path.Join(wb.tempRootDir, "[Content_Types].xml")
	f, err := os.Create(ctFilePath)
	if err != nil {
		return errors.Wrapf(err, "failed to create %s", ctFilePath)
	}
	defer func() {
		f.Close()
	}()

	_, err = f.WriteString(startContentTypes)
	if err != nil {
		return errors.Wrapf(err, "failed to append START_CONTENT_TYPES to %s", ctFilePath)
	}

	for i := 0; i < len(wb.worksheets); i++ {
		wsFileName := wb.worksheets[i].fileName
		_, err = f.WriteString(overrideWorksheetFormat(wsFileName))
		if err != nil {
			return errors.Wrapf(err, "failed to append override for %s to %s", wsFileName, ctFilePath)
		}
	}

	_, err = f.WriteString(endContentTypes)
	if err != nil {
		return errors.Wrapf(err, "failed to append END_CONTENT_TYPES to %s", ctFilePath)
	}

	return nil
}

func (wb *Workbook) createRootRelationships() error {
	filePath := path.Join(wb.tempDirs.rels, ".rels")
	f, err := os.Create(filePath)
	if err != nil {
		return errors.Wrapf(err, "failed to create %s", filePath)
	}
	defer f.Close()

	_, err = f.WriteString(rels)
	if err != nil {
		return errors.Wrapf(err, "failed to write to %s", filePath)
	}

	return nil
}

func (wb *Workbook) createWorkbookRelationships() error {
	filePath := path.Join(wb.tempDirs.xlRels, "workbook.xml.rels")
	f, err := os.Create(filePath)
	if err != nil {
		return errors.Wrapf(err, "failed to create %s", filePath)
	}
	defer f.Close()

	_, err = f.WriteString(startWorkbookRels)
	if err != nil {
		return errors.Wrapf(err, "failed to append START_WORKBOOK_RELS to %s", filePath)
	}

	for i := 0; i < len(wb.relationships); i++ {
		r := wb.relationships[i]
		_, err = f.WriteString(relationshipFormat(i+1, r.relType, r.target))
		if err != nil {
			return errors.Wrapf(err, "failed to append the relationship with the target %s to %s", r.target, filePath)
		}
	}

	_, err = f.WriteString(endWorkbookRels)
	if err != nil {
		return errors.Wrapf(err, "failed to append END_WORKBOOK_RELS to %s", filePath)
	}

	return nil
}

func (wb *Workbook) createMain() error {
	filePath := path.Join(wb.tempDirs.xl, "workbook.xml")
	f, err := os.Create(filePath)
	if err != nil {
		return errors.Wrapf(err, "failed to create %s", filePath)
	}
	defer f.Close()

	_, err = f.WriteString(startWorkbook)
	if err != nil {
		return errors.Wrapf(err, "failed to append START_WORKBOOK to %s", filePath)
	}

	for i := 0; i < len(wb.worksheets); i++ {
		ws := wb.worksheets[i]
		_, err = f.WriteString(sheetFormat(ws.name, ws.id))
	}

	_, err = f.WriteString(endWorkbook)
	if err != nil {
		return errors.Wrapf(err, "failed to append END_WORKBOOK to %s", filePath)
	}

	return nil
}

// Commit commits the workbook persisting the data to the specified file.
func (wb *Workbook) Commit() error {
	if wb.committed {
		return errors.New("can't commit a committed workbook")
	}

	if len(wb.worksheets) == 0 {
		return errors.New("a workbook needs at least one worksheet")
	}

	if wb.HasPendingWorksheets() {
		return errors.New("can't commit if there are still pending worksheets")
	}

	err := wb.createContentTypes()
	if err != nil {
		return errors.Wrap(err, "failed to create the content types file")
	}

	err = wb.createRootRelationships()
	if err != nil {
		return errors.Wrap(err, "failed to create the root relationships file")
	}

	err = wb.createWorkbookRelationships()
	if err != nil {
		return errors.Wrap(err, "failed to create the workbook relationships file")
	}

	err = wb.createMain()
	if err != nil {
		return errors.Wrap(err, "failed to create the main file")
	}

	zip := new(archivex.ZipFile)
	zip.Create(strings.TrimSuffix(wb.FilePath, path.Ext(wb.FilePath)))
	zip.AddAll(wb.tempRootDir, false)
	err = zip.Close()
	if err != nil {
		return errors.Wrap(err, "failed to close the zip file")
	}

	curPath := fmt.Sprintf("%s.zip", strings.TrimSuffix(wb.FilePath, path.Ext(wb.FilePath)))
	newPath := wb.FilePath
	err = os.Rename(curPath, newPath)
	if err != nil {
		return errors.Wrapf(err, "failed to rename file %s to %s", curPath, newPath)
	}

	wb.committed = true

	return nil
}

func (wb *Workbook) createTempDirs() error {
	if wb.tempDirsCreated {
		return errors.New("temporary directories already created")
	}

	tempRootDir, err := ioutil.TempDir("", "xlsx-")
	if err != nil {
		return errors.Wrap(err, "failed to create a temporary directory for the workbook files")
	}
	wb.tempRootDir = tempRootDir

	// /_rels
	wb.tempDirs.rels = path.Join(wb.tempRootDir, "_rels")
	err = os.Mkdir(wb.tempDirs.rels, os.ModeDir|os.ModePerm)
	if err != nil {
		return errors.Wrapf(err, "failed to create temporary directory \"%s\"", wb.tempDirs.rels)
	}

	// /xl
	wb.tempDirs.xl = path.Join(wb.tempRootDir, "xl")
	err = os.Mkdir(wb.tempDirs.xl, os.ModeDir|os.ModePerm)
	if err != nil {
		return errors.Wrapf(err, "failed to create temporary directory \"%s\"", wb.tempDirs.xl)
	}

	// /xl/_rels
	wb.tempDirs.xlRels = path.Join(wb.tempDirs.xl, "_rels")
	err = os.Mkdir(wb.tempDirs.xlRels, os.ModeDir|os.ModePerm)
	if err != nil {
		return errors.Wrapf(err, "failed to create temporary directory \"%s\"", wb.tempDirs.xlRels)
	}

	// /xl/worksheets
	wb.tempDirs.xlWorksheets = path.Join(wb.tempDirs.xl, "worksheets")
	err = os.Mkdir(wb.tempDirs.xlWorksheets, os.ModeDir|os.ModePerm)
	if err != nil {
		return errors.Wrapf(err, "failed to create temporary directory \"%s\"", wb.tempDirs.xlWorksheets)
	}

	wb.tempDirsCreated = true

	return nil
}
