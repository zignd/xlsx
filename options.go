package xlsx

import (
	"context"
	"fmt"
	"io"
	"os"
)

const (
	defaultBufferSize   = 64 * 1024
	defaultMaxCellRunes = 32_767
	defaultMaxSheets    = 1_024
)

// Compression controls how the final ZIP package is compressed.
type Compression uint8

const (
	// CompressionDefault uses the standard DEFLATE compression level.
	CompressionDefault Compression = iota
	// CompressionStore writes ZIP entries without compression.
	CompressionStore
	// CompressionBestSpeed favours speed over size.
	CompressionBestSpeed
	// CompressionBestCompression favours size over speed.
	CompressionBestCompression
)

// Options configures a workbook writer.
type Options struct {
	// TempDir is the parent directory for temporary worksheet files. The
	// library creates a private subdirectory inside it.
	TempDir string

	// BufferSize is the size of the per-worksheet write buffer. Zero uses the
	// package default.
	BufferSize int

	// TempDirMode is the permission mode for temporary directories.
	TempDirMode os.FileMode

	// FileMode is the permission mode for the final .xlsx file.
	FileMode os.FileMode

	// Compression selects the ZIP compression strategy.
	Compression Compression

	// KeepTempFiles leaves the temporary workbook directory in place after
	// Commit. It is useful for debugging.
	KeepTempFiles bool

	// Context cancels long-running writes. Nil means context.Background().
	Context context.Context

	// SharedStrings writes repeated strings through xl/sharedStrings.xml.
	SharedStrings bool

	// MaxCellLength limits text values. Zero uses Excel's 32,767-character
	// limit.
	MaxCellLength int

	// MaxRows limits rows per worksheet. Zero uses Excel's limit.
	MaxRows int

	// MaxColumns limits columns per worksheet. Zero uses Excel's limit.
	MaxColumns int

	// MaxSheets limits the number of worksheets. Zero uses 1,024.
	MaxSheets int
}

// DefaultOptions returns the default writer configuration.
func DefaultOptions() Options {
	return Options{
		TempDir:       os.TempDir(),
		BufferSize:    defaultBufferSize,
		TempDirMode:   0o700,
		FileMode:      0o644,
		Compression:   CompressionDefault,
		Context:       context.Background(),
		MaxCellLength: defaultMaxCellRunes,
		MaxRows:       maxWorksheetRows,
		MaxColumns:    maxWorksheetColumns,
		MaxSheets:     defaultMaxSheets,
	}
}

func (o Options) normalized() (Options, error) {
	defaults := DefaultOptions()

	if o.TempDir == "" {
		o.TempDir = defaults.TempDir
	}
	if o.BufferSize == 0 {
		o.BufferSize = defaults.BufferSize
	}
	if o.TempDirMode == 0 {
		o.TempDirMode = defaults.TempDirMode
	}
	if o.FileMode == 0 {
		o.FileMode = defaults.FileMode
	}
	if o.Context == nil {
		o.Context = defaults.Context
	}
	if o.MaxCellLength == 0 {
		o.MaxCellLength = defaults.MaxCellLength
	}
	if o.MaxRows == 0 {
		o.MaxRows = defaults.MaxRows
	}
	if o.MaxColumns == 0 {
		o.MaxColumns = defaults.MaxColumns
	}
	if o.MaxSheets == 0 {
		o.MaxSheets = defaults.MaxSheets
	}

	switch {
	case o.BufferSize < 0:
		return Options{}, fmt.Errorf("buffer size must be positive")
	case o.MaxCellLength < 0:
		return Options{}, fmt.Errorf("max cell length must be positive")
	case o.MaxRows < 0 || o.MaxRows > maxWorksheetRows:
		return Options{}, fmt.Errorf("max rows must be between 1 and %d", maxWorksheetRows)
	case o.MaxColumns < 0 || o.MaxColumns > maxWorksheetColumns:
		return Options{}, fmt.Errorf("max columns must be between 1 and %d", maxWorksheetColumns)
	case o.MaxSheets < 0:
		return Options{}, fmt.Errorf("max sheets must be positive")
	case o.Compression > CompressionBestCompression:
		return Options{}, fmt.Errorf("unknown compression setting %d", o.Compression)
	}

	return o, nil
}

func validateTempDir(dir string) error {
	info, err := os.Stat(dir)
	if err != nil {
		return fmt.Errorf("temporary directory %q is not accessible: %w", dir, err)
	}
	if !info.IsDir() {
		return fmt.Errorf("temporary directory %q is not a directory", dir)
	}
	probe, err := os.CreateTemp(dir, ".xlsx-write-test-*")
	if err != nil {
		return fmt.Errorf("temporary directory %q is not writable: %w", dir, err)
	}
	name := probe.Name()
	if err := probe.Close(); err != nil {
		return fmt.Errorf("failed to close temporary probe %s: %w", name, err)
	}
	if err := os.Remove(name); err != nil {
		return fmt.Errorf("failed to remove temporary probe %s: %w", name, err)
	}
	return nil
}

// NewWorkbookWithOptions creates a workbook with explicit writer options.
func NewWorkbookWithOptions(filePath string, opts Options) (*Workbook, error) {
	normalized, err := opts.normalized()
	if err != nil {
		return nil, err
	}
	if opts.TempDir != "" {
		if err := validateTempDir(normalized.TempDir); err != nil {
			return nil, err
		}
	}
	return &Workbook{
		FilePath:      filePath,
		tempDirs:      &workbookTempDirs{},
		options:       normalized,
		sharedStrings: newSharedStringsWriter(),
	}, nil
}

// NewWriter creates a workbook that writes its final package to w instead of a
// file path. The caller owns w; the workbook does not close it.
func NewWriter(w io.Writer, opts Options) (*Workbook, error) {
	if w == nil {
		return nil, fmt.Errorf("output writer cannot be nil")
	}
	normalized, err := opts.normalized()
	if err != nil {
		return nil, err
	}
	if opts.TempDir != "" {
		if err := validateTempDir(normalized.TempDir); err != nil {
			return nil, err
		}
	}
	return &Workbook{
		tempDirs:      &workbookTempDirs{},
		options:       normalized,
		output:        w,
		sharedStrings: newSharedStringsWriter(),
	}, nil
}
