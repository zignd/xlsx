package xlsx

import (
	"context"
	"fmt"
	"io"
	"os"
)

const (
	defaultMaxUncompressed = int64(1 << 30)
	defaultMaxEntries      = 10_000
	defaultSharedStringMem = int64(4 << 20)
)

// SharedStringsMode controls how shared strings are read.
type SharedStringsMode uint8

const (
	// SharedStringsAuto keeps small tables in memory and spills large tables to
	// disk.
	SharedStringsAuto SharedStringsMode = iota
	// SharedStringsMemory keeps every shared string in memory.
	SharedStringsMemory
	// SharedStringsDisk stores shared strings in temporary files and reads them
	// on demand.
	SharedStringsDisk
)

// ReadOptions configures a streaming XLSX reader.
type ReadOptions struct {
	// Context cancels long-running reads. Nil means context.Background().
	Context context.Context

	// SharedStringsMode selects the shared-string storage strategy.
	SharedStringsMode SharedStringsMode

	// SharedStringsMemoryLimit is the auto-mode threshold in bytes.
	SharedStringsMemoryLimit int64

	// TempDir is the parent directory for disk-backed shared strings.
	TempDir string

	// MaxUncompressedSize limits the total uncompressed package size.
	MaxUncompressedSize int64

	// MaxEntries limits the number of ZIP entries.
	MaxEntries int

	// MaxRows limits rows per worksheet.
	MaxRows int

	// MaxColumns limits columns per worksheet.
	MaxColumns int

	// MaxCellLength limits text values.
	MaxCellLength int
}

// DefaultReadOptions returns the default reader configuration.
func DefaultReadOptions() ReadOptions {
	return ReadOptions{
		Context:                  context.Background(),
		SharedStringsMode:        SharedStringsAuto,
		SharedStringsMemoryLimit: defaultSharedStringMem,
		TempDir:                  os.TempDir(),
		MaxUncompressedSize:      defaultMaxUncompressed,
		MaxEntries:               defaultMaxEntries,
		MaxRows:                  maxWorksheetRows,
		MaxColumns:               maxWorksheetColumns,
		MaxCellLength:            defaultMaxCellRunes,
	}
}

func (o ReadOptions) normalized() (ReadOptions, error) {
	defaults := DefaultReadOptions()
	if o.Context == nil {
		o.Context = defaults.Context
	}
	if o.SharedStringsMemoryLimit == 0 {
		o.SharedStringsMemoryLimit = defaults.SharedStringsMemoryLimit
	}
	if o.TempDir == "" {
		o.TempDir = defaults.TempDir
	}
	if o.MaxUncompressedSize == 0 {
		o.MaxUncompressedSize = defaults.MaxUncompressedSize
	}
	if o.MaxEntries == 0 {
		o.MaxEntries = defaults.MaxEntries
	}
	if o.MaxRows == 0 {
		o.MaxRows = defaults.MaxRows
	}
	if o.MaxColumns == 0 {
		o.MaxColumns = defaults.MaxColumns
	}
	if o.MaxCellLength == 0 {
		o.MaxCellLength = defaults.MaxCellLength
	}
	if o.SharedStringsMode > SharedStringsDisk {
		return ReadOptions{}, fmt.Errorf("unknown shared strings mode %d", o.SharedStringsMode)
	}
	if o.SharedStringsMemoryLimit < 0 || o.MaxUncompressedSize < 0 || o.MaxEntries < 0 || o.MaxRows < 0 || o.MaxColumns < 0 || o.MaxCellLength < 0 {
		return ReadOptions{}, fmt.Errorf("read limits cannot be negative")
	}
	return o, nil
}

// ReadOption customizes ReadOptions.
type ReadOption func(*ReadOptions)

// WithReadContext sets the reader context.
func WithReadContext(ctx context.Context) ReadOption {
	return func(o *ReadOptions) { o.Context = ctx }
}

// WithSharedStringsMode sets the shared-string storage mode.
func WithSharedStringsMode(mode SharedStringsMode) ReadOption {
	return func(o *ReadOptions) { o.SharedStringsMode = mode }
}

// WithReadTempDir sets the parent directory for disk-backed shared strings.
func WithReadTempDir(dir string) ReadOption {
	return func(o *ReadOptions) { o.TempDir = dir }
}

// WithReadLimits overrides the reader safety limits.
func WithReadLimits(maxUncompressedSize int64, maxEntries int) ReadOption {
	return func(o *ReadOptions) {
		o.MaxUncompressedSize = maxUncompressedSize
		o.MaxEntries = maxEntries
	}
}

// WithReadMaxRows limits rows per worksheet.
func WithReadMaxRows(rows int) ReadOption {
	return func(o *ReadOptions) { o.MaxRows = rows }
}

// WithReadMaxColumns limits columns per worksheet.
func WithReadMaxColumns(columns int) ReadOption {
	return func(o *ReadOptions) { o.MaxColumns = columns }
}

// WithReadMaxCellLength limits text values.
func WithReadMaxCellLength(length int) ReadOption {
	return func(o *ReadOptions) { o.MaxCellLength = length }
}

func applyReadOptions(opts []ReadOption) (ReadOptions, error) {
	options := DefaultReadOptions()
	for _, opt := range opts {
		if opt != nil {
			opt(&options)
		}
	}
	return options.normalized()
}

// Open opens an XLSX file for streaming reads.
func Open(path string, opts ...ReadOption) (*Reader, error) {
	options, err := applyReadOptions(opts)
	if err != nil {
		return nil, err
	}
	f, err := os.Open(path) // #nosec G304 -- the caller chooses the path to read.
	if err != nil {
		return nil, fmt.Errorf("failed to open %s: %w", path, err)
	}
	info, err := f.Stat()
	if err != nil {
		_ = f.Close()
		return nil, fmt.Errorf("failed to stat %s: %w", path, err)
	}
	if err := sniffZipFormat(f, info.Size()); err != nil {
		_ = f.Close()
		return nil, err
	}
	reader, err := newReader(f, info.Size(), f, options)
	if err != nil {
		_ = f.Close()
		return nil, err
	}
	return reader, nil
}

// OpenReaderAt opens an XLSX package from an io.ReaderAt.
func OpenReaderAt(r io.ReaderAt, size int64, opts ...ReadOption) (*Reader, error) {
	options, err := applyReadOptions(opts)
	if err != nil {
		return nil, err
	}
	if r == nil {
		return nil, fmt.Errorf("reader cannot be nil")
	}
	if size <= 0 {
		return nil, fmt.Errorf("%w: empty package", ErrCorruptFile)
	}
	if err := sniffZipFormat(r, size); err != nil {
		return nil, err
	}
	return newReader(r, size, nil, options)
}
