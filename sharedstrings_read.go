package xlsx

import (
	"archive/zip"
	"encoding/binary"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"math"
	"os"
)

const sharedStringIndexRecordSize = 12

type sharedStringItem struct {
	T string `xml:"t"`
	R []struct {
		T string `xml:"t"`
	} `xml:"r"`
}

func (s sharedStringItem) text() string {
	var result string
	result += s.T
	for _, run := range s.R {
		result += run.T
	}
	return result
}

type sharedStrings struct {
	values    []string
	dataFile  *os.File
	indexFile *os.File
	count     int
	tempDir   string
}

func openSharedStrings(entries map[string]*zip.File, partPath string, options ReadOptions) (*sharedStrings, error) {
	file, ok := entries[partPath]
	if !ok {
		return &sharedStrings{}, nil
	}

	mode := options.SharedStringsMode
	if mode == SharedStringsAuto {
		// #nosec G115 -- read options are validated as non-negative.
		if file.UncompressedSize64 > uint64(options.SharedStringsMemoryLimit) {
			mode = SharedStringsDisk
		} else {
			mode = SharedStringsMemory
		}
	}

	if mode == SharedStringsDisk {
		return readSharedStringsDisk(file, options)
	}
	return readSharedStringsMemory(file, options)
}

func readSharedStringsMemory(file *zip.File, options ReadOptions) (*sharedStrings, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, fmt.Errorf("%w: failed to open shared strings: %w", ErrCorruptFile, err)
	}
	defer rc.Close()

	result := &sharedStrings{}
	decoder := xml.NewDecoder(rc)
	for {
		if err := contextError(options.Context); err != nil {
			return nil, err
		}
		token, err := decoder.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("%w: shared strings: %w", ErrCorruptFile, err)
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "si" {
			continue
		}
		var item sharedStringItem
		if err := decoder.DecodeElement(&item, &start); err != nil {
			return nil, fmt.Errorf("%w: shared strings: %w", ErrCorruptFile, err)
		}
		result.values = append(result.values, item.text())
	}
	result.count = len(result.values)
	return result, nil
}

func readSharedStringsDisk(file *zip.File, options ReadOptions) (*sharedStrings, error) {
	tempDir, err := os.MkdirTemp(options.TempDir, "xlsx-read-*")
	if err != nil {
		return nil, fmt.Errorf("failed to create shared strings temp directory: %w", err)
	}
	dataFile, err := os.CreateTemp(tempDir, "shared-*.data")
	if err != nil {
		_ = os.RemoveAll(tempDir)
		return nil, fmt.Errorf("failed to create shared strings data file: %w", err)
	}
	indexFile, err := os.CreateTemp(tempDir, "shared-*.index")
	if err != nil {
		_ = dataFile.Close()
		_ = os.RemoveAll(tempDir)
		return nil, fmt.Errorf("failed to create shared strings index file: %w", err)
	}

	result := &sharedStrings{dataFile: dataFile, indexFile: indexFile, tempDir: tempDir}
	cleanupOnError := func() {
		_ = dataFile.Close()
		_ = indexFile.Close()
		_ = os.RemoveAll(tempDir)
	}

	rc, err := file.Open()
	if err != nil {
		cleanupOnError()
		return nil, fmt.Errorf("%w: failed to open shared strings: %w", ErrCorruptFile, err)
	}
	defer rc.Close()

	var offset int64
	decoder := xml.NewDecoder(rc)
	for {
		if err := contextError(options.Context); err != nil {
			cleanupOnError()
			return nil, err
		}
		token, err := decoder.Token()
		if errors.Is(err, io.EOF) {
			break
		}
		if err != nil {
			cleanupOnError()
			return nil, fmt.Errorf("%w: shared strings: %w", ErrCorruptFile, err)
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "si" {
			continue
		}
		var item sharedStringItem
		if err := decoder.DecodeElement(&item, &start); err != nil {
			cleanupOnError()
			return nil, fmt.Errorf("%w: shared strings: %w", ErrCorruptFile, err)
		}
		value := item.text()
		if len(value) > math.MaxUint32 {
			cleanupOnError()
			return nil, fmt.Errorf("shared string exceeds %d bytes", uint32(math.MaxUint32))
		}
		if _, err := dataFile.WriteString(value); err != nil {
			cleanupOnError()
			return nil, fmt.Errorf("failed to write shared string: %w", err)
		}

		var record [sharedStringIndexRecordSize]byte
		// #nosec G115 -- offset is a non-negative byte count.
		binary.LittleEndian.PutUint64(record[0:8], uint64(offset))
		// #nosec G115 -- length was checked against MaxUint32 above.
		binary.LittleEndian.PutUint32(record[8:12], uint32(len(value)))
		if _, err := indexFile.Write(record[:]); err != nil {
			cleanupOnError()
			return nil, fmt.Errorf("failed to write shared string index: %w", err)
		}
		offset += int64(len(value))
		result.count++
	}

	return result, nil
}

func (s *sharedStrings) lookup(index int) (string, error) {
	if index < 0 || index >= s.count {
		return "", fmt.Errorf("%w: shared string index %d is out of range", ErrCorruptFile, index)
	}
	if s.values != nil {
		return s.values[index], nil
	}
	if s.dataFile == nil || s.indexFile == nil {
		return "", fmt.Errorf("%w: shared string table is unavailable", ErrCorruptFile)
	}

	var record [sharedStringIndexRecordSize]byte
	if _, err := s.indexFile.ReadAt(record[:], int64(index)*sharedStringIndexRecordSize); err != nil {
		return "", fmt.Errorf("%w: shared string index: %w", ErrCorruptFile, err)
	}
	offset := binary.LittleEndian.Uint64(record[0:8])
	length := binary.LittleEndian.Uint32(record[8:12])
	if offset > math.MaxInt64 {
		return "", fmt.Errorf("%w: shared string offset is out of range", ErrCorruptFile)
	}
	buf := make([]byte, length)
	// #nosec G115 -- offset was checked against MaxInt64 above.
	if _, err := s.dataFile.ReadAt(buf, int64(offset)); err != nil {
		return "", fmt.Errorf("%w: shared string data: %w", ErrCorruptFile, err)
	}
	return string(buf), nil
}

func (s *sharedStrings) Close() error {
	if s == nil {
		return nil
	}
	var errs []error
	if s.dataFile != nil {
		errs = append(errs, s.dataFile.Close())
	}
	if s.indexFile != nil {
		errs = append(errs, s.indexFile.Close())
	}
	if s.tempDir != "" {
		errs = append(errs, os.RemoveAll(s.tempDir))
	}
	return errors.Join(errs...)
}
