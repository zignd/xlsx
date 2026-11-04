package xlsx

import (
	"fmt"
	"strings"
)

const contentTypeSharedStrings = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
const relationshipSharedStrings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"

type sharedStringsWriter struct {
	indices map[string]int
	values  []string
	count   int
}

func newSharedStringsWriter() *sharedStringsWriter {
	return &sharedStringsWriter{indices: make(map[string]int)}
}

func (s *sharedStringsWriter) index(value string) int {
	s.count++
	if index, ok := s.indices[value]; ok {
		return index
	}
	index := len(s.values)
	s.indices[value] = index
	s.values = append(s.values, value)
	return index
}

func (s *sharedStringsWriter) uniqueCount() int {
	return len(s.values)
}

func (s *sharedStringsWriter) xml() string {
	var b strings.Builder
	b.WriteString(xmlDeclaration)
	fmt.Fprintf(
		&b,
		`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">`,
		s.count, len(s.values),
	)
	for _, value := range s.values {
		b.WriteString(`<si><t xml:space="preserve">`)
		b.WriteString(xmlEscape(value))
		b.WriteString(`</t></si>`)
	}
	b.WriteString(`</sst>`)
	return b.String()
}
