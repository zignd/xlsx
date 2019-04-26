package xlsx

import (
	"fmt"
	"testing"
	"time"
)

func Test_DateCell_ShouldProperlyCreateCell(t *testing.T) {
	now := time.Now()
	expectedCellStr := fmt.Sprintf(`<c r="A1" s="1" t="n"><v>%f</v></c>`, timeToExcelTime(timeToUTCTime(now)))

	cellStr := dateCellFormat("A1", now)

	if cellStr != expectedCellStr {
		t.Errorf("cell string differs from the expected, found: %s, expected: %s", cellStr, expectedCellStr)
	}
}
