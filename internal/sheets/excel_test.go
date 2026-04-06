package sheets

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

func createTestExcelFile(t *testing.T, records [][]string) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	sheet := f.GetSheetName(0)
	for row, record := range records {
		for col, value := range record {
			cell, err := excelize.CoordinatesToCellName(col+1, row+1)
			if err != nil {
				t.Fatal(err)
			}
			if err := f.SetCellValue(sheet, cell, value); err != nil {
				t.Fatal(err)
			}
		}
	}

	path := filepath.Join(t.TempDir(), "test.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	return path
}

func TestIsExcelPath(t *testing.T) {
	tests := []struct {
		path string
		want bool
	}{
		{"file.xlsx", true},
		{"file.XLSX", true},
		{"file.csv", false},
		{"file.tsv", false},
		{"dir/file.xlsx", true},
		{"file.xlsx.bak", false},
	}

	for _, tt := range tests {
		if got := isExcelPath(tt.path); got != tt.want {
			t.Errorf("isExcelPath(%q) = %v, want %v", tt.path, got, tt.want)
		}
	}
}

func TestLoadExcelRecords(t *testing.T) {
	input := [][]string{
		{"Name", "Age", "City"},
		{"Alice", "30", "NYC"},
		{"Bob", "25", "LA"},
	}
	path := createTestExcelFile(t, input)

	records, err := loadExcelRecords(path)
	if err != nil {
		t.Fatal(err)
	}

	if len(records) != len(input) {
		t.Fatalf("got %d rows, want %d", len(records), len(input))
	}
	for i, row := range records {
		if len(row) != len(input[i]) {
			t.Fatalf("row %d: got %d cols, want %d", i, len(row), len(input[i]))
		}
		for j, cell := range row {
			if cell != input[i][j] {
				t.Errorf("cell [%d][%d] = %q, want %q", i, j, cell, input[i][j])
			}
		}
	}
}

func TestWriteAndReadExcelRoundTrip(t *testing.T) {
	records := [][]string{
		{"Header1", "Header2"},
		{"val1", "val2"},
		{"val3", "val4"},
	}

	path := filepath.Join(t.TempDir(), "output.xlsx")
	if err := writeExcelRecords(path, records); err != nil {
		t.Fatal(err)
	}

	got, err := loadExcelRecords(path)
	if err != nil {
		t.Fatal(err)
	}

	if len(got) != len(records) {
		t.Fatalf("got %d rows, want %d", len(got), len(records))
	}
	for i, row := range got {
		for j, cell := range row {
			if cell != records[i][j] {
				t.Errorf("cell [%d][%d] = %q, want %q", i, j, cell, records[i][j])
			}
		}
	}
}

func TestModelLoadExcelFile(t *testing.T) {
	input := [][]string{
		{"A", "B"},
		{"1", "2"},
	}
	path := createTestExcelFile(t, input)

	m := newModel()
	if err := m.loadFile(path); err != nil {
		t.Fatal(err)
	}

	if m.cellValue(0, 0) != "A" {
		t.Errorf("cell (0,0) = %q, want %q", m.cellValue(0, 0), "A")
	}
	if m.cellValue(1, 1) != "2" {
		t.Errorf("cell (1,1) = %q, want %q", m.cellValue(1, 1), "2")
	}
	if m.currentFilePath != path {
		t.Errorf("currentFilePath = %q, want %q", m.currentFilePath, path)
	}
}

func TestModelWriteExcelFile(t *testing.T) {
	m := newModel()
	m.setCellValue(0, 0, "X")
	m.setCellValue(0, 1, "Y")
	m.setCellValue(1, 0, "1")
	m.setCellValue(1, 1, "2")

	path := filepath.Join(t.TempDir(), "write_test.xlsx")
	if err := m.writeFile(path); err != nil {
		t.Fatal(err)
	}

	if _, err := os.Stat(path); err != nil {
		t.Fatalf("output file not created: %v", err)
	}

	m2 := newModel()
	if err := m2.loadFile(path); err != nil {
		t.Fatal(err)
	}

	if m2.cellValue(0, 0) != "X" {
		t.Errorf("cell (0,0) = %q, want %q", m2.cellValue(0, 0), "X")
	}
	if m2.cellValue(1, 1) != "2" {
		t.Errorf("cell (1,1) = %q, want %q", m2.cellValue(1, 1), "2")
	}
}

func TestLoadExcelFileNotFound(t *testing.T) {
	_, err := loadExcelRecords("/nonexistent/file.xlsx")
	if err == nil {
		t.Fatal("expected error for nonexistent file")
	}
}
