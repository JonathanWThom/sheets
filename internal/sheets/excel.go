package sheets

import (
	"strings"

	"github.com/xuri/excelize/v2"
)

func isExcelPath(path string) bool {
	return strings.HasSuffix(strings.ToLower(path), ".xlsx")
}

func loadExcelRecords(path string) ([][]string, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	sheet := f.GetSheetName(0)
	return f.GetRows(sheet)
}

func writeExcelRecords(path string, records [][]string) error {
	f := excelize.NewFile()
	defer f.Close()

	sheet := f.GetSheetName(0)
	for row, record := range records {
		for col, value := range record {
			cell, err := excelize.CoordinatesToCellName(col+1, row+1)
			if err != nil {
				return err
			}
			if err := f.SetCellValue(sheet, cell, value); err != nil {
				return err
			}
		}
	}

	return f.SaveAs(path)
}
