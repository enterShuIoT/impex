package importer

import (
	"os"
	"testing"

	"github.com/xuri/excelize/v2"
)

// TestRow simulates the user's LoadForecastRow
type TestRow struct {
	ClientAccount string            `excel:"用户编号"`
	Date          string            `excel:"日期"`
	TimeData      map[string]string `excel:"extra"`
}

func createTestExcel(t *testing.T, filename string) {
	// Create a dummy Excel file
	f := excelize.NewFile()
	sheetName := "Sheet1"
	index, _ := f.NewSheet(sheetName)
	f.SetActiveSheet(index)

	// Header
	headers := []string{"用户编号", "日期", "00:30", "01:00", "01:30"}
	for i, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(sheetName, cell, h)
	}

	// Data
	data := []string{"C123", "2023-10-01", "100", "200", "300"}
	for i, d := range data {
		cell, _ := excelize.CoordinatesToCellName(i+1, 2)
		f.SetCellValue(sheetName, cell, d)
	}

	if err := f.SaveAs(filename); err != nil {
		t.Fatal(err)
	}
}

func TestExcelImporter_Basic(t *testing.T) {
	filename := "test_import.xlsx"
	createTestExcel(t, filename)
	defer os.Remove(filename)

	// Config
	config := &ExcelImportConfig[TestRow]{
		SheetName: "Sheet1",
	}

	importer := NewExcelImporter(config)
	rows, err := importer.ImportLocal(filename)

	if err != nil {
		t.Fatalf("ImportLocal failed: %v", err)
	}
	if len(rows) != 1 {
		t.Fatalf("Expected 1 row, got %d", len(rows))
	}

	row := rows[0]
	if row.ClientAccount != "C123" {
		t.Errorf("Expected ClientAccount C123, got %s", row.ClientAccount)
	}
	if row.Date != "2023-10-01" {
		t.Errorf("Expected Date 2023-10-01, got %s", row.Date)
	}
	if row.TimeData == nil {
		t.Fatal("Expected TimeData to be initialized")
	}
	if val, ok := row.TimeData["00:30"]; !ok || val != "100" {
		t.Errorf("Expected 00:30=100, got %v", val)
	}
}

func TestExcelImporter_Stream(t *testing.T) {
	filename := "test_import_stream.xlsx"
	createTestExcel(t, filename)
	defer os.Remove(filename)

	config := &ExcelImportConfig[TestRow]{
		SheetName: "Sheet1",
	}

	importer := NewExcelImporter(config)
	ch := importer.ImportStreamLocal(filename)

	var count int
	for res := range ch {
		if res.Error != nil {
			t.Fatalf("Stream error at row %d: %v", res.RowIndex, res.Error)
		}
		
		count++
		row := res.Data
		if row.ClientAccount != "C123" {
			t.Errorf("Expected ClientAccount C123, got %s", row.ClientAccount)
		}
		if row.TimeData == nil {
			t.Fatal("Expected TimeData to be initialized")
		}
		if val, ok := row.TimeData["00:30"]; !ok || val != "100" {
			t.Errorf("Expected 00:30=100, got %v", val)
		}
	}

	if count != 1 {
		t.Fatalf("Expected 1 row, got %d", count)
	}
}
