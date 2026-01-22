package exporter

import (
	"math"
	"os"
	"testing"
)

type TestExportData struct {
	Name  string  `excel:"姓名,text"`
	Age   int     `excel:"年龄"`
	Score float64 `excel:"分数,width:20"`
}

func TestExcelExporter_ExportExample(t *testing.T) {
	// 1. Prepare data
	data := []TestExportData{
		{Name: "张三", Age: 25, Score: 88.5},
		{Name: "李四", Age: 30, Score: 92.0},
		{Name: "王五", Age: 28, Score: 76.5},
	}

	// 2. Configure exporter (Minimal config)
	config := &ExcelExportConfig[TestExportData]{
		FileName: "test_export.xlsx",
	}
	exporter := NewExcelExporter(config)

	// 3. Execute export
	resp, err := exporter.Export(data)
	if err != nil {
		t.Fatalf("Export failed: %v", err)
	}

	// 4. Verify results
	if resp.FileName != "test_export.xlsx" {
		t.Errorf("Expected filename test_export.xlsx, got %s", resp.FileName)
	}
	if len(resp.Content) == 0 {
		t.Error("Exported content is empty")
	}

	// Optional: Write file for manual check
	if err := os.WriteFile("test_export_output.xlsx", resp.Content, 0644); err != nil {
		t.Logf("Failed to write output file: %v", err)
	} else {
		t.Log("Successfully wrote test_export_output.xlsx")
		defer os.Remove("test_export_output.xlsx")
	}
}

// Simulate user's forecast export scenario
type ForecastExportItem struct {
	Name      string   `excel:"名称,text"`
	Value0030 *float64 `excel:"00:30"`
	Value0100 *float64 `excel:"01:00"`
}

func TestExcelExporter_UserScenario(t *testing.T) {
	val1 := 100.12345
	val2 := 200.67891
	data := []ForecastExportItem{
		{Name: "User1", Value0030: &val1, Value0100: nil},
		{Name: "User2", Value0030: nil, Value0100: &val2},
	}

	// Custom converter for 4 decimal places
	keep4Decimals := func(a any) any {
		if a == nil {
			return nil
		}
		if v, ok := a.(*float64); ok {
			if v == nil {
				return nil
			}
			return math.Round(*v*10000) / 10000
		}
		if v, ok := a.(float64); ok {
			return math.Round(v*10000) / 10000
		}
		return a
	}

	config := &ExcelExportConfig[ForecastExportItem]{
		FileName: "forecast.xlsx",
		CustomConverters: map[string]func(any) any{
			"Value0030": keep4Decimals,
			"Value0100": keep4Decimals,
		},
	}

	exporter := NewExcelExporter(config)
	resp, err := exporter.Export(data)
	if err != nil {
		t.Fatalf("Export failed: %v", err)
	}
	if len(resp.Content) == 0 {
		t.Error("Content empty")
	}
	// os.WriteFile("forecast_output.xlsx", resp.Content, 0644)
	// defer os.Remove("forecast_output.xlsx")
}
