package exporter

import (
	"bytes"
	"fmt"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// ExcelExportConfig configuration for Excel export
type ExcelExportConfig[T any] struct {
	FileName         string
	SheetName        string
	Headers          []string
	Dropdowns        map[int][]string
	CustomConverters map[string]func(any) any
	TextColumns      map[string]bool
	ColumnWidths     map[string]float64
}

// ExcelExporter generic exporter
type ExcelExporter[T any] struct {
	config   *ExcelExportConfig[T]
	fieldMap map[string]string // Header -> FieldName
}

// NewExcelExporter creates a new exporter instance
func NewExcelExporter[T any](config *ExcelExportConfig[T]) *ExcelExporter[T] {
	if config == nil {
		config = &ExcelExportConfig[T]{}
	}
	if config.SheetName == "" {
		config.SheetName = "Sheet1"
	}
	if config.FileName == "" {
		config.FileName = "export.xlsx"
	}
	if config.TextColumns == nil {
		config.TextColumns = make(map[string]bool)
	}
	if config.ColumnWidths == nil {
		config.ColumnWidths = make(map[string]float64)
	}

	exporter := &ExcelExporter[T]{config: config}
	exporter.parseTags()
	return exporter
}

func (e *ExcelExporter[T]) parseTags() {
	var zero T
	t := reflect.TypeOf(zero)
	if t.Kind() == reflect.Ptr {
		t = t.Elem()
	}
	if t.Kind() != reflect.Struct {
		return
	}

	e.fieldMap = make(map[string]string)
	var inferredHeaders []string

	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		tag := field.Tag.Get("excel")
		if tag == "" || tag == "-" {
			continue
		}

		parts := strings.Split(tag, ",")
		headerName := strings.TrimSpace(parts[0])
		e.fieldMap[headerName] = field.Name
		inferredHeaders = append(inferredHeaders, headerName)

		for _, opt := range parts[1:] {
			opt = strings.TrimSpace(opt)
			if opt == "text" {
				e.config.TextColumns[headerName] = true
			} else if strings.HasPrefix(opt, "width:") {
				valStr := strings.TrimPrefix(opt, "width:")
				if width, err := strconv.ParseFloat(valStr, 64); err == nil {
					e.config.ColumnWidths[headerName] = width
				}
			}
		}
	}

	// Only use inferred headers if config headers are empty
	if len(e.config.Headers) == 0 {
		e.config.Headers = inferredHeaders
	}
}

func (e *ExcelExporter[T]) Export(data []T) (*DownloadResponse, error) {
	f := excelize.NewFile()
	sheetName := e.config.SheetName
	index, _ := f.GetSheetIndex("Sheet1")
	if index != -1 {
		_ = f.SetSheetName("Sheet1", sheetName)
	}
	if err := e.setHeaders(f, sheetName); err != nil {
		return nil, err
	}

	if err := e.setDropdownValidations(f, sheetName); err != nil {
		return nil, err
	}

	if err := e.fillData(f, sheetName, data); err != nil {
		return nil, err
	}

	if err := e.setTextColumnStyle(f, sheetName); err != nil {
		return nil, err
	}

	if err := e.setHeaderStyle(f, sheetName); err != nil {
		return nil, err
	}

	if err := e.setColumnWidths(f, sheetName); err != nil {
		return nil, err
	}

	var buffer bytes.Buffer
	if err := f.Write(&buffer); err != nil {
		return nil, fmt.Errorf("buffer write failed: %v", err)
	}

	content := buffer.Bytes()

	response := &DownloadResponse{
		FileName:    e.config.FileName,
		FileSize:    int64(len(content)),
		ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		Content:     content,
	}

	return response, nil
}

func (e *ExcelExporter[T]) setHeaders(f *excelize.File, sheetName string) error {
	for col, header := range e.config.Headers {
		cell, err := excelize.CoordinatesToCellName(col+1, 1)
		if err != nil {
			return err
		}
		if err := f.SetCellValue(sheetName, cell, header); err != nil {
			return err
		}
	}
	return nil
}

func (e *ExcelExporter[T]) setDropdownValidations(f *excelize.File, sheetName string) error {
	if e.config.Dropdowns == nil {
		return nil
	}

	for colIndex, options := range e.config.Dropdowns {
		if len(options) == 0 {
			continue
		}

		colName, err := excelize.ColumnNumberToName(colIndex + 1)
		if err != nil {
			return err
		}

		dvRange := excelize.NewDataValidation(true)
		dvRange.SetSqref(fmt.Sprintf("%s2:%s1000", colName, colName))
		_ = dvRange.SetDropList(options)
		title := "Error"
		msg := "Invalid input"
		dvRange.SetError(excelize.DataValidationErrorStyleWarning, title, msg)

		if err := f.AddDataValidation(sheetName, dvRange); err != nil {
			return err
		}
	}

	return nil
}

func (e *ExcelExporter[T]) getTextCellStyle(f *excelize.File) (int, error) {
	// NumFmt 49 is '@' (Text)
	return f.NewStyle(&excelize.Style{
		NumFmt: 49,
		Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "center",
		},
	})
}

func (e *ExcelExporter[T]) setTextColumnStyle(f *excelize.File, sheetName string) error {
	if len(e.config.TextColumns) == 0 {
		return nil
	}

	styleID, err := e.getTextCellStyle(f)
	if err != nil {
		return err
	}

	for colIndex, header := range e.config.Headers {
		if e.config.TextColumns[header] {
			colName, err := excelize.ColumnNumberToName(colIndex + 1)
			if err != nil {
				return err
			}

			startCell := fmt.Sprintf("%s2", colName)
			endCell := fmt.Sprintf("%s10000", colName)

			if err := f.SetCellStyle(sheetName, startCell, endCell, styleID); err != nil {
				return err
			}
		}
	}
	return nil
}

func (e *ExcelExporter[T]) fillData(f *excelize.File, sheetName string, data []T) error {
	if len(data) == 0 {
		return nil
	}

	for rowIndex, item := range data {
		if err := e.fillRow(f, sheetName, rowIndex+2, item); err != nil {
			return fmt.Errorf("row %d error: %v", rowIndex+2, err)
		}
	}

	return nil
}

func (e *ExcelExporter[T]) fillRow(f *excelize.File, sheetName string, row int, item T) error {
	itemValue := reflect.ValueOf(item)
	if itemValue.Kind() == reflect.Ptr {
		itemValue = itemValue.Elem()
	}

	for colIndex, header := range e.config.Headers {
		cell, err := excelize.CoordinatesToCellName(colIndex+1, row)
		if err != nil {
			return err
		}

		fieldName, exists := e.fieldMap[header]
		if !exists {
			continue
		}

		fieldValue := itemValue.FieldByName(fieldName)
		if !fieldValue.IsValid() {
			continue
		}

		value := e.getFieldValue(fieldName, fieldValue)
		if e.config.TextColumns[header] {
			valueStr := fmt.Sprintf("%v", value)
			if err := f.SetCellStr(sheetName, cell, valueStr); err != nil {
				return err
			}
		} else {
			if err := f.SetCellValue(sheetName, cell, value); err != nil {
				return err
			}
		}
	}

	return nil
}

func (e *ExcelExporter[T]) getFieldValue(fieldName string, fieldValue reflect.Value) interface{} {
	if !fieldValue.IsValid() {
		return ""
	}

	// Handle pointer
	if fieldValue.Kind() == reflect.Ptr {
		if fieldValue.IsNil() {
			return ""
		}
		fieldValue = fieldValue.Elem()
	}

	// Check custom converter
	if converter, exists := e.config.CustomConverters[fieldName]; exists {
		// Pass the underlying value
		return converter(fieldValue.Interface())
	}

	// Default handling
	switch fieldValue.Kind() {
	case reflect.Struct:
		if fieldValue.Type() == reflect.TypeOf(time.Time{}) {
			if timeVal, ok := fieldValue.Interface().(time.Time); ok {
				return timeVal.Format("2006-01-02 15:04:05")
			}
		}
	}

	return fieldValue.Interface()
}

func (e *ExcelExporter[T]) setHeaderStyle(f *excelize.File, sheetName string) error {
	if len(e.config.Headers) == 0 {
		return nil
	}

	styleID, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{
			Bold:  true,
			Color: "FFFFFF",
			Size:  12,
		},
		Fill: excelize.Fill{
			Type:    "pattern",
			Color:   []string{"366092"},
			Pattern: 1,
		},
		Alignment: &excelize.Alignment{
			Horizontal: "center",
			Vertical:   "center",
		},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
		},
	})
	if err != nil {
		return err
	}

	startCell, _ := excelize.CoordinatesToCellName(1, 1)
	endCell, _ := excelize.CoordinatesToCellName(len(e.config.Headers), 1)

	return f.SetCellStyle(sheetName, startCell, endCell, styleID)
}

func (e *ExcelExporter[T]) setColumnWidths(f *excelize.File, sheetName string) error {
	// Default auto width logic or explicit width
	for colIndex, header := range e.config.Headers {
		colName, _ := excelize.ColumnNumberToName(colIndex + 1)
		
		if width, ok := e.config.ColumnWidths[header]; ok {
			if err := f.SetColWidth(sheetName, colName, colName, width); err != nil {
				return err
			}
		} else {
			// Default width
			if err := f.SetColWidth(sheetName, colName, colName, 15); err != nil {
				return err
			}
		}
	}
	return nil
}
