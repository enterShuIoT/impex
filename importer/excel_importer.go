package importer

import (
	"fmt"
	"io"
	"net/http"
	"path/filepath"
	"reflect"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// ExcelImportConfig configuration for Excel import
type ExcelImportConfig[T any] struct {
	SheetName        string
	StartRow         int
	HeaderRow        int
	FieldMappings    map[string]string            // Excel Column -> Struct Field
	DefaultValues    map[string]any
	Validators       map[string]func(any) error
	CustomConverters map[string]func(string) (any, error)
	SkipRows         map[int]bool
	RowHook          func(*T, []string, map[string]int) error
}

// ExcelImporter generic importer
type ExcelImporter[T any] struct {
	config        *ExcelImportConfig[T]
	dynamicField  string
	dynamicFilter *regexp.Regexp
}

// NewExcelImporter creates a new importer instance
func NewExcelImporter[T any](config *ExcelImportConfig[T]) *ExcelImporter[T] {
	if config == nil {
		config = &ExcelImportConfig[T]{}
	}
	if config.StartRow == 0 {
		config.StartRow = 2
	}
	if config.HeaderRow == 0 {
		config.HeaderRow = 1
	}

	importer := &ExcelImporter[T]{config: config}
	importer.parseTags()
	return importer
}

func (importer *ExcelImporter[T]) parseTags() {
	var zero T
	t := reflect.TypeOf(zero)
	if t.Kind() == reflect.Ptr {
		t = t.Elem()
	}
	if t.Kind() != reflect.Struct {
		return
	}

	if importer.config.FieldMappings == nil {
		importer.config.FieldMappings = make(map[string]string)
	}

	for i := 0; i < t.NumField(); i++ {
		field := t.Field(i)
		tag := field.Tag.Get("excel")
		if tag == "" {
			continue
		}

		parts := strings.Split(tag, ",")
		head := strings.TrimSpace(parts[0])

		if head == "*" || head == "extra" {
			importer.dynamicField = field.Name
			for _, part := range parts[1:] {
				part = strings.TrimSpace(part)
				if strings.HasPrefix(part, "pattern:") {
					pattern := strings.TrimPrefix(part, "pattern:")
					if regex, err := regexp.Compile(pattern); err == nil {
						importer.dynamicFilter = regex
					}
				}
			}
			continue
		}

		importer.config.FieldMappings[head] = field.Name
	}
}

func (importer *ExcelImporter[T]) Import(url string) ([]T, error) {
	body, _, err := downloadFromUrl(url)
	if err != nil {
		return nil, fmt.Errorf("download failed: %v", err)
	}
	f, err := excelize.OpenReader(body)
	if err != nil {
		return nil, fmt.Errorf("open excel failed: %v", err)
	}
	defer f.Close()
	return importer.importFromFile(f)
}

func (importer *ExcelImporter[T]) ImportLocal(filePath string) ([]T, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("open excel failed: %v", err)
	}
	defer f.Close()
	return importer.importFromFile(f)
}

func (importer *ExcelImporter[T]) ImportStream(url string) <-chan ImportResult[T] {
	ch := make(chan ImportResult[T])

	go func() {
		defer close(ch)

		body, _, err := downloadFromUrl(url)
		if err != nil {
			ch <- ImportResult[T]{Error: fmt.Errorf("download failed: %v", err)}
			return
		}
		
		f, err := excelize.OpenReader(body)
		if err != nil {
			ch <- ImportResult[T]{Error: fmt.Errorf("open excel failed: %v", err)}
			return
		}
		defer f.Close()

		importer.streamRows(f, ch)
	}()

	return ch
}

func (importer *ExcelImporter[T]) ImportStreamLocal(filePath string) <-chan ImportResult[T] {
	ch := make(chan ImportResult[T])

	go func() {
		defer close(ch)

		f, err := excelize.OpenFile(filePath)
		if err != nil {
			ch <- ImportResult[T]{Error: fmt.Errorf("open excel failed: %v", err)}
			return
		}
		defer f.Close()

		importer.streamRows(f, ch)
	}()

	return ch
}

func (importer *ExcelImporter[T]) streamRows(f *excelize.File, ch chan<- ImportResult[T]) {
	sheetName := importer.config.SheetName
	if sheetName == "" {
		if f.SheetCount < 1 {
			ch <- ImportResult[T]{Error: fmt.Errorf("excel file has no sheets")}
			return
		}
		sheetName = f.GetSheetName(0)
	}

	rows, err := f.Rows(sheetName)
	if err != nil {
		ch <- ImportResult[T]{Error: fmt.Errorf("read sheet failed: %v", err)}
		return
	}
	defer rows.Close()

	var columnIndexMap map[string]int
	rowIndex := 0

	for rows.Next() {
		rowIndex++
		
		// Skip rows
		if importer.config.SkipRows[rowIndex] {
			continue
		}

		// Read row columns
		row, err := rows.Columns()
		if err != nil {
			ch <- ImportResult[T]{RowIndex: rowIndex, Error: fmt.Errorf("read row %d failed: %v", rowIndex, err)}
			return
		}

		// Handle Header
		if rowIndex == importer.config.HeaderRow {
			columnIndexMap = importer.buildColumnIndexMap(row)
			
			// Validate headers
			missingColumns := make([]string, 0)
			for excelCol := range importer.config.FieldMappings {
				if _, exists := columnIndexMap[excelCol]; !exists {
					missingColumns = append(missingColumns, excelCol)
				}
			}
			if len(missingColumns) > 0 {
				ch <- ImportResult[T]{RowIndex: rowIndex, Error: fmt.Errorf("missing columns: %s", strings.Join(missingColumns, ", "))}
				return
			}
			continue
		}

		// Skip if before StartRow
		if rowIndex < importer.config.StartRow {
			continue
		}

		if importer.isEmptyRow(row) {
			continue
		}

		instance, err := importer.parseRow(row, columnIndexMap)
		if err != nil {
			ch <- ImportResult[T]{RowIndex: rowIndex, Error: err}
			continue // Continue processing other rows
		}

		ch <- ImportResult[T]{RowIndex: rowIndex, Data: instance}
	}
}

func (importer *ExcelImporter[T]) importFromFile(f *excelize.File) ([]T, error) {
	sheetName := importer.config.SheetName
	if sheetName == "" {
		if f.SheetCount < 1 {
			return nil, fmt.Errorf("excel file has no sheets")
		}
		sheetName = f.GetSheetName(0)
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("read sheet failed: %v", err)
	}

	if len(rows) < importer.config.HeaderRow {
		return nil, fmt.Errorf("insufficient rows")
	}

	headerRow := rows[importer.config.HeaderRow-1]
	columnIndexMap := importer.buildColumnIndexMap(headerRow)

	missingColumns := make([]string, 0)
	for excelCol := range importer.config.FieldMappings {
		if _, exists := columnIndexMap[excelCol]; !exists {
			missingColumns = append(missingColumns, excelCol)
		}
	}
	if len(missingColumns) > 0 {
		return nil, fmt.Errorf("missing columns: %s", strings.Join(missingColumns, ", "))
	}

	var result []T

	for i := importer.config.StartRow - 1; i < len(rows); i++ {
		if importer.config.SkipRows[i+1] {
			continue
		}

		row := rows[i]
		if importer.isEmptyRow(row) {
			continue
		}

		instance, err := importer.parseRow(row, columnIndexMap)
		if err != nil {
			return nil, fmt.Errorf("row %d error: %v", i+1, err)
		}

		result = append(result, instance)
	}

	return result, nil
}

func (importer *ExcelImporter[T]) parseRow(row []string, columnIndexMap map[string]int) (T, error) {
	var instance T
	val := reflect.ValueOf(&instance)
	if val.Kind() == reflect.Ptr {
		val = val.Elem()
	}
	if val.Kind() == reflect.Ptr {
		if val.IsNil() {
			val.Set(reflect.New(val.Type().Elem()))
		}
		val = val.Elem()
	}

	if err := importer.fillStruct(val, row, columnIndexMap, &instance); err != nil {
		return instance, err
	}

	if err := importer.validateData(val); err != nil {
		return instance, err
	}
	return instance, nil
}

func (importer *ExcelImporter[T]) buildColumnIndexMap(headerRow []string) map[string]int {
	indexMap := make(map[string]int)
	for idx, cellValue := range headerRow {
		cleanName := strings.Trim(strings.TrimSpace(cellValue), "*")
		indexMap[cleanName] = idx
	}
	return indexMap
}

func (importer *ExcelImporter[T]) isEmptyRow(row []string) bool {
	for _, cell := range row {
		if strings.TrimSpace(cell) != "" {
			return false
		}
	}
	return true
}

func (importer *ExcelImporter[T]) fillStruct(val reflect.Value, row []string, columnIndexMap map[string]int, instance *T) error {
	t := val.Type()
	usedColumns := make(map[int]bool)

	for i := 0; i < val.NumField(); i++ {
		field := val.Field(i)
		fieldType := t.Field(i)

		if !field.CanSet() {
			continue
		}

		if fieldType.Name == importer.dynamicField {
			continue
		}

		excelColumn := importer.findExcelColumnForField(fieldType)
		if excelColumn == "" {
			continue
		}

		colIndex, exists := columnIndexMap[excelColumn]
		if !exists {
			if defaultValue, hasDefault := importer.config.DefaultValues[fieldType.Name]; hasDefault {
				if err := importer.setFieldValue(field, defaultValue); err != nil {
					return err
				}
			}
			continue
		}

		usedColumns[colIndex] = true

		var cellValue string
		if colIndex < len(row) {
			cellValue = strings.TrimSpace(row[colIndex])
		}

		if cellValue == "" {
			if defaultValue, hasDefault := importer.config.DefaultValues[fieldType.Name]; hasDefault {
				if err := importer.setFieldValue(field, defaultValue); err != nil {
					return err
				}
			}
			continue
		}

		if err := importer.convertAndSetField(field, fieldType, cellValue); err != nil {
			return fmt.Errorf("field %s conversion failed: %v", fieldType.Name, err)
		}
	}

	// Handle dynamic field
	if importer.dynamicField != "" {
		field := val.FieldByName(importer.dynamicField)
		if field.IsValid() && field.CanSet() && field.Kind() == reflect.Map {
			if field.IsNil() {
				field.Set(reflect.MakeMap(field.Type()))
			}
			
			// Only support map[string]string or map[string]any
			keyKind := field.Type().Key().Kind()
			elemKind := field.Type().Elem().Kind()
			
			if keyKind == reflect.String {
				for colName, colIdx := range columnIndexMap {
					if !usedColumns[colIdx] && colIdx < len(row) {
						// Apply dynamic filter if set
						if importer.dynamicFilter != nil {
                            matched := importer.dynamicFilter.MatchString(colName)
                            if !matched {
							    continue
                            }
						}

						cellVal := strings.TrimSpace(row[colIdx])
						if cellVal != "" {
							var valToSet reflect.Value
							var err error

							switch elemKind {
							case reflect.String:
								valToSet = reflect.ValueOf(cellVal)
							case reflect.Interface:
								valToSet = reflect.ValueOf(cellVal)
							case reflect.Float64, reflect.Float32:
								if f, e := strconv.ParseFloat(cellVal, 64); e == nil {
									valToSet = reflect.ValueOf(f).Convert(field.Type().Elem())
								} else {
									err = e
								}
							case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
								if i, e := strconv.ParseInt(cellVal, 10, 64); e == nil {
									valToSet = reflect.ValueOf(i).Convert(field.Type().Elem())
								} else {
									err = e
								}
							case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
								if u, e := strconv.ParseUint(cellVal, 10, 64); e == nil {
									valToSet = reflect.ValueOf(u).Convert(field.Type().Elem())
								} else {
									err = e
								}
							case reflect.Bool:
								b := strings.ToLower(cellVal) == "true" || cellVal == "1" || cellVal == "是"
								valToSet = reflect.ValueOf(b)
							}

							if err == nil && valToSet.IsValid() {
								field.SetMapIndex(reflect.ValueOf(colName), valToSet)
							}
						}
					}
				}
			}
		}
	}

	if importer.config.RowHook != nil {
		if err := importer.config.RowHook(instance, row, columnIndexMap); err != nil {
			return err
		}
	}

	return nil
}

func (importer *ExcelImporter[T]) findExcelColumnForField(field reflect.StructField) string {
	for excelCol, structField := range importer.config.FieldMappings {
		if structField == field.Name {
			return excelCol
		}
	}
	return ""
}

func (importer *ExcelImporter[T]) convertAndSetField(field reflect.Value, fieldType reflect.StructField, cellValue string) error {
	if converter, exists := importer.config.CustomConverters[fieldType.Name]; exists {
		convertedValue, err := converter(cellValue)
		if err != nil {
			return err
		}
		return importer.setFieldValue(field, convertedValue)
	}
	var convertedValue interface{}
	switch field.Kind() {
	case reflect.String:
		convertedValue = cellValue
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		if cellValue == "" {
			convertedValue = 0
		} else {
			intVal, err := strconv.ParseInt(cellValue, 10, 64)
			if err != nil {
				return fmt.Errorf("invalid integer: %s", cellValue)
			}
			convertedValue = intVal
		}
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		if cellValue == "" {
			convertedValue = uint64(0)
		} else {
			uintVal, err := strconv.ParseUint(cellValue, 10, 64)
			if err != nil {
				return fmt.Errorf("invalid uint: %s", cellValue)
			}
			convertedValue = uintVal
		}
	case reflect.Float32, reflect.Float64:
		if cellValue == "" {
			convertedValue = 0.0
		} else {
			floatVal, err := strconv.ParseFloat(cellValue, 64)
			if err != nil {
				return fmt.Errorf("invalid float: %s", cellValue)
			}
			convertedValue = floatVal
		}
	case reflect.Bool:
		convertedValue = strings.ToLower(cellValue) == "true" || cellValue == "1" || cellValue == "是"
	case reflect.Struct:
		if fieldType.Type == reflect.TypeOf(time.Time{}) {
			timeVal, err := time.Parse("2006-01-02", cellValue)
			if err != nil {
				timeVal, err = time.Parse("2006/01/02", cellValue)
				if err != nil {
					return fmt.Errorf("invalid time: %s", cellValue)
				}
			}
			convertedValue = timeVal
		} else {
			return fmt.Errorf("unsupported struct type: %s", fieldType.Type.Name())
		}
	default:
		return fmt.Errorf("unsupported kind: %s", field.Kind())
	}
	return importer.setFieldValue(field, convertedValue)
}

func (importer *ExcelImporter[T]) setFieldValue(field reflect.Value, value interface{}) error {
	if value == nil {
		return nil
	}
	val := reflect.ValueOf(value)
	
	// Handle integer type mismatches (e.g. int64 to int)
	if val.Kind() != field.Kind() && val.Type().ConvertibleTo(field.Type()) {
		field.Set(val.Convert(field.Type()))
		return nil
	}

	if !val.Type().AssignableTo(field.Type()) {
		return fmt.Errorf("type mismatch: cannot assign %v to %v", val.Type(), field.Type())
	}
	
	field.Set(val)
	return nil
}

func (importer *ExcelImporter[T]) validateData(instance reflect.Value) error {
	for i := 0; i < instance.NumField(); i++ {
		field := instance.Field(i)
		fieldType := instance.Type().Field(i)

		if validator, exists := importer.config.Validators[fieldType.Name]; exists {
			if err := validator(field.Interface()); err != nil {
				return fmt.Errorf("validation failed: %v", err)
			}
		}
	}
	return nil
}

func downloadFromUrl(url string) (io.ReadCloser, string, error) {
	resp, err := http.Get(url)
	if err != nil {
		return nil, "", fmt.Errorf("request failed: %w", err)
	}

	if resp.StatusCode != http.StatusOK {
		_ = resp.Body.Close()
		return nil, "", fmt.Errorf("status code: %d", resp.StatusCode)
	}
	var fileName string
	disp := resp.Header.Get("Content-Disposition")
	if disp != "" {
		re := regexp.MustCompile(`filename="([^"]+)"`)
		matches := re.FindStringSubmatch(disp)
		if len(matches) > 1 {
			fileName = matches[1]
		}
	}
	if fileName == "" {
		fileName = filepath.Base(resp.Request.URL.Path)
	}

	return resp.Body, fileName, nil
}
