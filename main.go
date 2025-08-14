package excel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

// File represents an Excel file wrapper
type File struct {
	file *excelize.File
}

type (
	Alignment  excelize.Alignment
	Border     excelize.Border
	Font       excelize.Font
	Fill       excelize.Fill
	Protection excelize.Protection
	Style      struct {
		Border        []Border
		Fill          Fill
		Font          *Font
		Alignment     *Alignment
		Protection    *Protection
		NumFmt        int
		DecimalPlaces *int
		CustomNumFmt  *string
		NegRed        bool
	}
)

// CellData represents data for a single cell
type CellData struct {
	Row   int
	Col   int
	Value interface{}
	Sheet string
}

// RowData represents data for an entire row
type RowData struct {
	Row    int
	Values []interface{}
	Sheet  string
}

// NewExcelFile creates a new Excel file
func NewExcelFile() *File {
	return &File{
		file: excelize.NewFile(),
	}
}

// OpenExcelFile opens an existing Excel file
func OpenExcelFile(filename string) (*File, error) {
	file, err := excelize.OpenFile(filename)
	if err != nil {
		return nil, fmt.Errorf("failed to open file %s: %w", filename, err)
	}
	return &File{file: file}, nil
}

// CreateSheet creates a new worksheet
func (f *File) CreateSheet(sheetName string) error {
	_, err := f.file.NewSheet(sheetName)
	if err != nil {
		return fmt.Errorf("failed to create sheet %s: %w", sheetName, err)
	}
	return nil
}

// DeleteSheet deletes a worksheet
func (f *File) DeleteSheet(sheetName string) error {
	err := f.file.DeleteSheet(sheetName)
	if err != nil {
		return fmt.Errorf("failed to delete sheet %s: %w", sheetName, err)
	}
	return nil
}

// GetSheetList returns all sheet names
func (f *File) GetSheetList() []string {
	return f.file.GetSheetList()
}

// SetActiveSheet sets the active sheet
func (f *File) SetActiveSheet(sheetName string) error {
	index, err := f.file.GetSheetIndex(sheetName)
	if err != nil {
		return fmt.Errorf("failed to get sheet index for %s: %w", sheetName, err)
	}
	f.file.SetActiveSheet(index)
	return nil
}

// SetCellValue sets a single cell value
func (f *File) SetCellValue(sheet string, row, col int, value interface{}) error {
	cellName := f.getCellName(row, col)
	err := f.file.SetCellValue(sheet, cellName, value)
	if err != nil {
		return fmt.Errorf("failed to set cell value at %s: %w", cellName, err)
	}
	return nil
}

// GetCellValue gets a single cell value
func (f *File) GetCellValue(sheet string, row, col int) (string, error) {
	cellName := f.getCellName(row, col)
	value, err := f.file.GetCellValue(sheet, cellName)
	if err != nil {
		return "", fmt.Errorf("failed to get cell value at %s: %w", cellName, err)
	}
	return value, nil
}

// SetRowValues sets values for an entire row
func (f *File) SetRowValues(sheet string, row int, values []interface{}) error {
	for col, value := range values {
		err := f.SetCellValue(sheet, row, col+1, value)
		if err != nil {
			return err
		}
	}
	return nil
}

// GetRowValues gets values for an entire row up to the last non-empty cell
func (f *File) GetRowValues(sheet string, row int) ([]string, error) {
	rows, err := f.file.GetRows(sheet)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows from sheet %s: %w", sheet, err)
	}

	if row > len(rows) || row < 1 {
		return []string{}, nil
	}

	return rows[row-1], nil
}

// SetHeaders sets header row (first row)
func (f *File) SetHeaders(sheet string, headers []string) error {
	values := make([]interface{}, len(headers))
	for i, header := range headers {
		values[i] = header
	}
	return f.SetRowValues(sheet, 1, values)
}

// GetHeaders gets header row (first row)
func (f *File) GetHeaders(sheet string) ([]string, error) {
	return f.GetRowValues(sheet, 1)
}

// WriteData writes multiple rows of data starting from a specific row
func (f *File) WriteData(sheet string, startRow int, data [][]interface{}) error {
	for i, row := range data {
		err := f.SetRowValues(sheet, startRow+i, row)
		if err != nil {
			return err
		}
	}
	return nil
}

// ReadData reads all data from a sheet
func (f *File) ReadData(sheet string) ([][]string, error) {
	rows, err := f.file.GetRows(sheet)
	if err != nil {
		return nil, fmt.Errorf("failed to read data from sheet %s: %w", sheet, err)
	}
	return rows, nil
}

// ReadDataWithHeaders reads data and returns headers separately
func (f *File) ReadDataWithHeaders(sheet string) (headers []string, data [][]string, err error) {
	allData, err := f.ReadData(sheet)
	if err != nil {
		return nil, nil, err
	}

	if len(allData) == 0 {
		return []string{}, [][]string{}, nil
	}

	headers = allData[0]
	if len(allData) > 1 {
		data = allData[1:]
	} else {
		data = [][]string{}
	}

	return headers, data, nil
}

// SetCellStyle applies formatting to a cell
func (f *File) SetCellStyle(sheet string, row, col int, styleID int) error {
	cellName := f.getCellName(row, col)
	err := f.file.SetCellStyle(sheet, cellName, cellName, styleID)
	if err != nil {
		return fmt.Errorf("failed to set cell style at %s: %w", cellName, err)
	}
	return nil
}

// CreateStyle creates a new style and returns its ID
func (f *File) CreateStyle(style *Style) (int, error) {
	var border = make([]excelize.Border, len(style.Border))
	for _, b := range style.Border {
		border = append(border, (excelize.Border)(b))
	}
	exStyle := &excelize.Style{
		Border:        border,
		Fill:          (excelize.Fill)(style.Fill),
		Font:          (*excelize.Font)(style.Font),
		Alignment:     (*excelize.Alignment)(style.Alignment),
		Protection:    (*excelize.Protection)(style.Protection),
		NumFmt:        style.NumFmt,
		DecimalPlaces: style.DecimalPlaces,
		CustomNumFmt:  style.CustomNumFmt,
		NegRed:        style.NegRed,
	}

	styleID, err := f.file.NewStyle(exStyle)
	if err != nil {
		return 0, fmt.Errorf("failed to create style: %w", err)
	}
	return styleID, nil
}

// AutoFilter sets auto filter for a range
func (f *File) AutoFilter(sheet string, startRow, startCol, endRow, endCol int) error {
	startCell := f.getCellName(startRow, startCol)
	endCell := f.getCellName(endRow, endCol)
	rangeRef := fmt.Sprintf("%s:%s", startCell, endCell)

	err := f.file.AutoFilter(sheet, rangeRef, nil)
	if err != nil {
		return fmt.Errorf("failed to set auto filter: %w", err)
	}
	return nil
}

// SetColumnWidth sets the width of a column
func (f *File) SetColumnWidth(sheet string, col int, width float64) error {
	colName := f.getColumnName(col)
	err := f.file.SetColWidth(sheet, colName, colName, width)
	if err != nil {
		return fmt.Errorf("failed to set column width: %w", err)
	}
	return nil
}

// SaveAs saves the Excel file with a specific filename
func (f *File) SaveAs(filename string) error {
	err := f.file.SaveAs(filename)
	if err != nil {
		return fmt.Errorf("failed to save file as %s: %w", filename, err)
	}
	return nil
}

// Save saves the Excel file (only works if opened from an existing file)
func (f *File) Save() error {
	err := f.file.Save()
	if err != nil {
		return fmt.Errorf("failed to save file: %w", err)
	}
	return nil
}

// Close closes the Excel file
func (f *File) Close() error {
	return f.file.Close()
}

// Helper function to convert row, col to Excel cell name (e.g., A1, B2)
func (f *File) getCellName(row, col int) string {
	cellName, _ := excelize.CoordinatesToCellName(col, row)
	return cellName
}

// Helper function to get column name from column number
func (f *File) getColumnName(col int) string {
	var result string
	for col > 0 {
		col--
		result = string(rune('A'+col%26)) + result
		col /= 26
	}
	return result
}

// Utility functions for common operations

// CreateSimpleTable creates a simple table with headers and data
func (f *File) CreateSimpleTable(sheet string, headers []string, data [][]interface{}) error {
	// Set headers
	err := f.SetHeaders(sheet, headers)
	if err != nil {
		return err
	}

	// Set data starting from row 2
	err = f.WriteData(sheet, 2, data)
	if err != nil {
		return err
	}

	// Auto-size columns (set reasonable width)
	for i := range headers {
		err = f.SetColumnWidth(sheet, i+1, 15)
		if err != nil {
			return err
		}
	}

	return nil
}

// ConvertToMap converts sheet data to a slice of maps using headers as keys
func (f *File) ConvertToMap(sheet string) ([]map[string]string, error) {
	headers, data, err := f.ReadDataWithHeaders(sheet)
	if err != nil {
		return nil, err
	}

	var result []map[string]string
	for _, row := range data {
		rowMap := make(map[string]string)
		for i, header := range headers {
			if i < len(row) {
				rowMap[header] = row[i]
			} else {
				rowMap[header] = ""
			}
		}
		result = append(result, rowMap)
	}

	return result, nil
}
