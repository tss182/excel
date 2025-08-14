package excel

import (
	"os"
	"testing"
)

func TestNewExcelFile(t *testing.T) {
	f := NewExcelFile()
	if f == nil {
		t.Fatal("NewExcelFile returned nil")
	}
	if f.file == nil {
		t.Fatal("Excel file is nil")
	}
}

func TestFileOperations(t *testing.T) {
	f := NewExcelFile()
	defer f.Close()

	// Test creating and deleting sheets
	t.Run("Sheet Operations", func(t *testing.T) {
		err := f.CreateSheet("TestSheet")
		if err != nil {
			t.Errorf("Failed to create sheet: %v", err)
		}

		sheets := f.GetSheetList()
		found := false
		for _, sheet := range sheets {
			if sheet == "TestSheet" {
				found = true
				break
			}
		}
		if !found {
			t.Error("Created sheet not found in sheet list")
		}

		err = f.SetActiveSheet("TestSheet")
		if err != nil {
			t.Errorf("Failed to set active sheet: %v", err)
		}

		err = f.DeleteSheet("TestSheet")
		if err != nil {
			t.Errorf("Failed to delete sheet: %v", err)
		}
	})

	// Test cell operations
	t.Run("Cell Operations", func(t *testing.T) {
		sheet := "Sheet1"
		err := f.SetCellValue(sheet, 1, 1, "Test")
		if err != nil {
			t.Errorf("Failed to set cell value: %v", err)
		}

		value, err := f.GetCellValue(sheet, 1, 1)
		if err != nil {
			t.Errorf("Failed to get cell value: %v", err)
		}
		if value != "Test" {
			t.Errorf("Expected cell value 'Test', got '%s'", value)
		}
	})

	// Test row operations
	t.Run("Row Operations", func(t *testing.T) {
		sheet := "Sheet1"
		values := []interface{}{"Col1", "Col2", "Col3"}
		err := f.SetRowValues(sheet, 2, values)
		if err != nil {
			t.Errorf("Failed to set row values: %v", err)
		}

		rowValues, err := f.GetRowValues(sheet, 2)
		if err != nil {
			t.Errorf("Failed to get row values: %v", err)
		}
		if len(rowValues) != len(values) {
			t.Errorf("Expected %d values, got %d", len(values), len(rowValues))
		}
	})

	// Test headers
	t.Run("Header Operations", func(t *testing.T) {
		sheet := "Sheet1"
		headers := []string{"Header1", "Header2", "Header3"}
		err := f.SetHeaders(sheet, headers)
		if err != nil {
			t.Errorf("Failed to set headers: %v", err)
		}

		readHeaders, err := f.GetHeaders(sheet)
		if err != nil {
			t.Errorf("Failed to get headers: %v", err)
		}
		if len(readHeaders) != len(headers) {
			t.Errorf("Expected %d headers, got %d", len(headers), len(readHeaders))
		}
	})

	// Test data operations
	t.Run("Data Operations", func(t *testing.T) {
		sheet := "Sheet1"
		data := [][]interface{}{
			{"A1", "B1", "C1"},
			{"A2", "B2", "C2"},
		}
		err := f.WriteData(sheet, 3, data)
		if err != nil {
			t.Errorf("Failed to write data: %v", err)
		}

		readData, err := f.ReadData(sheet)
		if err != nil {
			t.Errorf("Failed to read data: %v", err)
		}
		if len(readData) < len(data) {
			t.Errorf("Expected at least %d rows, got %d", len(data), len(readData))
		}
	})
}

func TestFileIO(t *testing.T) {
	// Test saving and opening files
	t.Run("File IO Operations", func(t *testing.T) {
		f := NewExcelFile()
		defer f.Close()

		// Add some test data
		err := f.SetCellValue("Sheet1", 1, 1, "Test Data")
		if err != nil {
			t.Errorf("Failed to set cell value: %v", err)
		}

		// Save the file
		tempFile := "test_excel.xlsx"
		err = f.SaveAs(tempFile)
		if err != nil {
			t.Errorf("Failed to save file: %v", err)
		}

		// Open the saved file
		f2, err := OpenExcelFile(tempFile)
		if err != nil {
			t.Errorf("Failed to open file: %v", err)
		}
		defer f2.Close()

		// Verify the data
		value, err := f2.GetCellValue("Sheet1", 1, 1)
		if err != nil {
			t.Errorf("Failed to get cell value from opened file: %v", err)
		}
		if value != "Test Data" {
			t.Errorf("Expected 'Test Data', got '%s'", value)
		}

		// Clean up
		os.Remove(tempFile)
	})
}

func TestUtilityFunctions(t *testing.T) {
	f := NewExcelFile()
	defer f.Close()

	// Test CreateSimpleTable
	t.Run("Create Simple Table", func(t *testing.T) {
		headers := []string{"Name", "Age", "City"}
		data := [][]interface{}{
			{"John", 30, "New York"},
			{"Alice", 25, "London"},
		}

		err := f.CreateSimpleTable("Sheet1", headers, data)
		if err != nil {
			t.Errorf("Failed to create simple table: %v", err)
		}

		// Verify headers
		readHeaders, err := f.GetHeaders("Sheet1")
		if err != nil {
			t.Errorf("Failed to read headers: %v", err)
		}
		if len(readHeaders) != len(headers) {
			t.Errorf("Expected %d headers, got %d", len(headers), len(readHeaders))
		}
	})

	// Test ConvertToMap
	t.Run("Convert To Map", func(t *testing.T) {
		result, err := f.ConvertToMap("Sheet1")
		if err != nil {
			t.Errorf("Failed to convert to map: %v", err)
		}
		if len(result) != 2 {
			t.Errorf("Expected 2 rows in map, got %d", len(result))
		}
	})
}

func TestStyleOperations(t *testing.T) {
	f := NewExcelFile()
	defer f.Close()

	t.Run("Style Operations", func(t *testing.T) {
		style := &Style{
			Font: &Font{
				Bold: true,
			},
		}

		styleID, err := f.CreateStyle(style)
		if err != nil {
			t.Errorf("Failed to create style: %v", err)
		}

		err = f.SetCellStyle("Sheet1", 1, 1, styleID)
		if err != nil {
			t.Errorf("Failed to set cell style: %v", err)
		}
	})
}
