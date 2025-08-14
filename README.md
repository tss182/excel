# Excel Package for Go

A simple and powerful Go package for reading and creating Excel files, built on top of the excellent [excelize](https://github.com/xuri/excelize) library. This package provides a clean, easy-to-use interface for common Excel operations.

## Features

- ✅ Create new Excel files
- ✅ Read existing Excel files
- ✅ Multiple worksheet support
- ✅ Cell-level operations (read/write)
- ✅ Row-level operations (read/write)
- ✅ Bulk data operations
- ✅ Header management
- ✅ Auto-filtering
- ✅ Column width adjustment
- ✅ Cell styling and formatting
- ✅ Convert data to maps for easy access
- ✅ Simple table creation
- ✅ Comprehensive error handling

## Installation

```bash
go mod init your-project-name
go get github.com/tss182/excel
```

Then copy the `excel` package to your project.

## Quick Start

### Creating a New Excel File

```go
package main

import (
    "fmt"
    "github.com/tss182/excel" // Replace with your actual module path
)

func main() {
    // Create new Excel file
    ef := excel.New()
    defer ef.Close()
    
    // Set some data
    ef.SetCellValue("Sheet1", 1, 1, "Hello")
    ef.SetCellValue("Sheet1", 1, 2, "World")
    
    // Save the file
    err := ef.SaveAs("hello.xlsx")
    if err != nil {
        fmt.Printf("Error: %v\n", err)
    }
}
```

### Reading an Excel File

```go
package main

import (
    "fmt"
    "github.com/tss182/excel"
)

func main() {
    // Open existing file
    ef, err := excel.Open("data.xlsx")
    if err != nil {
        fmt.Printf("Error: %v\n", err)
        return
    }
    defer ef.Close()
    
    // Read cell value
    value, err := ef.GetCellValue("Sheet1", 1, 1)
    if err != nil {
        fmt.Printf("Error: %v\n", err)
        return
    }
    
    fmt.Printf("Cell A1: %s\n", value)
    
    // Read all data
    data, err := ef.ReadData("Sheet1")
    if err != nil {
        fmt.Printf("Error: %v\n", err)
        return
    }
    
    for i, row := range data {
        fmt.Printf("Row %d: %v\n", i+1, row)
    }
}
```

### Creating a Simple Table

```go
package main

import (
    "fmt"
    "github.com/tss182/excel"
)

func main() {
    ef := excel.New()
    defer ef.Close()
    
    // Define headers and data
    headers := []string{"Name", "Age", "City", "Salary"}
    data := [][]interface{}{
        {"John Doe", 30, "New York", 75000},
        {"Jane Smith", 25, "Los Angeles", 65000},
        {"Bob Johnson", 35, "Chicago", 80000},
        {"Alice Brown", 28, "Boston", 70000},
    }
    
    // Create table with auto-sizing
    err := ef.CreateSimpleTable("Employees", headers, data)
    if err != nil {
        fmt.Printf("Error: %v\n", err)
        return
    }
    
    // Add auto-filter
    err = ef.AutoFilter("Employees", 1, 1, len(data)+1, len(headers))
    if err != nil {
        fmt.Printf("Error adding filter: %v\n", err)
        return
    }
    
    // Save file
    err = ef.SaveAs("employees.xlsx")
    if err != nil {
        fmt.Printf("Error: %v\n", err)
    }
}
```

### Converting Data to Maps

```go
package main

import (
    "fmt"
    "github.com/tss182/excel"
)

func main() {
    ef, err := excel.Open("employees.xlsx")
    if err != nil {
        fmt.Printf("Error: %v\n", err)
        return
    }
    defer ef.Close()
    
    // Convert to slice of maps for easy access
    employees, err := ef.ConvertToMap("Employees")
    if err != nil {
        fmt.Printf("Error: %v\n", err)
        return
    }
    
    // Print employee information
    for i, emp := range employees {
        fmt.Printf("Employee %d:\n", i+1)
        fmt.Printf("  Name: %s\n", emp["Name"])
        fmt.Printf("  Age: %s\n", emp["Age"])
        fmt.Printf("  City: %s\n", emp["City"])
        fmt.Printf("  Salary: %s\n", emp["Salary"])
        fmt.Println()
    }
}
```

## API Reference

### File Operations

| Method             | Description |
|--------------------|-------------|
| `New()`            | Create a new Excel file |
| `Open(filename)`   | Open an existing Excel file |
| `SaveAs(filename)` | Save file with specified name |
| `Save()`           | Save file (for existing files) |
| `Close()`          | Close the file |

### Sheet Operations

| Method | Description |
|--------|-------------|
| `CreateSheet(name)` | Create a new worksheet |
| `DeleteSheet(name)` | Delete a worksheet |
| `GetSheetList()` | Get all sheet names |
| `SetActiveSheet(name)` | Set active sheet |

### Cell Operations

| Method | Description |
|--------|-------------|
| `SetCellValue(sheet, row, col, value)` | Set single cell value |
| `GetCellValue(sheet, row, col)` | Get single cell value |
| `SetCellStyle(sheet, row, col, styleID)` | Apply style to cell |

### Row Operations

| Method | Description |
|--------|-------------|
| `SetRowValues(sheet, row, values)` | Set entire row values |
| `GetRowValues(sheet, row)` | Get entire row values |
| `SetHeaders(sheet, headers)` | Set header row (row 1) |
| `GetHeaders(sheet)` | Get header row |

### Bulk Operations

| Method | Description |
|--------|-------------|
| `WriteData(sheet, startRow, data)` | Write multiple rows |
| `ReadData(sheet)` | Read all data from sheet |
| `ReadDataWithHeaders(sheet)` | Read data with headers separated |
| `ConvertToMap(sheet)` | Convert sheet to slice of maps |

### Utility Operations

| Method | Description |
|--------|-------------|
| `CreateSimpleTable(sheet, headers, data)` | Create formatted table |
| `AutoFilter(sheet, startRow, startCol, endRow, endCol)` | Add auto-filter |
| `SetColumnWidth(sheet, col, width)` | Set column width |
| `CreateStyle(style)` | Create new style, returns style ID |

## Advanced Usage

### Cell Formatting

```go
package main

import (
    "github.com/tss182/excel"
)

func main() {
    ef := excel.New()
    defer ef.Close()
    
    // Create a style for headers
    headerStyle := &excel.Style{
        Font: &excel.Font{
            Bold: true,
            Size: 12,
        },
        Fill: excel.Fill{
            Type:    "pattern",
            Color:   []string{"#4472C4"},
            Pattern: 1,
        },
        Alignment: &excel.Alignment{
            Horizontal: "center",
            Vertical:   "center",
        },
    }
    
    // Create the style
    styleID, err := ef.CreateStyle(headerStyle)
    if err != nil {
        // handle error
        return
    }
    
    // Set header values
    headers := []string{"Product", "Price", "Stock"}
    ef.SetHeaders("Sheet1", headers)
    
    // Apply style to header row
    for i := range headers {
        ef.SetCellStyle("Sheet1", 1, i+1, styleID)
    }
    
    ef.SaveAs("styled.xlsx")
}
```

### Working with Multiple Sheets

```go
package main

import (
    "github.com/tss182/excel"
)

func main() {
    ef := excel.New()
    defer ef.Close()
    
    // Create multiple sheets
    sheets := []string{"Sales", "Inventory", "Customers"}
    for _, sheet := range sheets {
        ef.CreateSheet(sheet)
    }
    
    // Add data to different sheets
    salesData := [][]interface{}{
        {"Product A", 1000, 50},
        {"Product B", 1500, 30},
    }
    ef.WriteData("Sales", 1, salesData)
    
    inventoryData := [][]interface{}{
        {"Product A", 100},
        {"Product B", 75},
    }
    ef.WriteData("Inventory", 1, inventoryData)
    
    // List all sheets
    sheetList := ef.GetSheetList()
    fmt.Printf("Sheets: %v\n", sheetList)
    
    ef.SaveAs("multi_sheet.xlsx")
}
```

## Error Handling

The package follows Go's idiomatic error handling patterns. Always check for errors:

```go
ef, err := excel.Open("data.xlsx")
if err != nil {
    log.Fatalf("Failed to open file: %v", err)
}
defer ef.Close()

err = ef.SetCellValue("Sheet1", 1, 1, "Hello")
if err != nil {
    log.Printf("Failed to set cell value: %v", err)
}
```

## Requirements

- Go 1.16 or later
- github.com/xuri/excelize/v2

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Built on top of the excellent [excelize](https://github.com/xuri/excelize) library
- Inspired by the need for a simpler Excel manipulation API in Go

## Support

If you encounter any issues or have questions, please open an issue on GitHub.
