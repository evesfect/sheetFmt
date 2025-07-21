package excel

import (
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

type Editor struct {
	file     *excelize.File
	filepath string
}

// OpenFile opens an existing Excel file
func OpenFile(filepath string) (*Editor, error) {
	file, err := excelize.OpenFile(filepath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %v", err)
	}
	return &Editor{
		file:     file,
		filepath: filepath,
	}, nil
}

// CreateNewFile creates a new Excel file in memory
func CreateNewFile() *Editor {
	file := excelize.NewFile()
	return &Editor{
		file:     file,
		filepath: "",
	}
}

// OpenOrCreateFile opens an existing file or creates a new one if it doesn't exist
func OpenOrCreateFile(filepath string) (*Editor, error) {
	// Check if file exists
	if _, err := os.Stat(filepath); os.IsNotExist(err) {
		// File doesn't exist, create new one
		file := excelize.NewFile()
		return &Editor{
			file:     file,
			filepath: filepath,
		}, nil
	} else if err != nil {
		// Error checking file status
		return nil, fmt.Errorf("error checking file status: %v", err)
	}
	// File exists, open it
	file, err := excelize.OpenFile(filepath)
	if err != nil {
		return nil, fmt.Errorf("failed to open existing file: %v", err)
	}
	return &Editor{
		file:     file,
		filepath: filepath,
	}, nil
}

// ReadColumnValues reads all values from a specific column
func (e *Editor) ReadColumnValues(sheet, column string) ([]string, error) {
	rows, err := e.file.GetRows(sheet)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows: %v", err)
	}
	var columnValues []string
	colIndex := columnToIndex(column)
	for _, row := range rows {
		if colIndex < len(row) {
			columnValues = append(columnValues, row[colIndex])
		} else {
			columnValues = append(columnValues, "")
		}
	}
	return columnValues, nil
}

// SetColumnHeader changes the header (first row) of a column
func (e *Editor) SetColumnHeader(sheet, column, headerName string) error {
	cell := column + "1"
	return e.file.SetCellValue(sheet, cell, headerName)
}

// SetCellValue sets a value in a specific cell
func (e *Editor) SetCellValue(sheet, cell string, value interface{}) error {
	return e.file.SetCellValue(sheet, cell, value)
}

// SetCellFormula sets a formula for a specific cell
func (e *Editor) SetCellFormula(sheet, cell, formula string) error {
	return e.file.SetCellFormula(sheet, cell, formula)
}

// GetColumnHeaders returns all column headers (first row)
func (e *Editor) GetColumnHeaders(sheet string) ([]string, error) {
	firstRow, err := e.file.GetRows(sheet)
	if err != nil {
		return nil, fmt.Errorf("failed to get first row: %v", err)
	}
	if len(firstRow) == 0 {
		return []string{}, nil
	}
	return firstRow[0], nil
}

// GetCellFormula returns the formula in a specific cell (if any)
func (e *Editor) GetCellFormula(sheet, cell string) (string, error) {
	return e.file.GetCellFormula(sheet, cell)
}

// GetCellValue returns the value in a specific cell
func (e *Editor) GetCellValue(sheet, cell string) (string, error) {
	return e.file.GetCellValue(sheet, cell)
}

func (e *Editor) GetCellDataType(sheet, cell string) (excelize.CellType, error) {
	return e.file.GetCellType(sheet, cell)
}

// PrintCellInfo prints detailed info about a cell (for debugging)
func (e *Editor) PrintCellInfo(sheet, cell string) {
	value, err := e.file.GetCellValue(sheet, cell)
	if err != nil {
		fmt.Printf("Error getting value for %s: %v\n", cell, err)
		return
	}

	formula, err := e.file.GetCellFormula(sheet, cell)
	if err != nil {
		fmt.Printf("Error getting formula for %s: %v\n", cell, err)
		return
	}

	if formula != "" {
		fmt.Printf("Cell %s: Formula='%s', Value='%s'\n", cell, formula, value)
	} else {
		fmt.Printf("Cell %s: Value='%s' (no formula)\n", cell, value)
	}
}

// GetSheetNames returns all sheet names in the workbook
func (e *Editor) GetSheetNames() []string {
	return e.file.GetSheetList()
}

// AddSheet creates a new sheet
func (e *Editor) AddSheet(sheetName string) error {
	_, err := e.file.NewSheet(sheetName)
	return err
}

// DeleteSheet removes a sheet
func (e *Editor) DeleteSheet(sheetName string) error {
	return e.file.DeleteSheet(sheetName)
}

// Save saves the Excel file to the original filepath
func (e *Editor) Save() error {
	if e.filepath == "" {
		return fmt.Errorf("no filepath specified, use SaveAs instead")
	}
	return e.file.SaveAs(e.filepath)
}

// SaveAs saves the Excel file with a new name
func (e *Editor) SaveAs(filepath string) error {
	e.filepath = filepath
	return e.file.SaveAs(filepath)
}

// Close closes the Excel file
func (e *Editor) Close() error {
	return e.file.Close()
}

// GetAllRows returns all rows from a sheet
func (e *Editor) GetAllRows(sheet string) ([][]string, error) {
	return e.file.GetRows(sheet)
}

// Helper function to convert column letter to index
func columnToIndex(column string) int {
	result := 0
	for _, char := range column {
		result = result*26 + int(char-'A'+1)
	}
	return result - 1 // Convert to 0-based index
}

// Helper function to convert column index to Excel column letter
/*func indexToColumn(index int) string {
	result := ""
	for index >= 0 {
		result = string(rune('A'+index%26)) + result
		index = index/26 - 1
	}
	return result
}*/

// InsertRows inserts the specified number of rows starting at the given row number
func (e *Editor) InsertRows(sheet string, startRow, numRows int) error {
	for i := 0; i < numRows; i++ {
		err := e.file.InsertRows(sheet, startRow, 1)
		if err != nil {
			return fmt.Errorf("failed to insert row at position %d: %v", startRow, err)
		}
	}
	return nil
}

// parseNumericValue attempts to parse a string as a number and returns the appropriate type
// Returns the original string if it's not a valid number, and a flag indicating if it's a float
func parseNumericValue(value string) (interface{}, bool) {
	// Trim whitespace
	trimmed := strings.TrimSpace(value)
	if trimmed == "" {
		return value, false
	}

	// Try to parse as integer first
	if intVal, err := strconv.ParseInt(trimmed, 10, 64); err == nil {
		return intVal, false
	}

	// Try to parse as float
	if floatVal, err := strconv.ParseFloat(trimmed, 64); err == nil {
		return floatVal, true
	}

	// Not a number, return as string
	return value, false
}

// SetCellValueSmart sets a cell value, automatically detecting if it's a number
// For float values, applies formatting to show 2 decimal places
func (e *Editor) SetCellValueSmart(sheet, cell string, value string) error {
	numericValue, isFloat := parseNumericValue(value)

	// Set the cell value
	err := e.file.SetCellValue(sheet, cell, numericValue)
	if err != nil {
		return err
	}

	// If it's a float, apply 2 decimal places formatting
	if isFloat {
		err = e.applyFloatFormatting(sheet, cell)
		if err != nil {
			return fmt.Errorf("failed to apply float formatting to cell %s: %v", cell, err)
		}
	}

	return nil
}

// applyFloatFormatting applies number formatting with 2 decimal places to a cell
func (e *Editor) applyFloatFormatting(sheet, cell string) error {
	// Create a style with 2 decimal places
	style, err := e.file.NewStyle(&excelize.Style{
		NumFmt: 2, // Built-in format for 2 decimal places (0.00)
	})
	if err != nil {
		return fmt.Errorf("failed to create float style: %v", err)
	}

	// Apply the style to the cell
	err = e.file.SetCellStyle(sheet, cell, cell, style)
	if err != nil {
		return fmt.Errorf("failed to apply float style: %v", err)
	}

	return nil
}
