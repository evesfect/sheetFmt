package excel

import (
	"fmt"
	"os"
	"regexp"
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

// DetectHeaderRow finds the row containing column headers using the strategy:
// Find rightmost column with data, then find first row with data in that column
func (e *Editor) DetectHeaderRow(sheet string) (int, error) {
	rows, err := e.file.GetRows(sheet)
	if err != nil {
		return 0, fmt.Errorf("failed to get rows: %v", err)
	}

	if len(rows) == 0 {
		return 1, nil // Default to row 1 if no data
	}

	// Find the rightmost column with data across all rows
	rightmostCol := 0
	for _, row := range rows {
		for colIdx := len(row) - 1; colIdx >= 0; colIdx-- {
			if strings.TrimSpace(row[colIdx]) != "" {
				if colIdx+1 > rightmostCol {
					rightmostCol = colIdx + 1 // Convert to 1-based column
				}
				break // Found the rightmost data in this row
			}
		}
	}

	if rightmostCol == 0 {
		return 1, nil // Default to row 1 if no data
	}

	// Now find the first row that has data in the rightmost column
	for rowIdx, row := range rows {
		if len(row) >= rightmostCol && strings.TrimSpace(row[rightmostCol-1]) != "" {
			return rowIdx + 1, nil // Convert to 1-based row number
		}
	}

	return 1, nil // Default to row 1 if not found
}

// GetColumnHeaders returns all column headers from the detected header row
func (e *Editor) GetColumnHeaders(sheet string) ([]string, error) {
	headerRow, err := e.DetectHeaderRow(sheet)
	if err != nil {
		return nil, fmt.Errorf("failed to detect header row: %v", err)
	}

	rows, err := e.file.GetRows(sheet)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows: %v", err)
	}

	if len(rows) < headerRow {
		return []string{}, nil
	}

	return rows[headerRow-1], nil // Convert back to 0-based index
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

// ConvertCandidateToTargetFormat converts a candidate target format file to proper target format
// by converting example formulas to column-wide templates and clearing example data
func ConvertCandidateToTargetFormat(candidateFile, outputFile string, formulaRow int) error {
	// Open the candidate file
	editor, err := OpenFile(candidateFile)
	if err != nil {
		return fmt.Errorf("failed to open candidate file: %v", err)
	}
	defer editor.Close()

	// Get all sheet names and process the first sheet (or we could make this configurable)
	sheetNames := editor.GetSheetNames()
	if len(sheetNames) == 0 {
		return fmt.Errorf("no sheets found in candidate file")
	}

	sheetName := sheetNames[0] // Process first sheet
	fmt.Printf("Processing sheet: %s\n", sheetName)

	// Detect header row
	headerRow, err := editor.DetectHeaderRow(sheetName)
	if err != nil {
		return fmt.Errorf("failed to detect header row: %v", err)
	}
	fmt.Printf("Detected header row: %d\n", headerRow)

	// The example row is the first row after headers
	exampleRow := headerRow + 1
	fmt.Printf("Example row: %d\n", exampleRow)

	// Get worksheet to find max column
	rows, err := editor.file.GetRows(sheetName)
	if err != nil {
		return fmt.Errorf("failed to get rows: %v", err)
	}

	if len(rows) < exampleRow {
		return fmt.Errorf("example row %d does not exist", exampleRow)
	}

	// Use the header row to determine max columns
	maxColumn := len(rows[headerRow-1]) // Convert to 0-based index for rows slice
	fmt.Printf("Max columns: %d\n", maxColumn)

	formulasFound := 0
	cellsCleared := 0

	// Process each column in the example row
	for colIdx := 0; colIdx < maxColumn; colIdx++ {
		colLetter := indexToColumn(colIdx)
		cellRef := fmt.Sprintf("%s%d", colLetter, exampleRow)

		fmt.Printf("DEBUG: Checking cell %s\n", cellRef)

		// Check cell value first
		cellValue, err := editor.GetCellValue(sheetName, cellRef)
		if err != nil {
			fmt.Printf("DEBUG: Error reading cell value %s: %v\n", cellRef, err)
			continue
		}
		fmt.Printf("DEBUG: Cell %s value: '%s'\n", cellRef, cellValue)

		// Check if this cell has a formula
		formula, err := editor.GetCellFormula(sheetName, cellRef)
		if err != nil {
			fmt.Printf("DEBUG: Error reading formula %s: %v\n", cellRef, err)
			continue
		}

		fmt.Printf("DEBUG: Cell %s formula: '%s'\n", cellRef, formula)

		// If cell has a formula OR starts with =, process it
		hasFormula := false
		var templateFormula string

		if formula != "" {
			fmt.Printf("Found formula in %s: %s\n", cellRef, formula)
			templateFormula = convertFormulaToTemplate(formula)
			hasFormula = true
		} else if strings.HasPrefix(strings.TrimSpace(cellValue), "=") {
			fmt.Printf("Found text formula in %s: %s\n", cellRef, cellValue)
			templateFormula = convertFormulaToTemplate(strings.TrimSpace(cellValue))
			hasFormula = true
		}

		if hasFormula {
			fmt.Printf("Converted to template: %s\n", templateFormula)

			// Place template formula at the configured formula row
			templateCellRef := fmt.Sprintf("%s%d", colLetter, formulaRow)

			// Set as text (since we want it to be read as a template later)
			err = editor.SetCellValue(sheetName, templateCellRef, templateFormula)
			if err != nil {
				fmt.Printf("Warning: Failed to set template formula at %s: %v\n", templateCellRef, err)
			} else {
				fmt.Printf("✓ Set template formula at %s: %s\n", templateCellRef, templateFormula)
				formulasFound++
			}
		}

		// Clear the cell in example row if it has any content
		if cellValue != "" || formula != "" {
			err = editor.SetCellValue(sheetName, cellRef, "")
			if err != nil {
				fmt.Printf("Warning: Failed to clear cell %s: %v\n", cellRef, err)
			} else {
				fmt.Printf("✓ Cleared cell %s\n", cellRef)
				cellsCleared++
			}
		}
	}

	fmt.Printf("Summary: Found %d formulas, cleared %d cells\n", formulasFound, cellsCleared)

	// Save as the target format file
	err = editor.SaveAs(outputFile)
	if err != nil {
		return fmt.Errorf("failed to save target format file: %v", err)
	}

	fmt.Printf("✓ Successfully converted candidate to target format: %s\n", outputFile)
	return nil
}

// convertFormulaToTemplate converts a cell-specific formula to a column-wide template
// Example: "=+I4*((100-J4)/100)" becomes "=+I*((100-J)/100)"
func convertFormulaToTemplate(formula string) string {
	// Use regex to find cell references (like I4, J4, AB4) and remove row numbers
	// Pattern matches: one or more uppercase letters followed by one or more digits
	re := regexp.MustCompile(`([A-Z]+)\d+`)

	// Replace with just the column letters
	template := re.ReplaceAllString(formula, "$1")

	return template
}

// Helper function to convert column index to Excel column letter
func indexToColumn(index int) string {
	result := ""
	for index >= 0 {
		result = string(rune('A'+index%26)) + result
		index = index/26 - 1
	}
	return result
}
