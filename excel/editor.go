package excel

import (
	"fmt"
	"os"

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
		fmt.Printf("File %s doesn't exist, creating new file...\n", filepath)
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
	fmt.Printf("File %s exists, opening...\n", filepath)
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
func indexToColumn(index int) string {
	result := ""
	for index >= 0 {
		result = string(rune('A'+index%26)) + result
		index = index/26 - 1
	}
	return result
}

// ApplyTargetFormat applies formatting from a target file to the current file
func (e *Editor) ApplyTargetFormat(targetFilePath string, targetSheet, currentSheet string) error {
	// Open the target format file
	targetEditor, err := OpenFile(targetFilePath)
	if err != nil {
		return fmt.Errorf("failed to open target format file: %v", err)
	}
	defer targetEditor.Close()

	// Get all rows from target sheet to determine the range
	targetRows, err := targetEditor.GetAllRows(targetSheet)
	if err != nil {
		return fmt.Errorf("failed to read target sheet: %v", err)
	}

	// Track cells we've processed
	processedCount := 0
	skippedCount := 0

	// Process each potential cell in the target range
	for rowIndex := 0; rowIndex < len(targetRows); rowIndex++ {
		// Get the maximum column count for this row and previous rows
		maxCols := 0
		for i := 0; i <= rowIndex && i < len(targetRows); i++ {
			if len(targetRows[i]) > maxCols {
				maxCols = len(targetRows[i])
			}
		}

		for colIndex := 0; colIndex < maxCols; colIndex++ {
			// Convert column index to Excel column letter (A, B, C, etc.)
			colLetter := indexToColumn(colIndex)
			cellAddress := fmt.Sprintf("%s%d", colLetter, rowIndex+1)

			// Check if this cell has a formula first
			formula, err := targetEditor.file.GetCellFormula(targetSheet, cellAddress)
			if err != nil {
				continue // Skip if we can't read the formula
			}

			if formula != "" {
				// It's a formula, copy the formula
				err = e.SetCellFormula(currentSheet, cellAddress, formula)
				if err != nil {
					return fmt.Errorf("failed to set formula in cell %s: %v", cellAddress, err)
				}
				fmt.Printf("Applied formula to %s: %s\n", cellAddress, formula)
				processedCount++
			} else {
				// Check if the cell has a non-empty value
				cellValue, err := targetEditor.file.GetCellValue(targetSheet, cellAddress)
				if err != nil {
					continue // Skip if we can't read the value
				}

				// Only apply if the target cell has actual content
				if cellValue != "" {
					err = e.SetCellValue(currentSheet, cellAddress, cellValue)
					if err != nil {
						return fmt.Errorf("failed to set value in cell %s: %v", cellAddress, err)
					}
					fmt.Printf("Applied value to %s: %s\n", cellAddress, cellValue)
					processedCount++
				} else {
					// Cell is empty in target - leave the edited file unchanged
					skippedCount++
				}
			}
		}
	}

	fmt.Printf("Target format applied: %d cells processed, %d empty cells skipped\n", processedCount, skippedCount)
	return nil
}

// ApplyTargetFormatToFile applies target format from one file to another file
func ApplyTargetFormatToFile(targetFilePath, editedFilePath string, targetSheet, editedSheet string) error {
	// Open the file to be edited
	editor, err := OpenOrCreateFile(editedFilePath)
	if err != nil {
		return fmt.Errorf("failed to open edited file: %v", err)
	}
	defer editor.Close()

	// Apply the target format
	err = editor.ApplyTargetFormat(targetFilePath, targetSheet, editedSheet)
	if err != nil {
		return err
	}

	// Save the changes
	err = editor.Save()
	if err != nil {
		return fmt.Errorf("failed to save edited file: %v", err)
	}

	fmt.Printf("Target format applied successfully from %s to %s\n", targetFilePath, editedFilePath)
	return nil
}
