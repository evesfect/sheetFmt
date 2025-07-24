package excel

import (
	"fmt"
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

// GetSheetNames returns all sheet names in the workbook
func (e *Editor) GetSheetNames() []string {
	return e.file.GetSheetList()
}

// Close closes the Excel file
func (e *Editor) Close() error {
	return e.file.Close()
}

// GetAllRows returns all rows from a sheet
func (e *Editor) GetAllRows(sheet string) ([][]string, error) {
	return e.file.GetRows(sheet)
}
