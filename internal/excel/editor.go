package excel

import (
	"fmt"
	"os"
	"path/filepath"
	"sheetFmt/internal/mapping"
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

// FormatFileWithTarget formats a single input file using target format and mappings
func FormatFileWithTarget(inputFilePath, targetFilePath, mappingFilePath, outputFilePath string, inputSheet, targetSheet string) error {
	// Load mapping configuration
	mappingConfig, err := mapping.LoadFromFile(mappingFilePath)
	if err != nil {
		return fmt.Errorf("failed to load mapping configuration: %v", err)
	}

	// Create reverse mapping: target column -> scanned column
	targetToScanned := make(map[string]string)
	for _, m := range mappingConfig.Mappings {
		if !m.IsIgnored && m.TargetColumn != "" {
			targetToScanned[m.TargetColumn] = m.ScannedColumn
		}
	}

	// Open target format file
	targetEditor, err := OpenFile(targetFilePath)
	if err != nil {
		return fmt.Errorf("failed to open target format file: %v", err)
	}
	defer targetEditor.Close()

	// Open input file
	inputEditor, err := OpenOrCreateFile(inputFilePath)
	if err != nil {
		return fmt.Errorf("failed to open input file %s: %v", inputFilePath, err)
	}
	defer inputEditor.Close()

	// Get target format headers
	targetHeaders, err := targetEditor.GetColumnHeaders(targetSheet)
	if err != nil {
		return fmt.Errorf("failed to get target headers: %v", err)
	}

	// Get input file headers
	inputHeaders, err := inputEditor.GetColumnHeaders(inputSheet)
	if err != nil {
		return fmt.Errorf("failed to get input headers: %v", err)
	}

	// Create header name to column index mapping for input
	inputHeaderMap := make(map[string]int)
	for i, header := range inputHeaders {
		inputHeaderMap[header] = i
	}

	// Get all input rows for data copying
	inputRows, err := inputEditor.GetAllRows(inputSheet)
	if err != nil {
		return fmt.Errorf("failed to get input rows: %v", err)
	}

	var errorMessages []string
	hasErrors := false
	inputFileName := filepath.Base(inputFilePath)

	// Process each target column
	for targetColIndex, targetHeader := range targetHeaders {
		targetColLetter := indexToColumn(targetColIndex)

		// Set the target header
		err = inputEditor.SetColumnHeader(inputSheet, targetColLetter, targetHeader)
		if err != nil {
			return fmt.Errorf("failed to set target header for column %s: %v", targetColLetter, err)
		}

		// Check if this target column has a mapping
		if scannedColumn, hasMapping := targetToScanned[targetHeader]; hasMapping {
			// Check if the scanned column exists in input file
			if inputColIndex, exists := inputHeaderMap[scannedColumn]; exists {
				// Copy data from input column to target column
				for rowIndex := 1; rowIndex < len(inputRows); rowIndex++ { // Skip header row
					cellAddress := fmt.Sprintf("%s%d", targetColLetter, rowIndex+1)

					// Check if the target format has a formula in this cell
					targetFormula, err := targetEditor.GetCellFormula(targetSheet, cellAddress)
					if err != nil {
						return fmt.Errorf("failed to check formula in target cell %s: %v", cellAddress, err)
					}

					if targetFormula != "" {
						// Target has a formula, copy the formula instead of data
						err = inputEditor.SetCellFormula(inputSheet, cellAddress, targetFormula)
						if err != nil {
							return fmt.Errorf("failed to set formula %s: %v", cellAddress, err)
						}
					} else {
						// Target has no formula, copy the data value
						if inputColIndex < len(inputRows[rowIndex]) {
							err = inputEditor.SetCellValue(inputSheet, cellAddress, inputRows[rowIndex][inputColIndex])
							if err != nil {
								return fmt.Errorf("failed to set cell value %s: %v", cellAddress, err)
							}
						}
					}
				}
			} else {
				// Mapping exists but column not found in input
				errorMessages = append(errorMessages, fmt.Sprintf("%s:%s:: mapped column '%s' not found in input", inputFileName, inputSheet, scannedColumn))
				hasErrors = true
			}
		} else {
			// No mapping for this target column - still need to copy formulas if they exist
			for rowIndex := 1; rowIndex < len(inputRows); rowIndex++ { // Skip header row
				cellAddress := fmt.Sprintf("%s%d", targetColLetter, rowIndex+1)

				// Check if the target format has a formula in this cell
				targetFormula, err := targetEditor.GetCellFormula(targetSheet, cellAddress)
				if err != nil {
					return fmt.Errorf("failed to check formula in target cell %s: %v", cellAddress, err)
				}

				if targetFormula != "" {
					// Copy formula from target format even if no mapping
					err = inputEditor.SetCellFormula(inputSheet, cellAddress, targetFormula)
					if err != nil {
						return fmt.Errorf("failed to set formula %s: %v", cellAddress, err)
					}
				}
			}

			errorMessages = append(errorMessages, fmt.Sprintf("%s:%s:: no mapping for '%s'", inputFileName, inputSheet, targetHeader))
			hasErrors = true
		}
	}

	// If there are any errors, print them and handle problematic file
	if hasErrors {
		for _, msg := range errorMessages {
			fmt.Println(msg)
		}

		// Copy file to problematic directory
		problematicDir := "data/problematic"
		err2 := os.MkdirAll(problematicDir, 0755)
		if err2 != nil {
			fmt.Printf("Failed to create problematic directory: %v\n", err2)
		} else {
			problematicPath := filepath.Join(problematicDir, filepath.Base(inputFilePath))
			err2 = copyFile(inputFilePath, problematicPath)
			if err2 != nil {
				fmt.Printf("Failed to copy problematic file: %v\n", err2)
			}
		}

		return fmt.Errorf("formatting failed")
	}

	// Save the formatted file
	err = inputEditor.SaveAs(outputFilePath)
	if err != nil {
		return fmt.Errorf("failed to save formatted file %s: %v", outputFilePath, err)
	}

	return nil
}

// FormatFile formats an entire Excel file with all its sheets
func FormatFile(inputFilePath, targetFilePath, mappingFilePath, targetSheet string) error {
	// Check if input file exists
	if _, err := os.Stat(inputFilePath); os.IsNotExist(err) {
		return fmt.Errorf("input file not found: %s", inputFilePath)
	}

	// Check if mapping file exists
	if _, err := os.Stat(mappingFilePath); os.IsNotExist(err) {
		return fmt.Errorf("mapping file not found: %s", mappingFilePath)
	}

	// Check if target format file exists
	if _, err := os.Stat(targetFilePath); os.IsNotExist(err) {
		return fmt.Errorf("target format file not found: %s", targetFilePath)
	}

	// Create results directory
	resultsDir := "data/results"
	err := os.MkdirAll(resultsDir, 0755)
	if err != nil {
		return fmt.Errorf("failed to create results directory: %v", err)
	}

	// Open input file to get sheet names
	inputEditor, err := OpenFile(inputFilePath)
	if err != nil {
		return fmt.Errorf("failed to open input file: %v", err)
	}
	defer inputEditor.Close()

	// Get all sheet names from input file
	inputSheets := inputEditor.GetSheetNames()
	if len(inputSheets) == 0 {
		return fmt.Errorf("no sheets found in input file")
	}

	// Get base filename without extension for output naming
	inputFileName := filepath.Base(inputFilePath)
	inputFileExt := filepath.Ext(inputFileName)
	inputFileBase := strings.TrimSuffix(inputFileName, inputFileExt)

	// Process each sheet separately
	for _, sheetName := range inputSheets {
		// Generate output file name: <inputfilename>-<sheetname>.xlsx
		outputFileName := fmt.Sprintf("%s-%s.xlsx", inputFileBase, sheetName)
		outputFilePath := filepath.Join(resultsDir, outputFileName)

		// Format the sheet
		err = FormatFileWithTarget(
			inputFilePath,
			targetFilePath,
			mappingFilePath,
			outputFilePath,
			sheetName,
			targetSheet,
		)

		if err != nil {
			fmt.Printf("Problematic file copied to: data/problematic/%s\n\n", filepath.Base(inputFilePath))
		} else {
			fmt.Printf("Format successful for %s\n\n", sheetName)
		}
	}

	return nil
}

// Helper function to copy files
func copyFile(src, dst string) error {
	input, err := os.ReadFile(src)
	if err != nil {
		return err
	}
	return os.WriteFile(dst, input, 0644)
}
