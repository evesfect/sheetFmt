package excel

import (
	"fmt"
	"os"
	"path/filepath"
	"sheetFmt/internal/mapping"
	"strings"
)

// FormatConfig holds configuration for formatting operations
type FormatConfig struct {
	InputFilePath   string
	TargetFilePath  string
	MappingFilePath string
	OutputFilePath  string
	InputSheet      string
	TargetSheet     string
}

// FormatResult contains the result of a formatting operation
type FormatResult struct {
	Success       bool
	ErrorMessages []string
	InputFileName string
}

// FormatFileWithTarget formats a single input file using target format and mappings
func FormatFileWithTarget(inputFilePath, targetFilePath, mappingFilePath, outputFilePath string, inputSheet, targetSheet string) error {
	config := &FormatConfig{
		InputFilePath:   inputFilePath,
		TargetFilePath:  targetFilePath,
		MappingFilePath: mappingFilePath,
		OutputFilePath:  outputFilePath,
		InputSheet:      inputSheet,
		TargetSheet:     targetSheet,
	}

	formatter := &fileFormatter{config: config}
	return formatter.format()
}

// FormatFile formats an entire Excel file with all its sheets
func FormatFile(inputFilePath, targetFilePath, mappingFilePath, targetSheet string) error {
	if err := validateInputFiles(inputFilePath, targetFilePath, mappingFilePath); err != nil {
		return err
	}

	if err := createResultsDirectory(); err != nil {
		return err
	}

	inputSheets, err := getInputSheets(inputFilePath)
	if err != nil {
		return err
	}

	return processAllSheets(inputFilePath, targetFilePath, mappingFilePath, targetSheet, inputSheets)
}

// fileFormatter handles the formatting of a single file
type fileFormatter struct {
	config          *FormatConfig
	mappingConfig   *mapping.MappingConfig
	targetToScanned map[string]string
	targetEditor    *Editor
	inputEditor     *Editor
	targetHeaders   []string
	inputHeaders    []string
	inputHeaderMap  map[string]int
	inputRows       [][]string
	columnFormulas  map[string]string
	result          *FormatResult
}

// format performs the complete formatting operation
func (f *fileFormatter) format() error {
	if err := f.initialize(); err != nil {
		return err
	}
	defer f.cleanup()

	if err := f.loadMappingAndHeaders(); err != nil {
		return err
	}

	if err := f.prepareDataAndFormulas(); err != nil {
		return err
	}

	if err := f.processAllColumns(); err != nil {
		return f.handleFormattingError()
	}

	return f.saveFormattedFile()
}

// initialize sets up the formatter and opens required files
func (f *fileFormatter) initialize() error {
	f.result = &FormatResult{
		InputFileName: filepath.Base(f.config.InputFilePath),
		ErrorMessages: []string{},
	}

	var err error
	f.targetEditor, err = OpenFile(f.config.TargetFilePath)
	if err != nil {
		return fmt.Errorf("failed to open target format file: %v", err)
	}

	f.inputEditor, err = OpenOrCreateFile(f.config.InputFilePath)
	if err != nil {
		return fmt.Errorf("failed to open input file %s: %v", f.config.InputFilePath, err)
	}

	return nil
}

// cleanup closes opened files
func (f *fileFormatter) cleanup() {
	if f.targetEditor != nil {
		f.targetEditor.Close()
	}
	if f.inputEditor != nil {
		f.inputEditor.Close()
	}
}

// loadMappingAndHeaders loads mapping configuration and column headers
func (f *fileFormatter) loadMappingAndHeaders() error {
	// Load mapping configuration
	var err error
	f.mappingConfig, err = mapping.LoadFromFile(f.config.MappingFilePath)
	if err != nil {
		return fmt.Errorf("failed to load mapping configuration: %v", err)
	}

	// Create reverse mapping
	f.targetToScanned = make(map[string]string)
	for _, m := range f.mappingConfig.Mappings {
		if !m.IsIgnored && m.TargetColumn != "" {
			f.targetToScanned[m.TargetColumn] = m.ScannedColumn
		}
	}

	// Get headers
	f.targetHeaders, err = f.targetEditor.GetColumnHeaders(f.config.TargetSheet)
	if err != nil {
		return fmt.Errorf("failed to get target headers: %v", err)
	}

	f.inputHeaders, err = f.inputEditor.GetColumnHeaders(f.config.InputSheet)
	if err != nil {
		return fmt.Errorf("failed to get input headers: %v", err)
	}

	// Create header mapping
	f.inputHeaderMap = make(map[string]int)
	for i, header := range f.inputHeaders {
		f.inputHeaderMap[header] = i
	}

	return nil
}

// prepareDataAndFormulas loads input data and processes column formulas
func (f *fileFormatter) prepareDataAndFormulas() error {
	// Get all input rows
	var err error
	f.inputRows, err = f.inputEditor.GetAllRows(f.config.InputSheet)
	if err != nil {
		return fmt.Errorf("failed to get input rows: %v", err)
	}

	// Check for column formulas and handle row insertion if needed
	f.columnFormulas = make(map[string]string)
	hasColumnFormulas := f.detectColumnFormulas()

	if hasColumnFormulas && f.hasExistingData() {
		if err := f.insertRowForFormulas(); err != nil {
			return err
		}
	}

	return nil
}

// detectColumnFormulas finds column formulas in target row 2
func (f *fileFormatter) detectColumnFormulas() bool {
	hasColumnFormulas := false

	for targetColIndex, targetHeader := range f.targetHeaders {
		targetColLetter := indexToColumn(targetColIndex)
		row2Cell := fmt.Sprintf("%s2", targetColLetter)

		formula, err := f.targetEditor.file.GetCellFormula(f.config.TargetSheet, row2Cell)
		if err != nil {
			continue
		}

		if formula != "" {
			f.columnFormulas[targetColLetter] = formula
			hasColumnFormulas = true
			fmt.Printf("Found column formula template in %s (%s): %s\n", row2Cell, targetHeader, formula)
		}
	}

	return hasColumnFormulas
}

// hasExistingData checks if there's existing data beyond the header
func (f *fileFormatter) hasExistingData() bool {
	originalDataRowCount := len(f.inputRows)
	if originalDataRowCount > 0 {
		originalDataRowCount-- // Exclude header row
	}
	return originalDataRowCount > 0
}

// insertRowForFormulas inserts a row to preserve existing data when column formulas are present
func (f *fileFormatter) insertRowForFormulas() error {
	err := f.inputEditor.InsertRows(f.config.InputSheet, 2, 1)
	if err != nil {
		return fmt.Errorf("failed to insert row for column formulas: %v", err)
	}
	fmt.Printf("Inserted row after header to preserve existing data\n")

	// Re-read input rows after insertion
	f.inputRows, err = f.inputEditor.GetAllRows(f.config.InputSheet)
	if err != nil {
		return fmt.Errorf("failed to get updated input rows: %v", err)
	}

	return nil
}

// processAllColumns processes each target column
func (f *fileFormatter) processAllColumns() error {
	for targetColIndex, targetHeader := range f.targetHeaders {
		targetColLetter := indexToColumn(targetColIndex)

		if err := f.processTargetColumn(targetHeader, targetColLetter); err != nil {
			return err
		}
	}

	return nil
}

// processTargetColumn processes a single target column
func (f *fileFormatter) processTargetColumn(targetHeader, targetColLetter string) error {
	// Set the target header
	err := f.inputEditor.SetColumnHeader(f.config.InputSheet, targetColLetter, targetHeader)
	if err != nil {
		return fmt.Errorf("failed to set target header for column %s: %v", targetColLetter, err)
	}

	// Check if this column has a column formula
	if formula, hasColumnFormula := f.columnFormulas[targetColLetter]; hasColumnFormula {
		return f.applyColumnFormula(targetColLetter, formula)
	}

	// Process regular column mapping
	return f.processColumnMapping(targetHeader, targetColLetter)
}

// applyColumnFormula applies column formula to all data rows
func (f *fileFormatter) applyColumnFormula(targetColLetter, formula string) error {
	for rowIndex := 2; rowIndex < len(f.inputRows); rowIndex++ {
		cellAddress := fmt.Sprintf("%s%d", targetColLetter, rowIndex)
		adjustedFormula := adjustFormulaForRow(formula, rowIndex)

		err := f.inputEditor.SetCellFormula(f.config.InputSheet, cellAddress, adjustedFormula)
		if err != nil {
			return fmt.Errorf("failed to set column formula in cell %s: %v", cellAddress, err)
		}
	}
	fmt.Printf("Applied column formula to %s: %s\n", targetColLetter, formula)
	return nil
}

// processColumnMapping handles mapping and data copying for a column
func (f *fileFormatter) processColumnMapping(targetHeader, targetColLetter string) error {
	// Check if this target column has a mapping
	if scannedColumn, hasMapping := f.targetToScanned[targetHeader]; hasMapping {
		return f.processColumnWithMapping(targetColLetter, scannedColumn)
	}

	// No mapping - copy formulas if they exist
	f.copyFormulasOnly(targetColLetter)
	f.addError(fmt.Sprintf("%s:%s:: no mapping for '%s'", f.result.InputFileName, f.config.InputSheet, targetHeader))
	return nil
}

// processColumnWithMapping processes a column that has a mapping
func (f *fileFormatter) processColumnWithMapping(targetColLetter, scannedColumn string) error {
	inputColIndex, exists := f.inputHeaderMap[scannedColumn]
	if !exists {
		f.addError(fmt.Sprintf("%s:%s:: mapped column '%s' not found in input", f.result.InputFileName, f.config.InputSheet, scannedColumn))
		return nil
	}

	// Copy data from input column to target column
	for rowIndex := 1; rowIndex < len(f.inputRows); rowIndex++ {
		cellAddress := fmt.Sprintf("%s%d", targetColLetter, rowIndex+1)

		if err := f.processCellMapping(cellAddress, rowIndex, inputColIndex); err != nil {
			return err
		}
	}

	return nil
}

// processCellMapping handles mapping for a single cell
func (f *fileFormatter) processCellMapping(cellAddress string, rowIndex, inputColIndex int) error {
	// Check if the target format has a formula in this cell
	targetFormula, err := f.targetEditor.GetCellFormula(f.config.TargetSheet, cellAddress)
	if err != nil {
		return fmt.Errorf("failed to check formula in target cell %s: %v", cellAddress, err)
	}

	if targetFormula != "" {
		// Target has a formula, copy the formula instead of data
		return f.inputEditor.SetCellFormula(f.config.InputSheet, cellAddress, targetFormula)
	}

	// Target has no formula, copy the data value with smart type detection
	if inputColIndex < len(f.inputRows[rowIndex]) {
		cellValue := f.inputRows[rowIndex][inputColIndex]
		return f.inputEditor.SetCellValueSmart(f.config.InputSheet, cellAddress, cellValue)
	}

	return nil
}

// copyFormulasOnly copies formulas from target format when no mapping exists
func (f *fileFormatter) copyFormulasOnly(targetColLetter string) {
	for rowIndex := 1; rowIndex < len(f.inputRows); rowIndex++ {
		cellAddress := fmt.Sprintf("%s%d", targetColLetter, rowIndex+1)

		targetFormula, err := f.targetEditor.GetCellFormula(f.config.TargetSheet, cellAddress)
		if err != nil {
			continue
		}

		if targetFormula != "" {
			f.inputEditor.SetCellFormula(f.config.InputSheet, cellAddress, targetFormula)
		}
	}
}

// addError adds an error message to the result
func (f *fileFormatter) addError(message string) {
	f.result.ErrorMessages = append(f.result.ErrorMessages, message)
}

// handleFormattingError handles errors during formatting
func (f *fileFormatter) handleFormattingError() error {
	// Print all error messages
	for _, msg := range f.result.ErrorMessages {
		fmt.Println(msg)
	}

	// Copy file to problematic directory
	f.copyToProblematicDirectory()

	return fmt.Errorf("formatting failed")
}

// copyToProblematicDirectory copies the problematic file to a special directory
func (f *fileFormatter) copyToProblematicDirectory() {
	problematicDir := "data/problematic"
	err := os.MkdirAll(problematicDir, 0755)
	if err != nil {
		fmt.Printf("Failed to create problematic directory: %v\n", err)
		return
	}

	problematicPath := filepath.Join(problematicDir, filepath.Base(f.config.InputFilePath))
	err = copyFile(f.config.InputFilePath, problematicPath)
	if err != nil {
		fmt.Printf("Failed to copy problematic file: %v\n", err)
	}
}

// saveFormattedFile saves the formatted file
func (f *fileFormatter) saveFormattedFile() error {
	if len(f.result.ErrorMessages) > 0 {
		return f.handleFormattingError()
	}

	err := f.inputEditor.SaveAs(f.config.OutputFilePath)
	if err != nil {
		return fmt.Errorf("failed to save formatted file %s: %v", f.config.OutputFilePath, err)
	}

	return nil
}

// validateInputFiles validates that all required input files exist
func validateInputFiles(inputFilePath, targetFilePath, mappingFilePath string) error {
	if _, err := os.Stat(inputFilePath); os.IsNotExist(err) {
		return fmt.Errorf("input file not found: %s", inputFilePath)
	}

	if _, err := os.Stat(mappingFilePath); os.IsNotExist(err) {
		return fmt.Errorf("mapping file not found: %s", mappingFilePath)
	}

	if _, err := os.Stat(targetFilePath); os.IsNotExist(err) {
		return fmt.Errorf("target format file not found: %s", targetFilePath)
	}

	return nil
}

// createResultsDirectory creates the results directory if it doesn't exist
func createResultsDirectory() error {
	resultsDir := "data/results"
	return os.MkdirAll(resultsDir, 0755)
}

// getInputSheets returns all sheet names from the input file
func getInputSheets(inputFilePath string) ([]string, error) {
	inputEditor, err := OpenFile(inputFilePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open input file: %v", err)
	}
	defer inputEditor.Close()

	inputSheets := inputEditor.GetSheetNames()
	if len(inputSheets) == 0 {
		return nil, fmt.Errorf("no sheets found in input file")
	}

	return inputSheets, nil
}

// processAllSheets processes each sheet in the input file
func processAllSheets(inputFilePath, targetFilePath, mappingFilePath, targetSheet string, inputSheets []string) error {
	inputFileName := filepath.Base(inputFilePath)
	inputFileExt := filepath.Ext(inputFileName)
	inputFileBase := strings.TrimSuffix(inputFileName, inputFileExt)

	for _, sheetName := range inputSheets {
		outputFileName := fmt.Sprintf("%s-%s.xlsx", inputFileBase, sheetName)
		outputFilePath := filepath.Join("data/results", outputFileName)

		err := FormatFileWithTarget(
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

// adjustFormulaForRow adjusts formula for a specific row
// "=An+Cn+1" becomes "=A5+C5+1" for row 5
func adjustFormulaForRow(formula string, targetRow int) string {
	return strings.ReplaceAll(formula, "n", fmt.Sprintf("%d", targetRow))
}

// Helper function to copy files
func copyFile(src, dst string) error {
	input, err := os.ReadFile(src)
	if err != nil {
		return err
	}
	return os.WriteFile(dst, input, 0644)
}
