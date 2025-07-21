package excel

import (
	"bufio"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
)

// ScanAllColumnsInDirectory scans all .xlsx files in the specified directory
// and extracts all unique column names from all sheets, saving them to scanned_columns file
func ScanAllColumnsInDirectory(inputDir, outputDir string) error {
	// Create the input directory if it doesn't exist
	if err := os.MkdirAll(inputDir, 0755); err != nil {
		return fmt.Errorf("failed to create input directory: %v", err)
	}

	// Create the output directory if it doesn't exist
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return fmt.Errorf("failed to create output directory: %v", err)
	}

	// Get all .xlsx files in the directory
	xlsxFiles, err := getXlsxFiles(inputDir)
	if err != nil {
		return fmt.Errorf("failed to get xlsx files: %v", err)
	}

	if len(xlsxFiles) == 0 {
		fmt.Printf("No .xlsx files found in directory: %s\n", inputDir)
		return nil
	}

	fmt.Printf("Found %d .xlsx files to scan\n", len(xlsxFiles))

	// Set to store unique column names
	uniqueColumns := make(map[string]bool)

	// Process each Excel file
	for _, filePath := range xlsxFiles {
		fmt.Printf("Scanning file: %s\n", filepath.Base(filePath))

		err := scanFileColumns(filePath, uniqueColumns)
		if err != nil {
			fmt.Printf("Warning: Failed to scan file %s: %v\n", filepath.Base(filePath), err)
			continue
		}
	}

	// Convert map to sorted slice
	columnNames := make([]string, 0, len(uniqueColumns))
	for column := range uniqueColumns {
		if strings.TrimSpace(column) != "" { // Skip empty column names
			columnNames = append(columnNames, column)
		}
	}
	sort.Strings(columnNames)

	// Write to scanned_columns file in output directory
	outputFilePath := filepath.Join(outputDir, "scanned_columns")
	err = writeColumnsToFile(outputFilePath, columnNames)
	if err != nil {
		return fmt.Errorf("failed to write columns to file: %v", err)
	}

	fmt.Printf("✓ Scanning completed successfully!\n")
	fmt.Printf("✓ Found %d unique column names across all files\n", len(columnNames))
	fmt.Printf("✓ Results saved to '%s' file\n", outputFilePath)

	return nil
}

// getXlsxFiles returns all .xlsx files in the specified directory
func getXlsxFiles(dir string) ([]string, error) {
	var xlsxFiles []string

	err := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		if !info.IsDir() && strings.ToLower(filepath.Ext(path)) == ".xlsx" {
			xlsxFiles = append(xlsxFiles, path)
		}

		return nil
	})

	return xlsxFiles, err
}

// scanFileColumns scans all sheets in a single Excel file and adds column names to the set
func scanFileColumns(filePath string, uniqueColumns map[string]bool) error {
	// Open the Excel file
	editor, err := OpenFile(filePath)
	if err != nil {
		return fmt.Errorf("failed to open file: %v", err)
	}
	defer editor.Close()

	// Get all sheet names
	sheetNames := editor.GetSheetNames()

	// Process each sheet
	for _, sheetName := range sheetNames {
		fmt.Printf("  - Scanning sheet: %s\n", sheetName)

		// Get column headers from this sheet
		headers, err := editor.GetColumnHeaders(sheetName)
		if err != nil {
			fmt.Printf("    Warning: Failed to read headers from sheet %s: %v\n", sheetName, err)
			continue
		}

		// Add each header to the unique set
		for _, header := range headers {
			trimmedHeader := strings.TrimSpace(header)
			if trimmedHeader != "" {
				uniqueColumns[trimmedHeader] = true
			}
		}

		fmt.Printf("    Found %d column headers\n", len(headers))
	}

	return nil
}

// writeColumnsToFile writes the column names to a plain text file
func writeColumnsToFile(filename string, columns []string) error {
	file, err := os.Create(filename)
	if err != nil {
		return fmt.Errorf("failed to create file: %v", err)
	}
	defer file.Close()

	writer := bufio.NewWriter(file)
	defer writer.Flush()

	for _, column := range columns {
		_, err := writer.WriteString(column + "\n")
		if err != nil {
			return fmt.Errorf("failed to write column: %v", err)
		}
	}

	return nil
}
