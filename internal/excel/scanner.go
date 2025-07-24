package excel

import (
	"bufio"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sheetFmt/internal/logger"
	"sort"
	"strings"
)

// cleanColumnName cleans column names by removing HTML tags, extra whitespace, and taking first line
func cleanColumnName(rawName string) string {
	if rawName == "" {
		return ""
	}

	// Convert to string and strip basic whitespace
	cleaned := strings.TrimSpace(rawName)

	// Remove HTML/XML tags using regex
	// This handles both self-closing and regular tags
	htmlTagRegex := regexp.MustCompile(`<[^>]+>`)
	cleaned = htmlTagRegex.ReplaceAllString(cleaned, "")

	// Split by newlines and take the first non-empty line
	lines := strings.Split(cleaned, "\n")
	var firstLine string
	for _, line := range lines {
		trimmedLine := strings.TrimSpace(line)
		if trimmedLine != "" {
			firstLine = trimmedLine
			break
		}
	}

	if firstLine == "" {
		return ""
	}

	// Remove extra whitespace and normalize (replace multiple spaces with single space)
	spaceRegex := regexp.MustCompile(`\s+`)
	cleaned = spaceRegex.ReplaceAllString(firstLine, " ")
	cleaned = strings.TrimSpace(cleaned)

	return cleaned
}

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
		logger.Warn("No .xlsx files found in directory", "directory", inputDir)
		return nil
	}

	logger.Info("Starting scan", "file_count", len(xlsxFiles), "directory", inputDir)

	// Set to store unique column names
	uniqueColumns := make(map[string]bool)
	// Track cleaning for debugging
	cleaningStats := make(map[string]string) // cleaned -> original (for first occurrence)

	// Process each Excel file
	for _, filePath := range xlsxFiles {
		fileName := filepath.Base(filePath)
		logger.Info("Scanning file", "file", fileName)

		err := scanFileColumns(filePath, uniqueColumns, cleaningStats)
		if err != nil {
			logger.Warn("Failed to scan file", "file", fileName, "error", err)
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

	// Log cleaning statistics
	cleanedCount := 0
	for cleaned, original := range cleaningStats {
		if cleaned != original {
			cleanedCount++
			logger.Debug("Column name cleaned", "original", original, "cleaned", cleaned)
		}
	}

	logger.Info("Column name processing completed",
		"total_unique_columns", len(columnNames),
		"cleaned_columns", cleanedCount)

	// Write to scanned_columns file in output directory
	outputFilePath := filepath.Join(outputDir, "scanned_columns")
	err = writeColumnsToFile(outputFilePath, columnNames)
	if err != nil {
		return fmt.Errorf("failed to write columns to file: %v", err)
	}

	logger.Info("Scanning completed successfully",
		"unique_columns", len(columnNames),
		"output_file", outputFilePath)

	return nil
}

// scanFileColumns scans all sheets in a single Excel file and adds column names to the set
func scanFileColumns(filePath string, uniqueColumns map[string]bool, cleaningStats map[string]string) error {
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
		logger.Debug("Scanning sheet", "sheet", sheetName, "file", filepath.Base(filePath))

		// Get column headers from this sheet
		headers, err := editor.GetColumnHeaders(sheetName)
		if err != nil {
			logger.Warn("Failed to read headers", "sheet", sheetName, "error", err)
			continue
		}

		// Add each header to the unique set after cleaning
		for _, header := range headers {
			rawHeader := strings.TrimSpace(header)
			if rawHeader == "" {
				continue
			}

			// Clean the header name
			cleanHeader := cleanColumnName(rawHeader)
			if cleanHeader != "" {
				uniqueColumns[cleanHeader] = true

				// Track cleaning for debugging (only store first occurrence)
				if _, exists := cleaningStats[cleanHeader]; !exists {
					cleaningStats[cleanHeader] = rawHeader
				}
			}
		}

		logger.Debug("Found headers in sheet", "sheet", sheetName, "header_count", len(headers))
	}

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
