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
	logger.Info("Starting directory scan", "input_dir", inputDir, "output_dir", outputDir)

	// Create the input directory if it doesn't exist
	if err := os.MkdirAll(inputDir, 0755); err != nil {
		logger.Error("Failed to create input directory", "directory", inputDir, "error", err)
		return fmt.Errorf("failed to create input directory: %v", err)
	}

	// Create the output directory if it doesn't exist
	if err := os.MkdirAll(outputDir, 0755); err != nil {
		logger.Error("Failed to create output directory", "directory", outputDir, "error", err)
		return fmt.Errorf("failed to create output directory: %v", err)
	}

	// Get all .xlsx files in the directory
	xlsxFiles, err := getXlsxFiles(inputDir)
	if err != nil {
		logger.Error("Failed to get xlsx files from directory", "directory", inputDir, "error", err)
		return fmt.Errorf("failed to get xlsx files: %v", err)
	}

	if len(xlsxFiles) == 0 {
		logger.Warn("No .xlsx files found in directory", "directory", inputDir)
		fmt.Printf("No .xlsx files found in directory: %s\n", inputDir)
		return nil
	}

	logger.Info("Excel files discovered", "file_count", len(xlsxFiles), "directory", inputDir)
	fmt.Printf("Found %d Excel files to scan\n", len(xlsxFiles))

	// Set to store unique column names
	uniqueColumns := make(map[string]bool)
	// Track cleaning for debugging
	cleaningStats := make(map[string]string) // cleaned -> original (for first occurrence)
	// Track where each unique header was first found
	headerSources := make(map[string]string) // cleaned -> first file name

	// Track scanning statistics
	var (
		totalFilesProcessed   = 0
		totalFilesWithErrors  = 0
		totalSheetsProcessed  = 0
		totalSheetsWithErrors = 0
		totalHeadersFound     = 0
		totalHeadersCleaned   = 0
		totalEmptyHeaders     = 0
	)

	// Process each Excel file
	for i, filePath := range xlsxFiles {
		fileName := filepath.Base(filePath)
		logger.Info("Processing file",
			"file", fileName,
			"progress", fmt.Sprintf("%d/%d", i+1, len(xlsxFiles)),
			"path", filePath)
		fmt.Printf("[%d/%d] Scanning: %s\n", i+1, len(xlsxFiles), fileName)

		fileStats, err := scanFileColumns(filePath, uniqueColumns, cleaningStats, headerSources)
		if err != nil {
			logger.Error("Failed to scan file completely",
				"file", fileName,
				"error", err,
				"file_path", filePath)
			fmt.Printf("Error scanning file: %s - %v\n", fileName, err)
			totalFilesWithErrors++
		} else {
			logger.Info("File scan completed successfully",
				"file", fileName,
				"sheets_processed", fileStats.SheetsProcessed,
				"sheets_with_errors", fileStats.SheetsWithErrors,
				"headers_found", fileStats.HeadersFound,
				"empty_headers", fileStats.EmptyHeaders)
			fmt.Printf("Successfully scanned: %s (%d sheets, %d headers)\n",
				fileName, fileStats.SheetsProcessed, fileStats.HeadersFound)
		}

		// Update totals
		totalFilesProcessed++
		totalSheetsProcessed += fileStats.SheetsProcessed
		totalSheetsWithErrors += fileStats.SheetsWithErrors
		totalHeadersFound += fileStats.HeadersFound
		totalEmptyHeaders += fileStats.EmptyHeaders
	}

	// Convert map to sorted slice
	columnNames := make([]string, 0, len(uniqueColumns))
	for column := range uniqueColumns {
		if strings.TrimSpace(column) != "" { // Skip empty column names
			columnNames = append(columnNames, column)
		}
	}
	sort.Strings(columnNames)

	// Log every single unique header found with first file
	logger.Info("ALL UNIQUE HEADERS FOUND", "total_count", len(columnNames))
	for i, header := range columnNames {
		original := cleaningStats[header]
		firstFile := headerSources[header]
		if original != header {
			logger.Info("UNIQUE HEADER", "index", i+1, "header", header, "original", original, "first_found_in", firstFile)
		} else {
			logger.Info("UNIQUE HEADER", "index", i+1, "header", header, "first_found_in", firstFile)
		}
	}

	// Count cleaned columns
	for cleaned, original := range cleaningStats {
		if cleaned != original {
			totalHeadersCleaned++
		}
	}

	// Log comprehensive scanning statistics
	logger.Info("Scan statistics",
		"total_files_found", len(xlsxFiles),
		"total_files_processed", totalFilesProcessed,
		"total_files_with_errors", totalFilesWithErrors,
		"total_sheets_processed", totalSheetsProcessed,
		"total_sheets_with_errors", totalSheetsWithErrors,
		"total_headers_found", totalHeadersFound,
		"total_headers_cleaned", totalHeadersCleaned,
		"total_empty_headers", totalEmptyHeaders,
		"unique_columns_final", len(columnNames))

	// Write to scanned_columns file in output directory
	outputFilePath := filepath.Join(outputDir, "scanned_columns")
	logger.Info("Writing scanned columns to file", "output_file", outputFilePath, "column_count", len(columnNames))

	err = writeColumnsToFile(outputFilePath, columnNames)
	if err != nil {
		logger.Error("Failed to write columns to file", "output_file", outputFilePath, "error", err)
		return fmt.Errorf("failed to write columns to file: %v", err)
	}

	logger.Info("Scanning completed successfully",
		"unique_columns", len(columnNames),
		"output_file", outputFilePath)

	// Print final summary
	fmt.Printf("\nScan Summary:\n")
	fmt.Printf("   Files processed: %d/%d\n", totalFilesProcessed-totalFilesWithErrors, len(xlsxFiles))
	fmt.Printf("   Sheets processed: %d\n", totalSheetsProcessed)
	fmt.Printf("   Headers found: %d\n", totalHeadersFound)
	fmt.Printf("   Unique columns: %d\n", len(columnNames))
	if totalHeadersCleaned > 0 {
		fmt.Printf("   Headers cleaned: %d\n", totalHeadersCleaned)
	}
	if totalFilesWithErrors > 0 {
		fmt.Printf("   Files with errors: %d\n", totalFilesWithErrors)
	}
	if totalSheetsWithErrors > 0 {
		fmt.Printf("   Sheets with errors: %d\n", totalSheetsWithErrors)
	}
	fmt.Printf("   Output: %s\n", outputFilePath)

	return nil
}

// FileStats holds statistics about scanning a single file
type FileStats struct {
	SheetsProcessed  int
	SheetsWithErrors int
	HeadersFound     int
	EmptyHeaders     int
}

// scanFileColumns scans all sheets in a single Excel file and adds column names to the set
func scanFileColumns(filePath string, uniqueColumns map[string]bool, cleaningStats map[string]string, headerSources map[string]string) (FileStats, error) {
	stats := FileStats{}
	fileName := filepath.Base(filePath)

	// Open the Excel file
	logger.Debug("Opening Excel file", "file", fileName, "path", filePath)
	editor, err := OpenFile(filePath)
	if err != nil {
		logger.Error("Failed to open Excel file", "file", fileName, "error", err, "path", filePath)
		return stats, fmt.Errorf("failed to open file: %v", err)
	}
	defer func() {
		if closeErr := editor.Close(); closeErr != nil {
			logger.Warn("Failed to close Excel file", "file", fileName, "error", closeErr)
		}
	}()

	// Get all sheet names
	sheetNames := editor.GetSheetNames()
	logger.Debug("Retrieved sheet names", "file", fileName, "sheet_count", len(sheetNames), "sheets", sheetNames)

	if len(sheetNames) == 0 {
		logger.Warn("No sheets found in Excel file", "file", fileName)
		return stats, fmt.Errorf("no sheets found in file")
	}

	// Process each sheet
	for sheetIndex, sheetName := range sheetNames {
		logger.Debug("Processing sheet",
			"sheet", sheetName,
			"file", fileName,
			"sheet_progress", fmt.Sprintf("%d/%d", sheetIndex+1, len(sheetNames)))

		// Get column headers from this sheet
		headers, err := editor.GetColumnHeaders(sheetName)
		if err != nil {
			logger.Error("Failed to read headers from sheet",
				"sheet", sheetName,
				"file", fileName,
				"error", err)
			stats.SheetsWithErrors++
			continue
		}

		logger.Debug("Retrieved headers from sheet",
			"sheet", sheetName,
			"file", fileName,
			"raw_header_count", len(headers))

		if len(headers) == 0 {
			logger.Warn("No headers found in sheet", "sheet", sheetName, "file", fileName)
			stats.SheetsProcessed++
			continue
		}

		// Track sheet-level statistics
		sheetHeadersFound := 0
		sheetEmptyHeaders := 0
		sheetCleanedHeaders := 0

		// Add each header to the unique set after cleaning
		for headerIndex, header := range headers {
			rawHeader := strings.TrimSpace(header)

			if rawHeader == "" {
				sheetEmptyHeaders++
				logger.Debug("Found empty header",
					"sheet", sheetName,
					"file", fileName,
					"header_index", headerIndex)
				continue
			}

			// Clean the header name
			cleanHeader := cleanColumnName(rawHeader)
			if cleanHeader != "" {
				// Track first occurrence of this header
				if !uniqueColumns[cleanHeader] {
					headerSources[cleanHeader] = fileName
				}

				uniqueColumns[cleanHeader] = true
				sheetHeadersFound++

				// Track cleaning statistics
				if cleanHeader != rawHeader {
					sheetCleanedHeaders++
					// Track cleaning for debugging (only store first occurrence)
					if _, exists := cleaningStats[cleanHeader]; !exists {
						cleaningStats[cleanHeader] = rawHeader
					}
					logger.Debug("Header cleaned",
						"sheet", sheetName,
						"file", fileName,
						"original", rawHeader,
						"cleaned", cleanHeader)
				}
			} else {
				logger.Debug("Header became empty after cleaning",
					"sheet", sheetName,
					"file", fileName,
					"header_index", headerIndex,
					"original", rawHeader)
				sheetEmptyHeaders++
			}
		}

		// Log sheet processing results
		logger.Info("Sheet processed",
			"sheet", sheetName,
			"file", fileName,
			"headers_found", sheetHeadersFound,
			"empty_headers", sheetEmptyHeaders,
			"cleaned_headers", sheetCleanedHeaders,
			"total_raw_headers", len(headers))

		// Update file statistics
		stats.SheetsProcessed++
		stats.HeadersFound += sheetHeadersFound
		stats.EmptyHeaders += sheetEmptyHeaders

		// Warn if sheet had no valid headers
		if sheetHeadersFound == 0 && len(headers) > 0 {
			logger.Warn("Sheet had headers but none were valid after processing",
				"sheet", sheetName,
				"file", fileName,
				"raw_headers_count", len(headers))
		}
	}

	logger.Debug("File processing completed",
		"file", fileName,
		"sheets_processed", stats.SheetsProcessed,
		"sheets_with_errors", stats.SheetsWithErrors,
		"total_headers_found", stats.HeadersFound,
		"total_empty_headers", stats.EmptyHeaders)

	return stats, nil
}

// getXlsxFiles returns all .xlsx files in the specified directory
func getXlsxFiles(dir string) ([]string, error) {
	logger.Debug("Scanning directory for Excel files", "directory", dir)

	var xlsxFiles []string

	err := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			logger.Warn("Error accessing path during directory walk", "path", path, "error", err)
			return err
		}

		if !info.IsDir() && strings.ToLower(filepath.Ext(path)) == ".xlsx" {
			// Check for potential issues with file
			if strings.HasPrefix(info.Name(), "~$") {
				logger.Debug("Skipping temporary Excel file", "file", info.Name(), "path", path)
				return nil
			}

			if info.Size() == 0 {
				logger.Warn("Found zero-size Excel file", "file", info.Name(), "path", path)
			}

			if info.Size() > 100*1024*1024 { // 100MB
				logger.Warn("Found very large Excel file", "file", info.Name(), "size_mb", info.Size()/(1024*1024), "path", path)
			}

			xlsxFiles = append(xlsxFiles, path)
			logger.Debug("Found Excel file", "file", info.Name(), "size_bytes", info.Size(), "path", path)
		}

		return nil
	})

	if err != nil {
		logger.Error("Error during directory walk", "directory", dir, "error", err)
		return nil, err
	}

	logger.Info("Directory scan completed", "directory", dir, "xlsx_files_found", len(xlsxFiles))
	return xlsxFiles, err
}

// writeColumnsToFile writes the column names to a plain text file
func writeColumnsToFile(filename string, columns []string) error {
	logger.Debug("Writing columns to file", "filename", filename, "column_count", len(columns))

	file, err := os.Create(filename)
	if err != nil {
		logger.Error("Failed to create output file", "filename", filename, "error", err)
		return fmt.Errorf("failed to create file: %v", err)
	}
	defer func() {
		if closeErr := file.Close(); closeErr != nil {
			logger.Warn("Failed to close output file", "filename", filename, "error", closeErr)
		}
	}()

	writer := bufio.NewWriter(file)
	defer func() {
		if flushErr := writer.Flush(); flushErr != nil {
			logger.Warn("Failed to flush output file buffer", "filename", filename, "error", flushErr)
		}
	}()

	for i, column := range columns {
		_, err := writer.WriteString(column + "\n")
		if err != nil {
			logger.Error("Failed to write column to file",
				"filename", filename,
				"column_index", i,
				"column", column,
				"error", err)
			return fmt.Errorf("failed to write column: %v", err)
		}
	}

	logger.Info("Successfully wrote columns to file", "filename", filename, "column_count", len(columns))
	return nil
}
