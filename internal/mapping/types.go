package mapping

import (
	"bufio"
	"encoding/json"
	"fmt"
	"os"
	"sheetFmt/internal/excel"
	"sheetFmt/internal/logger"
	"strings"
)

// ColumnMapping represents a mapping between scanned and target columns
type ColumnMapping struct {
	ScannedColumn string `json:"scanned_column"`
	TargetColumn  string `json:"target_column"`
	IsIgnored     bool   `json:"is_ignored"`
}

// MappingConfig holds all column mappings
type MappingConfig struct {
	Mappings []ColumnMapping `json:"mappings"`
}

// SaveToFile saves the mapping configuration to a JSON file
func (mc *MappingConfig) SaveToFile(filepath string) error {
	file, err := json.MarshalIndent(mc, "", "  ")
	if err != nil {
		return err
	}

	err = writeToFile(filepath, file)
	if err != nil {
		return err
	}

	logger.Info("Saved mapping configuration", "path", filepath, "mappings_count", len(mc.Mappings))
	return nil
}

// LoadFromFile loads mapping configuration from a JSON file
func LoadFromFile(filepath string) (*MappingConfig, error) {
	data, err := readFromFile(filepath)
	if err != nil {
		return nil, err
	}

	var config MappingConfig
	err = json.Unmarshal(data, &config)
	if err != nil {
		return nil, err
	}

	logger.Info("Loaded mapping configuration", "path", filepath, "mappings_count", len(config.Mappings))
	return &config, nil
}

// ReadColumnsFromFile reads column names from a text file (one per line)
func ReadColumnsFromFile(filepath string) ([]string, error) {
	file, err := os.Open(filepath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file %s: %v", filepath, err)
	}
	defer file.Close()

	var columns []string
	scanner := bufio.NewScanner(file)

	for scanner.Scan() {
		line := strings.TrimSpace(scanner.Text())
		if line != "" {
			columns = append(columns, line)
		}
	}

	if err := scanner.Err(); err != nil {
		return nil, fmt.Errorf("error reading file %s: %v", filepath, err)
	}

	logger.Info("Read columns from file", "path", filepath, "column_count", len(columns))
	return columns, nil
}

// CreateDefaultTargetColumnsFile creates a sample target_columns file if it doesn't exist
func CreateDefaultTargetColumnsFile(filepath string) error {
	if _, err := os.Stat(filepath); err == nil {
		return nil // File already exists
	}

	defaultColumns := []string{
		"ID",
		"Name",
		"Email",
		"Phone",
		"Address",
		"Date",
		"Amount",
		"Description",
	}

	file, err := os.Create(filepath)
	if err != nil {
		return fmt.Errorf("failed to create target columns file: %v", err)
	}
	defer file.Close()

	writer := bufio.NewWriter(file)
	defer writer.Flush()

	for _, column := range defaultColumns {
		_, err := writer.WriteString(column + "\n")
		if err != nil {
			return fmt.Errorf("failed to write column: %v", err)
		}
	}

	logger.Info("Created default target columns file", "path", filepath)
	return nil
}

// AppendTargetFormatHeadersToFile reads headers from target format file and appends unique ones to target_columns file
func AppendTargetFormatHeadersToFile(targetFormatFile, targetSheet, targetColumnsFile string) error {
	// Open the target format file
	editor, err := excel.OpenFile(targetFormatFile)
	if err != nil {
		return fmt.Errorf("failed to open target format file: %v", err)
	}
	defer editor.Close()

	// Get headers from the target format file
	headers, err := editor.GetColumnHeaders(targetSheet)
	if err != nil {
		return fmt.Errorf("failed to get headers from target format file: %v", err)
	}

	// Read existing target columns (if file exists)
	var existingColumns []string
	if _, err := os.Stat(targetColumnsFile); err == nil {
		existingColumns, err = ReadColumnsFromFile(targetColumnsFile)
		if err != nil {
			return fmt.Errorf("failed to read existing target columns: %v", err)
		}
	}

	// Create a map for quick lookup of existing columns
	existingMap := make(map[string]bool)
	for _, col := range existingColumns {
		existingMap[strings.TrimSpace(col)] = true
	}

	// Add unique headers from target format
	var newColumns []string
	addedCount := 0

	for _, header := range headers {
		trimmedHeader := strings.TrimSpace(header)
		if trimmedHeader != "" && !existingMap[trimmedHeader] {
			newColumns = append(newColumns, trimmedHeader)
			existingMap[trimmedHeader] = true
			addedCount++
		}
	}

	// Combine existing and new columns
	allColumns := append(existingColumns, newColumns...)

	// Write back to target_columns file
	err = writeColumnsToFile(targetColumnsFile, allColumns)
	if err != nil {
		return fmt.Errorf("failed to write updated target columns: %v", err)
	}

	logger.Info("Appended target format headers",
		"added_count", addedCount,
		"total_columns", len(allColumns),
		"target_file", targetColumnsFile)

	return nil
}

// writeColumnsToFile writes column names to a plain text file (helper function)
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

// Helper functions
func writeToFile(filepath string, data []byte) error {
	return os.WriteFile(filepath, data, 0644)
}

func readFromFile(filepath string) ([]byte, error) {
	return os.ReadFile(filepath)
}
