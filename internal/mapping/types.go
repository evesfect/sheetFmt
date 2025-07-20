package mapping

import (
	"bufio"
	"encoding/json"
	"fmt"
	"os"
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

	return writeToFile(filepath, file)
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

	return nil
}

// Helper functions
func writeToFile(filepath string, data []byte) error {
	return os.WriteFile(filepath, data, 0644)
}

func readFromFile(filepath string) ([]byte, error) {
	return os.ReadFile(filepath)
}
