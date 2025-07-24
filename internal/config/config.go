package config

import (
	"fmt"
	"os"
	"path/filepath"
	"sheetFmt/internal/logger"

	"github.com/BurntSushi/toml"
)

type Config struct {
	Scan   ScanConfig   `toml:"scan"`
	UI     UIConfig     `toml:"ui"`
	Format FormatConfig `toml:"format"`
}

type ScanConfig struct {
	InputDirectory  string `toml:"input_directory"`
	OutputDirectory string `toml:"output_directory"`
}

type UIConfig struct {
	ColumnsPerRow int `toml:"columns_per_row"`
	RowsPerPage   int `toml:"rows_per_page"`
}

type FormatConfig struct {
	TargetFormatFile     string `toml:"target_format_file"`
	TargetSheet          string `toml:"target_sheet"`
	TableEndTolerance    int    `toml:"table_end_tolerance"`
	CleanFormulaOnlyRows bool   `toml:"clean_formula_only_rows"`
}

// LoadConfig loads configuration from the specified config file path
func LoadConfig(configPath string) (*Config, error) {
	// Check if config file exists
	if _, err := os.Stat(configPath); os.IsNotExist(err) {
		// Create configs directory if it doesn't exist
		configDir := filepath.Dir(configPath)
		if err := os.MkdirAll(configDir, 0755); err != nil {
			return nil, fmt.Errorf("failed to create config directory: %v", err)
		}

		// Create default config file
		defaultConfig := &Config{
			Scan: ScanConfig{
				InputDirectory:  "data/input",
				OutputDirectory: "data/output",
			},
			UI: UIConfig{
				ColumnsPerRow: 6,
				RowsPerPage:   2,
			},
			Format: FormatConfig{
				TargetFormatFile:     "configs/target_format.xlsx",
				TargetSheet:          "Sheet1",
				TableEndTolerance:    1,
				CleanFormulaOnlyRows: true,
			},
		}

		err = SaveConfig(configPath, defaultConfig)
		if err != nil {
			return nil, fmt.Errorf("failed to create default config: %v", err)
		}

		logger.Info("Created default config file", "path", configPath)
		return defaultConfig, nil
	}

	// Load existing config
	var config Config
	_, err := toml.DecodeFile(configPath, &config)
	if err != nil {
		return nil, fmt.Errorf("failed to load config file %s: %v", configPath, err)
	}

	// Set defaults if missing
	if config.UI.ColumnsPerRow == 0 {
		config.UI.ColumnsPerRow = 6
	}
	if config.UI.RowsPerPage == 0 {
		config.UI.RowsPerPage = 2
	}
	if config.Format.TargetFormatFile == "" {
		config.Format.TargetFormatFile = "configs/target_format.xlsx"
	}
	if config.Format.TargetSheet == "" {
		config.Format.TargetSheet = "Sheet1"
	}
	if config.Format.TableEndTolerance == 0 {
		config.Format.TableEndTolerance = 1
	}

	logger.Info("Loaded configuration", "path", configPath)
	return &config, nil
}

// SaveConfig saves configuration to the specified config file path
func SaveConfig(configPath string, config *Config) error {
	file, err := os.Create(configPath)
	if err != nil {
		return fmt.Errorf("failed 	to create config file: %v", err)
	}
	defer file.Close()

	encoder := toml.NewEncoder(file)
	err = encoder.Encode(config)
	if err != nil {
		return fmt.Errorf("failed to encode config: %v", err)
	}

	logger.Info("Saved configuration", "path", configPath)
	return nil
}
