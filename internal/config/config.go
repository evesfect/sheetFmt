package config

import (
	"fmt"
	"os"
	"path/filepath"

	"github.com/BurntSushi/toml"
)

type Config struct {
	Scan ScanConfig `toml:"scan"`
	UI   UIConfig   `toml:"ui"`
}

type ScanConfig struct {
	InputDirectory  string `toml:"input_directory"`
	OutputDirectory string `toml:"output_directory"`
}

type UIConfig struct {
	ColumnsPerRow int `toml:"columns_per_row"`
	RowsPerPage   int `toml:"rows_per_page"`
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
		}

		err = SaveConfig(configPath, defaultConfig)
		if err != nil {
			return nil, fmt.Errorf("failed to create default config: %v", err)
		}

		fmt.Printf("Created default config file: %s\n", configPath)
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

	return &config, nil
}

// SaveConfig saves configuration to the specified config file path
func SaveConfig(configPath string, config *Config) error {
	file, err := os.Create(configPath)
	if err != nil {
		return fmt.Errorf("failed to create config file: %v", err)
	}
	defer file.Close()

	encoder := toml.NewEncoder(file)
	err = encoder.Encode(config)
	if err != nil {
		return fmt.Errorf("failed to encode config: %v", err)
	}

	return nil
}
