package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"sheetFmt/internal/config"
	"sheetFmt/internal/excel"
	"sheetFmt/internal/mapping"
)

func main() {
	if len(os.Args) < 2 {
		printUsage()
		return
	}

	command := os.Args[1]

	cfg, err := config.LoadConfig("configs/config.toml")
	if err != nil {
		log.Fatal("Error loading config:", err)
	}

	switch command {
	case "scan":
		runScan(cfg)
	case "map":
		runMapping(cfg)
	case "format":
		fmt.Println("Format functionality coming soon...")
	default:
		fmt.Printf("Unknown command: %s\n", command)
		printUsage()
	}
}

func printUsage() {
	fmt.Println("SheetFmt - Excel Formatting Tool")
	fmt.Println("\nUsage:")
	fmt.Println("  sheetfmt scan    - Scan Excel files for column names")
	fmt.Println("  sheetfmt map     - Open interactive mapping tool")
	fmt.Println("  sheetfmt format  - Format Excel files (coming soon)")
}

func runScan(cfg *config.Config) {
	fmt.Println("\nScanning Excel files for column names...")
	err := excel.ScanAllColumnsInDirectory(cfg.Scan.InputDirectory, cfg.Scan.OutputDirectory)
	if err != nil {
		log.Fatal("Error scanning Excel files:", err)
	}
}

func runMapping(cfg *config.Config) {
	// File paths
	scannedColumnsFile := filepath.Join(cfg.Scan.OutputDirectory, "scanned_columns")
	targetColumnsFile := filepath.Join(cfg.Scan.OutputDirectory, "target_columns")
	mappingOutputFile := filepath.Join(cfg.Scan.OutputDirectory, "column_mapping.json")

	// Create target_columns file if it doesn't exist
	err := mapping.CreateDefaultTargetColumnsFile(targetColumnsFile)
	if err != nil {
		log.Fatal("Error creating target columns file:", err)
	}

	// Check if scanned_columns exists
	if _, err := os.Stat(scannedColumnsFile); os.IsNotExist(err) {
		fmt.Printf("❌ Scanned columns file not found: %s\n", scannedColumnsFile)
		fmt.Println("Please run 'sheetfmt scan' first to generate scanned columns.")
		return
	}

	fmt.Printf("📂 Using files:\n")
	fmt.Printf("   Scanned columns: %s\n", scannedColumnsFile)
	fmt.Printf("   Target columns:  %s\n", targetColumnsFile)
	fmt.Printf("   Output mapping:  %s\n", mappingOutputFile)
	fmt.Printf("📏 Grid: %dx%d (cols x rows)\n", cfg.UI.ColumnsPerRow, cfg.UI.RowsPerPage)
	fmt.Println()

	// Convert config to mapping UIConfig
	uiConfig := mapping.UIConfig{
		ColumnsPerRow: cfg.UI.ColumnsPerRow,
		RowsPerPage:   cfg.UI.RowsPerPage,
	}

	// Run the mapping TUI
	err = mapping.RunMappingTUI(scannedColumnsFile, targetColumnsFile, mappingOutputFile, uiConfig)
	if err != nil {
		log.Fatal("Error running mapping tool:", err)
	}
}
