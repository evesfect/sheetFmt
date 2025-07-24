package main

import (
	"fmt"
	"os"
	"path/filepath"
	"sheetFmt/internal/config"
	"sheetFmt/internal/excel"
	"sheetFmt/internal/logger"
	"sheetFmt/internal/mapping"
	"strings"
)

func main() {
	if len(os.Args) < 2 {
		printUsage()
		return
	}

	command := os.Args[1]

	cfg, err := config.LoadConfig("configs/config.toml")
	if err != nil {
		logger.Error("Failed to load config", "error", err)
		fmt.Printf("Error loading config: %v\n", err)
		os.Exit(1)
	}

	switch command {
	case "scan":
		runScan(cfg)
	case "map":
		runMapping(cfg)
	case "format":
		if len(os.Args) < 3 {
			fmt.Println("Error: format command requires input file path")
			fmt.Println("Usage: sheetfmt format <input_file_path>")
			return
		}
		runFormat(cfg, os.Args[2])
	case "append-target-headers":
		runAppendTargetHeaders(cfg)
	case "format-all":
		runFormatAll(cfg)
	default:
		fmt.Printf("Unknown command: %s\n", command)
		printUsage()
	}
}

func runFormatAll(cfg *config.Config) {
	logger.Info("Starting format-all operation", "input_directory", cfg.Scan.InputDirectory)

	// Check if mapping file exists
	mappingFilePath := filepath.Join(cfg.Scan.OutputDirectory, "column_mapping.json")
	if _, err := os.Stat(mappingFilePath); os.IsNotExist(err) {
		fmt.Printf("Mapping file not found: %s\n", mappingFilePath)
		fmt.Println("Please run 'sheetfmt map' first to create column mappings.")
		return
	}

	// Get all .xlsx files in input directory
	xlsxFiles, err := getXlsxFiles(cfg.Scan.InputDirectory)
	if err != nil {
		logger.Error("Failed to get Excel files", "error", err)
		fmt.Printf("Error getting Excel files: %v\n", err)
		return
	}

	if len(xlsxFiles) == 0 {
		fmt.Printf("No .xlsx files found in directory: %s\n", cfg.Scan.InputDirectory)
		return
	}

	logger.Info("Found files to format", "file_count", len(xlsxFiles))

	// Create results directory
	resultsDir := filepath.Join(cfg.Scan.OutputDirectory, "results")
	err = os.MkdirAll(resultsDir, 0755)
	if err != nil {
		logger.Error("Failed to create results directory", "error", err)
		fmt.Printf("Error creating results directory: %v\n", err)
		return
	}

	// Track statistics
	successCount := 0
	errorCount := 0

	// Process each file
	for i, inputFile := range xlsxFiles {
		fileName := filepath.Base(inputFile)
		fmt.Printf("\n[%d/%d] Processing: %s\n", i+1, len(xlsxFiles), fileName)

		logger.Info("Processing file", "file", fileName, "progress", fmt.Sprintf("%d/%d", i+1, len(xlsxFiles)))

		err := excel.FormatFile(
			inputFile,
			cfg.Format.TargetFormatFile,
			mappingFilePath,
			cfg.Format.TargetSheet,
			cfg.Format.TableEndTolerance,
			cfg.Format.CleanFormulaOnlyRows,
		)

		if err != nil {
			logger.Error("Failed to format file", "file", fileName, "error", err)
			fmt.Printf("❌ Error formatting file: %v\n", err)
			errorCount++
		} else {
			logger.Info("Successfully formatted file", "file", fileName)
			fmt.Printf("✓ Successfully formatted\n")
			successCount++
		}
	}

	// Print summary
	logger.Info("Format-all operation completed",
		"success_count", successCount,
		"error_count", errorCount)

	fmt.Printf("\n========================================\n")
	fmt.Printf("Formatting complete!\n")
	fmt.Printf("✓ Success: %d files\n", successCount)
	if errorCount > 0 {
		fmt.Printf("❌ Errors: %d files\n", errorCount)
		fmt.Printf("Check data/problematic directory for files with errors\n")
	}
	fmt.Printf("Results saved to: %s\n", resultsDir)
}

// Helper function to get all .xlsx files in a directory
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

func printUsage() {
	fmt.Println("SheetFmt - Excel Formatting Tool")
	fmt.Println("\nUsage:")
	fmt.Println("  sheetfmt scan                         - Scan Excel files for column names")
	fmt.Println("  sheetfmt map                          - Open interactive mapping tool")
	fmt.Println("  sheetfmt format <input_file>          - Format single Excel file")
	fmt.Println("  sheetfmt format-all                   - Format all Excel files in input directory")
	fmt.Println("  sheetfmt append-target-headers        - Add target format headers to target_columns file")
}

func runScan(cfg *config.Config) {
	logger.Info("Starting scan operation")
	fmt.Println("\nScanning Excel files for column names...")
	err := excel.ScanAllColumnsInDirectory(cfg.Scan.InputDirectory, cfg.Scan.OutputDirectory)
	if err != nil {
		logger.Error("Scan operation failed", "error", err)
		fmt.Printf("Error scanning Excel files: %v\n", err)
		os.Exit(1)
	}
}

func runMapping(cfg *config.Config) {
	scannedColumnsFile := filepath.Join(cfg.Scan.OutputDirectory, "scanned_columns")
	targetColumnsFile := filepath.Join(cfg.Scan.OutputDirectory, "target_columns")
	mappingOutputFile := filepath.Join(cfg.Scan.OutputDirectory, "column_mapping.json")

	logger.Info("Starting mapping operation",
		"scanned_file", scannedColumnsFile,
		"target_file", targetColumnsFile,
		"output_file", mappingOutputFile)

	err := mapping.CreateDefaultTargetColumnsFile(targetColumnsFile)
	if err != nil {
		logger.Error("Failed to create target columns file", "error", err)
		fmt.Printf("Error creating target columns file: %v\n", err)
		os.Exit(1)
	}

	if _, err := os.Stat(scannedColumnsFile); os.IsNotExist(err) {
		fmt.Printf("Scanned columns file not found: %s\n", scannedColumnsFile)
		fmt.Println("Please run 'sheetfmt scan' first to generate scanned columns.")
		return
	}

	fmt.Printf("Using files:\n")
	fmt.Printf("   Scanned columns: %s\n", scannedColumnsFile)
	fmt.Printf("   Target columns:  %s\n", targetColumnsFile)
	fmt.Printf("   Output mapping:  %s\n", mappingOutputFile)
	fmt.Printf("Grid: %dx%d (cols x rows)\n", cfg.UI.ColumnsPerRow, cfg.UI.RowsPerPage)
	fmt.Println()

	uiConfig := mapping.UIConfig{
		ColumnsPerRow: cfg.UI.ColumnsPerRow,
		RowsPerPage:   cfg.UI.RowsPerPage,
	}

	err = mapping.RunMappingTUI(scannedColumnsFile, targetColumnsFile, mappingOutputFile, uiConfig)
	if err != nil {
		logger.Error("Mapping operation failed", "error", err)
		fmt.Printf("Error running mapping tool: %v\n", err)
		os.Exit(1)
	}
}

func runFormat(cfg *config.Config, inputFilePath string) {
	mappingFilePath := filepath.Join(cfg.Scan.OutputDirectory, "column_mapping.json")

	logger.Info("Starting format operation", "input_file", inputFilePath)

	err := excel.FormatFile(
		inputFilePath,
		cfg.Format.TargetFormatFile,
		mappingFilePath,
		cfg.Format.TargetSheet,
		cfg.Format.TableEndTolerance,
		cfg.Format.CleanFormulaOnlyRows,
	)
	if err != nil {
		logger.Error("Format operation failed", "error", err)
		fmt.Printf("Error formatting file: %v\n", err)
		os.Exit(1)
	}
}

func runAppendTargetHeaders(cfg *config.Config) {
	targetColumnsFile := filepath.Join(cfg.Scan.OutputDirectory, "target_columns")

	logger.Info("Starting append target headers operation")
	fmt.Println("\nAppending target format headers to target_columns file...")
	err := mapping.AppendTargetFormatHeadersToFile(
		cfg.Format.TargetFormatFile,
		cfg.Format.TargetSheet,
		targetColumnsFile,
	)
	if err != nil {
		logger.Error("Append target headers operation failed", "error", err)
		fmt.Printf("Error appending target headers: %v\n", err)
		os.Exit(1)
	}
}
