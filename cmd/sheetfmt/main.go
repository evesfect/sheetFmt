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
		if len(os.Args) < 3 {
			fmt.Println("Error: format command requires input file path")
			fmt.Println("Usage: sheetfmt format <input_file_path>")
			return
		}
		runFormat(cfg, os.Args[2])
	case "append-target-headers":
		runAppendTargetHeaders(cfg)
	case "convert-candidate":
		if len(os.Args) < 3 {
			fmt.Println("Error: convert-candidate command requires candidate file path")
			fmt.Println("Usage: sheetfmt convert-candidate <candidate_file_path>")
			return
		}
		runConvertCandidate(cfg, os.Args[2])
	default:
		fmt.Printf("Unknown command: %s\n", command)
		printUsage()
	}
}

func printUsage() {
	fmt.Println("SheetFmt - Excel Formatting Tool")
	fmt.Println("\nUsage:")
	fmt.Println("  sheetfmt scan                         - Scan Excel files for column names")
	fmt.Println("  sheetfmt map                          - Open interactive mapping tool")
	fmt.Println("  sheetfmt format <input_file>          - Format single Excel file")
	fmt.Println("  sheetfmt append-target-headers        - Add target format headers to target_columns file")
	fmt.Println("  sheetfmt convert-candidate <file>     - Convert candidate format to target format")
}

func runConvertCandidate(cfg *config.Config, candidateFilePath string) {
	fmt.Println("\nConverting candidate format to target format...")
	err := excel.ConvertCandidateToTargetFormat(
		candidateFilePath,
		cfg.Format.TargetFormatFile,
		cfg.Format.FormulaRow,
	)
	if err != nil {
		log.Fatal("Error converting candidate format:", err)
	}
}

func runScan(cfg *config.Config) {
	fmt.Println("\nScanning Excel files for column names...")
	err := excel.ScanAllColumnsInDirectory(cfg.Scan.InputDirectory, cfg.Scan.OutputDirectory)
	if err != nil {
		log.Fatal("Error scanning Excel files:", err)
	}
}

func runMapping(cfg *config.Config) {
	scannedColumnsFile := filepath.Join(cfg.Scan.OutputDirectory, "scanned_columns")
	targetColumnsFile := filepath.Join(cfg.Scan.OutputDirectory, "target_columns")
	mappingOutputFile := filepath.Join(cfg.Scan.OutputDirectory, "column_mapping.json")

	err := mapping.CreateDefaultTargetColumnsFile(targetColumnsFile)
	if err != nil {
		log.Fatal("Error creating target columns file:", err)
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
		log.Fatal("Error running mapping tool:", err)
	}
}

func runFormat(cfg *config.Config, inputFilePath string) {
	mappingFilePath := filepath.Join(cfg.Scan.OutputDirectory, "column_mapping.json")

	err := excel.FormatFile(
		inputFilePath,
		cfg.Format.TargetFormatFile,
		mappingFilePath,
		cfg.Format.TargetSheet,
		cfg.Format.FormulaRow,
		cfg.Format.TableEndTolerance,
		cfg.Format.CleanFormulaOnlyRows,
	)
	if err != nil {
		log.Fatal("Error formatting file:", err)
	}
}

func runAppendTargetHeaders(cfg *config.Config) {
	targetColumnsFile := filepath.Join(cfg.Scan.OutputDirectory, "target_columns")

	fmt.Println("\nAppending target format headers to target_columns file...")
	err := mapping.AppendTargetFormatHeadersToFile(
		cfg.Format.TargetFormatFile,
		cfg.Format.TargetSheet,
		targetColumnsFile,
	)
	if err != nil {
		log.Fatal("Error appending target headers:", err)
	}
}
