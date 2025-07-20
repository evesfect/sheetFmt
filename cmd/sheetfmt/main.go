package main

import (
	"fmt"
	"log"
	"sheetFmt/internal/config"
	"sheetFmt/internal/excel"
)

func main() {
	cfg, err := config.LoadConfig("configs/config.toml")
	if err != nil {
		log.Fatal("Error loading config:", err)
	}

	fmt.Println("\nScanning Excel files for column names...")
	err = excel.ScanAllColumnsInDirectory(cfg.Scan.InputDirectory, cfg.Scan.OutputDirectory)
	if err != nil {
		log.Fatal("Error scanning Excel files:", err)
	}

	editor, err := excel.OpenFile("data/input/example.xlsx")
	if err != nil {
		log.Fatal("Error opening file:", err)
	}

	err = editor.SetCellValue("Sheet1", "B2", "5")
	if err != nil {
		log.Fatal("Error setting cell value:", err)
	}
	editor.SaveAs("data/output/example.xlsx")
}
