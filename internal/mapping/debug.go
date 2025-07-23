package mapping

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"time"
)

var debugLogger *log.Logger

func initDebugLogger() {
	// Create logs directory
	logDir := "logs"
	os.MkdirAll(logDir, 0755)

	// Create log file with timestamp
	logFile := filepath.Join(logDir, fmt.Sprintf("ai_mapping_%s.log", time.Now().Format("2006-01-02_15-04-05")))

	file, err := os.OpenFile(logFile, os.O_CREATE|os.O_WRONLY|os.O_APPEND, 0666)
	if err != nil {
		return
	}

	debugLogger = log.New(file, "", log.LstdFlags)
	debugLogger.Printf("=== AI Mapping Session Started ===")
}

func debugLog(format string, args ...interface{}) {
	if debugLogger != nil {
		debugLogger.Printf(format, args...)
	}
}

func saveAIMappingsToFile(unmappedColumns, targetColumns []string, aiMappings []AIMapping, err error) {
	// Create debug directory
	debugDir := "logs/ai_debug"
	os.MkdirAll(debugDir, 0755)

	timestamp := time.Now().Format("2006-01-02_15-04-05")
	debugFile := filepath.Join(debugDir, fmt.Sprintf("ai_mapping_%s.txt", timestamp))

	file, fileErr := os.Create(debugFile)
	if fileErr != nil {
		return
	}
	defer file.Close()

	fmt.Fprintf(file, "AI Mapping Debug - %s\n", time.Now().Format("2006-01-02 15:04:05"))
	fmt.Fprintf(file, "===========================================\n\n")

	fmt.Fprintf(file, "UNMAPPED COLUMNS SENT TO AI (%d):\n", len(unmappedColumns))
	for i, col := range unmappedColumns {
		fmt.Fprintf(file, "%d. %s\n", i+1, col)
	}

	fmt.Fprintf(file, "\nTARGET COLUMNS (%d):\n", len(targetColumns))
	for i, col := range targetColumns {
		fmt.Fprintf(file, "%d. %s\n", i+1, col)
	}

	fmt.Fprintf(file, "\nAI RESPONSE:\n")
	if err != nil {
		fmt.Fprintf(file, "ERROR: %v\n", err)
	} else {
		fmt.Fprintf(file, "SUCCESS - Generated %d mappings:\n", len(aiMappings))
		for i, mapping := range aiMappings {
			fmt.Fprintf(file, "%d. '%s' → '%s' (%.2f confidence)\n",
				i+1, mapping.ScannedColumn, mapping.TargetColumn, mapping.Confidence)
		}

		if len(aiMappings) == 0 {
			fmt.Fprintf(file, "No mappings generated (all were NO_MATCH or low confidence)\n")
		}
	}

	fmt.Fprintf(file, "\n===========================================\n")
}
