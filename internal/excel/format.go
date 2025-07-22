package excel

import (
	"fmt"
	"os"
	"os/exec"
	"strconv"
)

// FormatFileWithTarget formats a single input file using Python script
func FormatFileWithTarget(inputFilePath, targetFilePath, mappingFilePath, outputFilePath string, inputSheet, targetSheet string, formulaRow int) error {
	// Get the path to the Python script
	scriptPath := "internal/format/format_excel.py"

	// Check if Python script exists
	if _, err := os.Stat(scriptPath); os.IsNotExist(err) {
		return fmt.Errorf("python formatting script not found: %s", scriptPath)
	}

	// Run Python script
	cmd := exec.Command("python", scriptPath, inputFilePath, targetFilePath, mappingFilePath, targetSheet, outputFilePath, inputSheet, strconv.Itoa(formulaRow))

	// Set environment to use UTF-8 encoding for Python
	cmd.Env = append(os.Environ(), "PYTHONIOENCODING=utf-8")

	// Capture output
	output, err := cmd.CombinedOutput()

	if err != nil {
		fmt.Printf("Python script output: %s\n", string(output))
		return fmt.Errorf("python formatting failed: %v", err)
	}

	fmt.Printf("%s", string(output))
	return nil
}

// FormatFile formats an entire Excel file with all its sheets using Python script
func FormatFile(inputFilePath, targetFilePath, mappingFilePath, targetSheet string, formulaRow int, tableEndTolerance int, cleanFormulaOnlyRows bool) error {
	// Validate input files
	if err := validateInputFiles(inputFilePath, targetFilePath, mappingFilePath); err != nil {
		return err
	}

	// Get the path to the Python script
	scriptPath := "internal/format/format_excel.py"

	// Check if Python script exists
	if _, err := os.Stat(scriptPath); os.IsNotExist(err) {
		return fmt.Errorf("python formatting script not found: %s", scriptPath)
	}

	// Convert bool to string for Python
	cleanFlag := "false"
	if cleanFormulaOnlyRows {
		cleanFlag = "true"
	}

	// Run Python script for all sheets
	cmd := exec.Command("python", scriptPath, inputFilePath, targetFilePath, mappingFilePath, targetSheet, strconv.Itoa(formulaRow), strconv.Itoa(tableEndTolerance), cleanFlag)

	// Set environment to use UTF-8 encoding for Python
	cmd.Env = append(os.Environ(), "PYTHONIOENCODING=utf-8")

	// Capture output
	output, err := cmd.CombinedOutput()

	if err != nil {
		fmt.Printf("Python script output: %s\n", string(output))
		return fmt.Errorf("python formatting failed: %v", err)
	}

	fmt.Printf("%s", string(output))
	return nil
}

// validateInputFiles validates that all required input files exist
func validateInputFiles(inputFilePath, targetFilePath, mappingFilePath string) error {
	if _, err := os.Stat(inputFilePath); os.IsNotExist(err) {
		return fmt.Errorf("input file not found: %s", inputFilePath)
	}

	if _, err := os.Stat(mappingFilePath); os.IsNotExist(err) {
		return fmt.Errorf("mapping file not found: %s", mappingFilePath)
	}

	if _, err := os.Stat(targetFilePath); os.IsNotExist(err) {
		return fmt.Errorf("target format file not found: %s", targetFilePath)
	}

	return nil
}
