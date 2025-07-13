package main

import (
	"fmt"
	"log"
	"sheetFmt/excel"
)

func main() {

	targetFile := "target_format.xlsx"
	editedFile := "edited_file.xlsx"

	err := excel.ApplyTargetFormatToFile(targetFile, editedFile, "Sheet1", "Sheet1")
	if err != nil {
		log.Fatal("Error applying target format:", err)
	}

	fmt.Println("✓ Target format applied successfully!")
	fmt.Printf("✓ Check '%s' to see the results\n", editedFile)
}
