package main

import (
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func main() {
	// Exit on missing filename
	if len(os.Args) < 2 || os.Args[1] == "" {
		fmt.Println("Syntax: dumpXLSX <filename.xlsx>")
		os.Exit(1)
	}

	// Open file and panic on error
	fmt.Println("Reading ", os.Args[1])
	xlsx, err := excelize.OpenFile(os.Args[1])
	if err != nil {
		panic(err)
	}

	// Read all sheets in map
	for i, sheet := range xlsx.GetSheetMap() {
		//Output sheet header
		fmt.Printf("----- %d. %s -----\n", i, sheet)

		// Get rows
		rows := xlsx.GetRows(sheet)
		// Create a row number prefix pattern long enough to fit all row numbers
		prefixPattern := fmt.Sprintf("%% %dd ", len(fmt.Sprintf("%d", len(rows))))

		// Walk through rows
		for j, row := range rows {
			// Output row number as prefix
			fmt.Printf(prefixPattern, j)
			// Output row content
			for _, cell := range row {
				fmt.Print(cell, "\t")
			}
			fmt.Println()
		}
	}
}
