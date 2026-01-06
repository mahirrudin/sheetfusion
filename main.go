package main

import (
	"flag"
	"fmt"
	"os"
)

const version = "1.0.0"

func main() {
	// Define command-line flags
	input := flag.String("input", "", "Comma-separated list of Excel files OR directory path (required)")
	inputShort := flag.String("i", "", "Shorthand for -input")
	output := flag.String("output", "merged.xlsx", "Output filename")
	outputShort := flag.String("o", "merged.xlsx", "Shorthand for -output")
	sheet := flag.String("sheet", "", "Specific sheet name to merge (default: first sheet)")
	sheetShort := flag.String("s", "", "Shorthand for -sheet")
	startRow := flag.Int("startrow", 0, "Row number to start merging from (1-indexed, default: auto-detect)")
	startRowShort := flag.Int("r", 0, "Shorthand for -startrow")
	showVersion := flag.Bool("version", false, "Show version information")
	showHelp := flag.Bool("help", false, "Show help message")

	flag.Parse()

	// Show version
	if *showVersion {
		fmt.Printf("SheetFusion v%s\n", version)
		fmt.Println("Excel file merger tool")
		os.Exit(0)
	}

	// Show help
	if *showHelp {
		printHelp()
		os.Exit(0)
	}

	// Determine which flag was used (prefer long form)
	inputValue := *input
	if inputValue == "" {
		inputValue = *inputShort
	}

	outputValue := *output
	if *outputShort != "merged.xlsx" {
		outputValue = *outputShort
	}

	sheetValue := *sheet
	if sheetValue == "" {
		sheetValue = *sheetShort
	}

	startRowValue := *startRow
	if *startRowShort != 0 {
		startRowValue = *startRowShort
	}

	// Validate input
	if inputValue == "" {
		fmt.Println("Error: input is required")
		fmt.Println()
		printHelp()
		os.Exit(1)
	}

	// Collect Excel files
	fmt.Println("SheetFusion - Excel File Merger")
	fmt.Println("================================")
	fmt.Println()

	files, err := collectExcelFiles(inputValue)
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}

	fmt.Printf("Found %d Excel file(s) to merge:\n", len(files))
	for i, file := range files {
		fmt.Printf("  %d. %s\n", i+1, file)
	}
	fmt.Println()

	if sheetValue != "" {
		fmt.Printf("Target sheet: %s\n", sheetValue)
	} else {
		fmt.Println("Target sheet: First sheet in each file")
	}
	if startRowValue > 0 {
		fmt.Printf("Start row: %d\n", startRowValue)
	} else {
		fmt.Println("Start row: Auto-detect (skip headers for subsequent files)")
	}
	fmt.Printf("Output file: %s\n", outputValue)
	fmt.Println()

	// Perform merge
	opts := MergeOptions{
		InputFiles: files,
		OutputFile: outputValue,
		SheetName:  sheetValue,
		StartRow:   startRowValue,
	}

	if err := mergeExcelFiles(opts); err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}
}

func printHelp() {
	fmt.Println("SheetFusion - Excel File Merger")
	fmt.Println()
	fmt.Println("USAGE:")
	fmt.Println("  sheetfusion -input <files|directory> [options]")
	fmt.Println()
	fmt.Println("OPTIONS:")
	fmt.Println("  -input, -i      Comma-separated list of Excel files OR directory path (required)")
	fmt.Println("  -output, -o     Output filename (default: merged.xlsx)")
	fmt.Println("  -sheet, -s      Specific sheet name to merge (default: first sheet)")
	fmt.Println("  -startrow, -r   Row number to start merging from (1-indexed, default: auto)")
	fmt.Println("  -version        Show version information")
	fmt.Println("  -help           Show this help message")
	fmt.Println()
	fmt.Println("SUPPORTED FORMATS:")
	fmt.Println("  - .xlsx files (Excel 2007+) ✓")
	fmt.Println("  - .xls files (Excel 97-2003) ✗ Not supported - please convert to .xlsx first")
	fmt.Println()
	fmt.Println("EXAMPLES:")
	fmt.Println("  # Merge specific files")
	fmt.Println("  sheetfusion -input \"file1.xlsx,file2.xlsx\" -output result.xlsx")
	fmt.Println()
	fmt.Println("  # Merge all Excel files in a directory")
	fmt.Println("  sheetfusion -input /path/to/directory -output combined.xlsx")
	fmt.Println()
	fmt.Println("  # Merge specific sheet from multiple files")
	fmt.Println("  sheetfusion -i \"data1.xlsx,data2.xlsx\" -s \"Sales\" -o sales_merged.xlsx")
	fmt.Println()
	fmt.Println("  # Start merging from row 3 (skip first 2 rows)")
	fmt.Println("  sheetfusion -i \"file1.xlsx,file2.xlsx\" -r 3 -o merged.xlsx")
}
