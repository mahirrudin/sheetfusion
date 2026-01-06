package main

import (
	"fmt"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

// MergeOptions holds configuration for the merge operation
type MergeOptions struct {
	InputFiles []string
	OutputFile string
	SheetName  string // If empty, use first sheet from each file
	StartRow   int    // Row number to start merging from (1-indexed, 0 = use default behavior)
}

// mergeExcelFiles combines multiple Excel files into a single output file
func mergeExcelFiles(opts MergeOptions) error {
	// Track temporary files for cleanup
	var tempFiles []string
	defer cleanupTempFiles(tempFiles)

	// Preprocess: Convert any .xls files to .xlsx
	processedFiles := make([]string, 0, len(opts.InputFiles))
	for _, inputFile := range opts.InputFiles {
		if strings.ToLower(filepath.Ext(inputFile)) == ".xls" {
			// Convert .xls to .xlsx
			convertedPath, err := convertXLStoXLSX(inputFile)
			if err != nil {
				return fmt.Errorf("failed to convert %s: %w", inputFile, err)
			}
			processedFiles = append(processedFiles, convertedPath)
			tempFiles = append(tempFiles, convertedPath)
		} else {
			// Use .xlsx file as-is
			processedFiles = append(processedFiles, inputFile)
		}
	}

	// Create a new Excel file for output
	output := excelize.NewFile()
	defer output.Close()

	// Create the output sheet
	outputSheetName := "MergedData"
	outputSheetIndex, err := output.NewSheet(outputSheetName)
	if err != nil {
		return fmt.Errorf("failed to create output sheet: %w", err)
	}

	currentRow := 1
	isFirstFile := true

	// Process each input file
	for fileIdx, inputFile := range processedFiles {
		// Display original filename for user feedback
		originalFile := opts.InputFiles[fileIdx]
		fmt.Printf("Processing file %d/%d: %s\n", fileIdx+1, len(opts.InputFiles), originalFile)

		// Open the Excel file
		f, err := excelize.OpenFile(inputFile)
		if err != nil {
			return fmt.Errorf("failed to open file %s: %w", inputFile, err)
		}

		// Determine which sheet to read
		var sheetName string
		if opts.SheetName != "" {
			// Use specified sheet name
			sheetName = opts.SheetName
			// Verify the sheet exists
			sheetIndex, err := f.GetSheetIndex(sheetName)
			if err != nil || sheetIndex == -1 {
				f.Close()
				return fmt.Errorf("sheet '%s' not found in file %s", sheetName, inputFile)
			}
		} else {
			// Use the first sheet
			sheetName = f.GetSheetName(0)
			if sheetName == "" {
				f.Close()
				return fmt.Errorf("no sheets found in file %s", inputFile)
			}
		}

		// Read all rows from the sheet
		rows, err := f.GetRows(sheetName)
		if err != nil {
			f.Close()
			return fmt.Errorf("failed to read rows from sheet '%s' in file %s: %w", sheetName, inputFile, err)
		}

		// Write rows to output
		startRow := currentRow

		// Determine which rows to process
		var rowsToProcess [][]string
		if opts.StartRow > 0 {
			// User specified a start row
			if isFirstFile {
				// For first file, include rows from StartRow onwards
				if opts.StartRow <= len(rows) {
					rowsToProcess = rows[opts.StartRow-1:]
				}
			} else {
				// For subsequent files, skip to StartRow (assumes same structure)
				if opts.StartRow <= len(rows) {
					rowsToProcess = rows[opts.StartRow-1:]
				}
			}
		} else {
			// Default behavior: use all rows, skip header for subsequent files
			if !isFirstFile {
				// Skip header row for subsequent files (assuming first row is header)
				if len(rows) > 0 {
					rowsToProcess = rows[1:]
				}
			} else {
				rowsToProcess = rows
			}
		}

		for _, row := range rowsToProcess {
			for colIdx, cellValue := range row {
				cell, err := excelize.CoordinatesToCellName(colIdx+1, currentRow)
				if err != nil {
					f.Close()
					return fmt.Errorf("failed to convert coordinates: %w", err)
				}
				if err := output.SetCellValue(outputSheetName, cell, cellValue); err != nil {
					f.Close()
					return fmt.Errorf("failed to set cell value: %w", err)
				}
			}
			currentRow++
		}

		f.Close()

		if isFirstFile {
			isFirstFile = false
		}

		fmt.Printf("  Added %d rows (total rows now: %d)\n", currentRow-startRow, currentRow-1)
	}

	// Set the active sheet
	output.SetActiveSheet(outputSheetIndex)

	// Delete the default Sheet1 if it exists and is not our output sheet
	if outputSheetName != "Sheet1" {
		output.DeleteSheet("Sheet1")
	}

	// Save the output file
	if err := output.SaveAs(opts.OutputFile); err != nil {
		return fmt.Errorf("failed to save output file: %w", err)
	}

	fmt.Printf("\n✓ Successfully merged %d files into %s\n", len(opts.InputFiles), opts.OutputFile)
	fmt.Printf("  Total rows: %d\n", currentRow-1)

	return nil
}
