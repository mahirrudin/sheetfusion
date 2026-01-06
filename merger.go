package main

import (
	"fmt"

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
	for fileIdx, inputFile := range opts.InputFiles {
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

		// Determine which rows to process and track source row indices
		var startRowIdx, endRowIdx int
		if opts.StartRow > 0 {
			// User specified a start row
			if opts.StartRow <= len(rows) {
				startRowIdx = opts.StartRow - 1
				endRowIdx = len(rows)
			} else {
				f.Close()
				continue
			}
		} else {
			// Default behavior: use all rows, skip header for subsequent files
			if !isFirstFile && len(rows) > 0 {
				startRowIdx = 1 // Skip header row
			} else {
				startRowIdx = 0
			}
			endRowIdx = len(rows)
		}

		// Process each row
		for sourceRowIdx := startRowIdx; sourceRowIdx < endRowIdx; sourceRowIdx++ {
			row := rows[sourceRowIdx]
			for colIdx := range row {
				// Get source cell coordinates (1-indexed)
				sourceCell, err := excelize.CoordinatesToCellName(colIdx+1, sourceRowIdx+1)
				if err != nil {
					f.Close()
					return fmt.Errorf("failed to convert source coordinates: %w", err)
				}

				// Get destination cell coordinates
				destCell, err := excelize.CoordinatesToCellName(colIdx+1, currentRow)
				if err != nil {
					f.Close()
					return fmt.Errorf("failed to convert dest coordinates: %w", err)
				}

				// Get the cell type to preserve formatting
				cellType, err := f.GetCellType(sheetName, sourceCell)
				var sourceStyle *excelize.Style

				if err == nil && cellType != excelize.CellTypeUnset {
					// Get cell style from source
					sourceStyleID, err := f.GetCellStyle(sheetName, sourceCell)
					if err == nil && sourceStyleID != 0 {
						// Get the actual style object from source
						sourceStyle, err = f.GetStyle(sourceStyleID)
						if err == nil && sourceStyle != nil {
							// Create a new style in the output workbook with the same properties
							newStyleID, err := output.NewStyle(sourceStyle)
							if err == nil {
								// Apply the new style to the destination cell
								output.SetCellStyle(outputSheetName, destCell, destCell, newStyleID)
							}
						}
					}
				}

				// Get and set cell value
				cellValue, err := f.GetCellValue(sheetName, sourceCell)
				if err != nil {
					f.Close()
					return fmt.Errorf("failed to get cell value: %w", err)
				}

				// Check if we should convert currency text to number
				if sourceStyle != nil && sourceStyle.CustomNumFmt != nil {
					customFmt := *sourceStyle.CustomNumFmt
					if shouldConvertToNumber(int(cellType), customFmt, cellValue) {
						// Parse currency text and set as number
						if numValue, ok := parseCurrencyText(cellValue); ok {
							if err := output.SetCellValue(outputSheetName, destCell, numValue); err != nil {
								f.Close()
								return fmt.Errorf("failed to set cell value: %w", err)
							}
							// Don't set the text value - we already set the numeric value
							continue
						}
					}
				}

				// Set the cell value as-is (for non-currency or if parsing failed)
				if err := output.SetCellValue(outputSheetName, destCell, cellValue); err != nil {
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
