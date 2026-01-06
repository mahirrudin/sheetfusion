package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/extrame/xls"
	"github.com/xuri/excelize/v2"
)

// convertXLStoXLSX converts a .xls file to .xlsx format
// Returns the path to the temporary .xlsx file
func convertXLStoXLSX(xlsPath string) (string, error) {
	fmt.Printf("  Converting .xls to .xlsx format...\n")

	// Open the XLS file
	xlsFile, err := xls.Open(xlsPath, "utf-8")
	if err != nil {
		return "", fmt.Errorf("failed to open .xls file: %w", err)
	}

	// Create a new XLSX file
	xlsx := excelize.NewFile()
	defer xlsx.Close()

	// Process each sheet in the XLS file
	for sheetIdx := 0; sheetIdx < xlsFile.NumSheets(); sheetIdx++ {
		sheet := xlsFile.GetSheet(sheetIdx)
		if sheet == nil {
			continue
		}

		sheetName := sheet.Name

		// Create sheet in XLSX (skip creating for the first sheet as it exists by default)
		if sheetIdx == 0 {
			// Rename the default sheet
			xlsx.SetSheetName("Sheet1", sheetName)
		} else {
			_, err := xlsx.NewSheet(sheetName)
			if err != nil {
				return "", fmt.Errorf("failed to create sheet '%s': %w", sheetName, err)
			}
		}

		// Copy data from XLS to XLSX
		maxRow := int(sheet.MaxRow)
		for rowIdx := 0; rowIdx <= maxRow; rowIdx++ {
			row := sheet.Row(rowIdx)
			if row == nil {
				continue
			}

			maxCol := row.LastCol()
			for colIdx := 0; colIdx <= int(maxCol); colIdx++ {
				cellValue := row.Col(colIdx)

				// Convert cell coordinates (1-indexed for excelize)
				cell, err := excelize.CoordinatesToCellName(colIdx+1, rowIdx+1)
				if err != nil {
					return "", fmt.Errorf("failed to convert coordinates: %w", err)
				}

				// Set cell value in XLSX
				if err := xlsx.SetCellValue(sheetName, cell, cellValue); err != nil {
					return "", fmt.Errorf("failed to set cell value: %w", err)
				}
			}
		}
	}

	// Generate temporary file path
	tempDir := os.TempDir()
	baseName := strings.TrimSuffix(filepath.Base(xlsPath), filepath.Ext(xlsPath))
	tempPath := filepath.Join(tempDir, fmt.Sprintf("%s_converted_%d.xlsx", baseName, os.Getpid()))

	// Save the XLSX file
	if err := xlsx.SaveAs(tempPath); err != nil {
		return "", fmt.Errorf("failed to save converted file: %w", err)
	}

	fmt.Printf("  ✓ Converted successfully\n")
	return tempPath, nil
}

// cleanupTempFiles removes temporary converted files
func cleanupTempFiles(files []string) {
	for _, file := range files {
		if err := os.Remove(file); err != nil {
			// Silently ignore errors during cleanup
			continue
		}
	}
}
