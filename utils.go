package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
)

// isExcelFile checks if a file has an Excel extension
// Supports both .xlsx (Excel 2007+) and .xls (Excel 97-2003)
func isExcelFile(filename string) bool {
	ext := strings.ToLower(filepath.Ext(filename))
	return ext == ".xlsx" || ext == ".xls"
}

// collectExcelFiles gathers all Excel files from the input specification
// Input can be comma-separated files or a directory path
func collectExcelFiles(input string) ([]string, error) {
	var files []string

	// Check if input is a directory
	info, err := os.Stat(input)
	if err == nil && info.IsDir() {
		// Read all files from directory
		entries, err := os.ReadDir(input)
		if err != nil {
			return nil, fmt.Errorf("failed to read directory: %w", err)
		}

		for _, entry := range entries {
			if !entry.IsDir() {
				fullPath := filepath.Join(input, entry.Name())
				if isExcelFile(fullPath) {
					files = append(files, fullPath)
				}
			}
		}

		if len(files) == 0 {
			return nil, fmt.Errorf("no Excel files found in directory: %s", input)
		}

		return files, nil
	}

	// Treat as comma-separated file list
	parts := strings.Split(input, ",")
	for _, part := range parts {
		filename := strings.TrimSpace(part)
		if filename == "" {
			continue
		}

		// Check if file exists
		if _, err := os.Stat(filename); os.IsNotExist(err) {
			return nil, fmt.Errorf("file not found: %s", filename)
		}

		if !isExcelFile(filename) {
			return nil, fmt.Errorf("not a supported Excel file (.xls or .xlsx): %s", filename)
		}

		files = append(files, filename)
	}

	if len(files) == 0 {
		return nil, fmt.Errorf("no valid Excel files specified")
	}

	return files, nil
}

// fileExists checks if a file exists
func fileExists(filename string) bool {
	_, err := os.Stat(filename)
	return err == nil
}
