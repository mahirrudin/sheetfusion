package main

import (
	"strconv"
	"strings"
)

// parseCurrencyText attempts to parse a currency text value into a float64
// Examples: " $50,000 " -> 50000, " $(1,234)" -> -1234, " $0 " -> 0
func parseCurrencyText(text string) (float64, bool) {
	// Trim spaces
	text = strings.TrimSpace(text)

	if text == "" {
		return 0, false
	}

	// Check if it looks like currency (starts with $ or has $ somewhere)
	if !strings.Contains(text, "$") {
		return 0, false
	}

	// Remove currency symbol and spaces
	text = strings.ReplaceAll(text, "$", "")
	text = strings.TrimSpace(text)

	// Check for negative (parentheses format)
	isNegative := false
	if strings.HasPrefix(text, "(") && strings.HasSuffix(text, ")") {
		isNegative = true
		text = strings.Trim(text, "()")
		text = strings.TrimSpace(text)
	}

	// Remove thousand separators (commas)
	text = strings.ReplaceAll(text, ",", "")

	// Handle dash as zero
	if text == "-" || text == "" {
		return 0, true
	}

	// Try to parse as float
	value, err := strconv.ParseFloat(text, 64)
	if err != nil {
		return 0, false
	}

	if isNegative {
		value = -value
	}

	return value, true
}

// isCurrencyFormat checks if a custom number format is a currency format
func isCurrencyFormat(customFmt string) bool {
	if customFmt == "" {
		return false
	}
	// Check if format contains $ symbol
	return strings.Contains(customFmt, "$") || strings.Contains(customFmt, "\\$")
}

// shouldConvertToNumber checks if a cell should be converted from text to number
// Returns true if:
// - Cell type is string/inlineString
// - Cell has currency formatting
// - Cell value looks like a currency value
func shouldConvertToNumber(cellType int, customFmt string, value string) bool {
	// Check if it's a text cell (type 6=inlineString, 7=string)
	if cellType != 6 && cellType != 7 {
		return false
	}

	// Check if it has currency formatting
	if !isCurrencyFormat(customFmt) {
		return false
	}

	// Check if value looks like currency
	_, ok := parseCurrencyText(value)
	return ok
}
