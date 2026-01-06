# SheetFusion

A powerful command-line tool to merge multiple Excel files into a single file.

## Features

- ✅ Merge multiple Excel files into one
- ✅ **Automatic .xls to .xlsx conversion** - no manual conversion needed!
- ✅ Accept individual files or entire directories
- ✅ Specify custom output filename
- ✅ Select specific sheet names to merge
- ✅ **Specify start row** - skip metadata or multiple header rows
- ✅ Automatic header detection and deduplication
- ✅ Cross-platform support (Linux & Windows)
- ✅ Progress reporting during merge

## Installation

### Download Pre-built Binaries

Download the appropriate binary for your platform from the `build/` directory:
- **Linux**: `sheetfusion-linux-amd64`
- **Windows**: `sheetfusion-windows-amd64.exe`

### Build from Source

Requirements:
- Go 1.16 or higher

```bash
# Clone or download the source
cd sheetfusion

# Download dependencies
go mod download

# Build for your current platform
go build -o sheetfusion .

# Or use Make to build for all platforms
make all
```

## Usage

### Basic Syntax

```bash
sheetfusion -input <files|directory> [options]
```

### Options

| Flag | Shorthand | Description | Default |
|------|-----------|-------------|---------|
| `-input` | `-i` | Comma-separated list of Excel files OR directory path | *required* |
| `-output` | `-o` | Output filename | `merged.xlsx` |
| `-sheet` | `-s` | Specific sheet name to merge | First sheet |
| `-startrow` | `-r` | Row number to start merging from (1-indexed) | Auto-detect |
| `-version` | | Show version information | |
| `-help` | | Show help message | |

### Examples

#### Merge specific files
```bash
sheetfusion -input "file1.xlsx,file2.xlsx" -output result.xlsx
```

#### Merge all Excel files in a directory
```bash
sheetfusion -input /path/to/directory -output combined.xlsx
```

#### Merge a specific sheet from multiple files
```bash
sheetfusion -i "data1.xlsx,data2.xlsx" -s "Sales" -o sales_merged.xlsx
```

#### Using short flags
```bash
sheetfusion -i "./excel_files" -o merged.xlsx -s "Sheet1"
```

#### Start merging from a specific row
```bash
# Skip first 2 rows (e.g., title and metadata) and start from row 3
sheetfusion -i "file1.xlsx,file2.xlsx" -r 3 -o merged.xlsx
```

## How It Works

1. **File Collection**: The tool scans the input (files or directory) for Excel files (.xlsx, .xls)
2. **Sheet Selection**: For each file, it reads either the specified sheet or the first sheet
3. **Header Handling**: The first file's header row is preserved; subsequent files skip their header rows
4. **Data Merging**: All data rows are combined sequentially into a single output sheet
5. **Output**: The merged data is saved to the specified output file

## Building for Different Platforms

### Using Make

```bash
# Build for Linux
make linux

# Build for Windows
make windows

# Build for both platforms
make all

# Clean build artifacts
make clean
```

### Manual Build

```bash
# Linux
GOOS=linux GOARCH=amd64 go build -o sheetfusion-linux-amd64 .

# Windows
GOOS=windows GOARCH=amd64 go build -o sheetfusion-windows-amd64.exe .
```

## Requirements

- Input files must be valid Excel files (`.xlsx` or `.xls`)
  - **`.xlsx` files** (Excel 2007+) are used directly
  - **`.xls` files** (Excel 97-2003) are automatically converted to `.xlsx` during processing
- All files must contain the specified sheet name (if using `-sheet` option)
- Files should have compatible data structures for meaningful merging

## Limitations

- Assumes the first row of each file is a header row
- Merges data sequentially (does not perform joins or complex data operations)
- Does not preserve formatting, formulas, or charts (data only)

## License

MIT License

## Author

Created with Go and the excellent [excelize](https://github.com/xuri/excelize) library.
