#!/bin/bash
# XLS to XLSX Converter for Linux using LibreOffice Calc
# Usage: ./convert_xls_to_xlsx.sh <input.xls> [output.xlsx]
#        ./convert_xls_to_xlsx.sh <directory>

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Check if LibreOffice is installed
if ! command -v libreoffice &> /dev/null; then
    echo -e "${RED}Error: LibreOffice is not installed.${NC}"
    echo "Please install it using:"
    echo "  Ubuntu/Debian: sudo apt install libreoffice-calc"
    echo "  Fedora: sudo dnf install libreoffice-calc"
    echo "  Arch: sudo pacman -S libreoffice-fresh"
    exit 1
fi

# Function to convert a single file
convert_file() {
    local input_file="$1"
    local output_file="$2"
    
    # Get absolute paths
    input_abs=$(realpath "$input_file")
    output_dir=$(dirname "$input_abs")
    
    if [ -n "$output_file" ]; then
        output_abs=$(realpath "$output_file" 2>/dev/null || echo "$(pwd)/$output_file")
        output_dir=$(dirname "$output_abs")
        output_name=$(basename "$output_abs")
    else
        # Auto-generate output filename
        output_name=$(basename "$input_file" .xls).xlsx
    fi
    
    echo -e "${YELLOW}Converting: $input_file${NC}"
    
    # Convert using LibreOffice in headless mode
    if libreoffice --headless --convert-to xlsx:"Calc MS Excel 2007 XML" \
        --outdir "$output_dir" "$input_abs" &>/dev/null; then
        
        # Rename if custom output name was specified
        if [ -n "$output_file" ]; then
            converted_file="$output_dir/$(basename "$input_file" .xls).xlsx"
            if [ "$converted_file" != "$output_abs" ]; then
                mv "$converted_file" "$output_abs" 2>/dev/null || true
            fi
        fi
        
        echo -e "${GREEN}✓ Converted: $output_name${NC}"
        return 0
    else
        echo -e "${RED}✗ Failed to convert: $input_file${NC}"
        return 1
    fi
}

# Main script
if [ $# -eq 0 ]; then
    echo "XLS to XLSX Converter (Linux)"
    echo ""
    echo "Usage:"
    echo "  $0 <input.xls> [output.xlsx]    # Convert single file"
    echo "  $0 <directory>                   # Convert all .xls files in directory"
    echo ""
    echo "Examples:"
    echo "  $0 data.xls"
    echo "  $0 data.xls converted_data.xlsx"
    echo "  $0 /path/to/excel/files"
    exit 0
fi

input="$1"

# Check if input is a directory
if [ -d "$input" ]; then
    echo -e "${YELLOW}Converting all .xls files in: $input${NC}"
    echo ""
    
    count=0
    success=0
    failed=0
    
    # Use find to get all .xls files
    while IFS= read -r -d '' file; do
        ((count++))
        if convert_file "$file"; then
            ((success++))
        else
            ((failed++))
        fi
    done < <(find "$input" -maxdepth 1 -type f -name "*.xls" -print0)
    
    echo ""
    if [ $count -eq 0 ]; then
        echo -e "${YELLOW}No .xls files found in directory.${NC}"
    else
        echo -e "${GREEN}✓ Successfully converted $success of $count file(s)${NC}"
        if [ $failed -gt 0 ]; then
            echo -e "${RED}✗ Failed to convert $failed file(s)${NC}"
        fi
    fi
    
elif [ -f "$input" ]; then
    # Single file conversion
    if [[ "$input" != *.xls ]]; then
        echo -e "${RED}Error: Input file must have .xls extension${NC}"
        exit 1
    fi
    
    output="$2"
    convert_file "$input" "$output"
    echo ""
    echo -e "${GREEN}✓ Conversion complete!${NC}"
    
else
    echo -e "${RED}Error: File or directory not found: $input${NC}"
    exit 1
fi
