# XLS to XLSX Conversion Scripts

These scripts help you convert `.xls` files (Excel 97-2003) to `.xlsx` format (Excel 2007+) before using SheetFusion.

## Linux Script (convert_xls_to_xlsx.sh)

Uses LibreOffice Calc command-line interface.

### Requirements
- LibreOffice Calc must be installed

**Install LibreOffice:**
```bash
# Ubuntu/Debian
sudo apt install libreoffice-calc

# Fedora
sudo dnf install libreoffice-calc

# Arch
sudo pacman -S libreoffice-fresh
```

### Usage

**Convert single file:**
```bash
./convert_xls_to_xlsx.sh input.xls
./convert_xls_to_xlsx.sh input.xls output.xlsx
```

**Convert all .xls files in a directory:**
```bash
./convert_xls_to_xlsx.sh /path/to/directory
```

### Examples
```bash
# Convert data.xls to data.xlsx
./convert_xls_to_xlsx.sh data.xls

# Convert with custom output name
./convert_xls_to_xlsx.sh old_data.xls new_data.xlsx

# Convert all .xls files in current directory
./convert_xls_to_xlsx.sh .

# Convert all .xls files in specific directory
./convert_xls_to_xlsx.sh ~/Documents/excel_files
```

---

## Windows Script (convert_xls_to_xlsx.ps1)

Uses Microsoft Excel COM automation.

### Requirements
- Microsoft Excel must be installed

### Usage

**Convert single file:**
```powershell
.\convert_xls_to_xlsx.ps1 -Input "input.xls"
.\convert_xls_to_xlsx.ps1 -Input "input.xls" -Output "output.xlsx"
```

**Convert all .xls files in a directory:**
```powershell
.\convert_xls_to_xlsx.ps1 -Input "C:\path\to\directory"
```

### Examples
```powershell
# Convert data.xls to data.xlsx
.\convert_xls_to_xlsx.ps1 -Input "data.xls"

# Convert with custom output name
.\convert_xls_to_xlsx.ps1 -Input "old_data.xls" -Output "new_data.xlsx"

# Convert all .xls files in current directory
.\convert_xls_to_xlsx.ps1 -Input "."

# Convert all .xls files in specific directory
.\convert_xls_to_xlsx.ps1 -Input "C:\Users\YourName\Documents\ExcelFiles"
```

### PowerShell Execution Policy

If you get an execution policy error, run PowerShell as Administrator and execute:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

## Workflow Example

**Complete workflow for merging .xls files:**

### Linux
```bash
# 1. Convert all .xls files to .xlsx
./convert_xls_to_xlsx.sh /path/to/xls/files

# 2. Merge the converted .xlsx files
./sheetfusion -i /path/to/xls/files -o merged.xlsx
```

### Windows
```powershell
# 1. Convert all .xls files to .xlsx
.\convert_xls_to_xlsx.ps1 -Input "C:\path\to\xls\files"

# 2. Merge the converted .xlsx files
.\sheetfusion.exe -i "C:\path\to\xls\files" -o "merged.xlsx"
```

---

## Features

✅ **Single file conversion** - Convert one file at a time
✅ **Batch conversion** - Convert entire directories
✅ **Custom output names** - Specify output filename
✅ **Auto-naming** - Automatically generates .xlsx filename
✅ **Error handling** - Clear error messages
✅ **Progress reporting** - Shows conversion status
✅ **Dependency checking** - Verifies required software is installed

---

## Troubleshooting

### Linux

**"libreoffice: command not found"**
- Install LibreOffice Calc (see Requirements section above)

**"Permission denied"**
```bash
chmod +x convert_xls_to_xlsx.sh
```

### Windows

**"Excel is not installed"**
- Install Microsoft Office with Excel

**"Execution policy error"**
- Run PowerShell as Administrator
- Execute: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

**"File is locked"**
- Close Excel if it's running
- Make sure the .xls file isn't open in another program

---

## Notes

- Converted files preserve all data and formatting
- Original .xls files are not deleted
- Scripts can be run multiple times safely
- Both scripts support relative and absolute paths
