.PHONY: all linux windows clean

# Binary names
BINARY_NAME=sheetfusion
LINUX_BINARY=$(BINARY_NAME)-linux-amd64
WINDOWS_BINARY=$(BINARY_NAME)-windows-amd64.exe

# Build directory
BUILD_DIR=build

all: linux windows

linux:
	@echo "Building for Linux..."
	@mkdir -p $(BUILD_DIR)
	GOOS=linux GOARCH=amd64 go build -o $(BUILD_DIR)/$(LINUX_BINARY) .
	@echo "✓ Linux binary created: $(BUILD_DIR)/$(LINUX_BINARY)"

windows:
	@echo "Building for Windows..."
	@mkdir -p $(BUILD_DIR)
	GOOS=windows GOARCH=amd64 go build -o $(BUILD_DIR)/$(WINDOWS_BINARY) .
	@echo "✓ Windows binary created: $(BUILD_DIR)/$(WINDOWS_BINARY)"

clean:
	@echo "Cleaning build artifacts..."
	@rm -rf $(BUILD_DIR)
	@echo "✓ Build directory cleaned"

test:
	@echo "Running tests..."
	go test -v ./...

deps:
	@echo "Downloading dependencies..."
	go mod download
	go mod tidy
	@echo "✓ Dependencies updated"
