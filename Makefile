# Excel to PDF Converter Makefile
# Universal commands for Terminal, VS Code, and GitHub Actions

.PHONY: help setup dev test lint fmt clean docker/build docker/run streamlit

# Default target
help:
	@echo "Excel to PDF Converter - Available Commands:"
	@echo ""
	@echo "Setup:"
	@echo "  setup          Install dependencies and setup environment"
	@echo "  dev            Run Streamlit web application"
	@echo ""
	@echo "Development:"
	@echo "  test           Run tests"
	@echo "  lint           Run linting checks"
	@echo "  fmt            Format code"
	@echo ""
	@echo "PDF Generation:"
	@echo "  generate-pdf   Generate PDF from project Excel file (Pro Forma (4 Products).xlsx)"
	@echo "  convert        Convert any Excel file to PDF (usage: make convert FILE=path/to/file.xlsx)"
	@echo "  convert-proforma Convert proforma sheets only (usage: make convert-proforma FILE=path/to/file.xlsx)"
	@echo ""
	@echo "Docker:"
	@echo "  docker/build   Build Docker image"
	@echo "  docker/run     Run Docker container"
	@echo ""
	@echo "Cleanup:"
	@echo "  clean          Clean temporary files and output"

# Setup
setup:
	@echo "Setting up Excel to PDF Converter..."
	python -m pip install --upgrade pip
	pip install -e .
	pip install -e ".[dev]"
	@echo "Setup complete!"

# Development server
dev:
	@echo "Starting Streamlit application..."
	streamlit run streamlit_app.py --server.port 8501 --server.address 0.0.0.0

# Testing
test:
	@echo "Running tests..."
	python -m pytest tests/ -v

# Linting
lint:
	@echo "Running linting checks..."
	ruff check src/ streamlit_app.py tests/
	mypy src/ streamlit_app.py

# Formatting
fmt:
	@echo "Formatting code..."
	black src/ streamlit_app.py tests/
	ruff check --fix src/ streamlit_app.py tests/

# Generate PDF from project Excel file
generate-pdf:
	@echo "Generating PDF from Pro Forma (4 Products).xlsx..."
	python generate_pdf.py

# Convert any Excel file to PDF
convert:
	@if [ -z "$(FILE)" ]; then \
		echo "Error: Please specify FILE=path/to/your/file.xlsx"; \
		exit 1; \
	fi
	@echo "Converting $(FILE) to PDF..."
	python -m excel_pdf_converter "$(FILE)" -v

# Convert proforma sheets only from any file
convert-proforma:
	@if [ -z "$(FILE)" ]; then \
		echo "Error: Please specify FILE=path/to/your/file.xlsx"; \
		exit 1; \
	fi
	@echo "Converting proforma sheets from $(FILE) to PDF..."
	python -m excel_pdf_converter "$(FILE)" --proforma-only -v

# List sheets in Excel file
list-sheets:
	@if [ -z "$(FILE)" ]; then \
		echo "Error: Please specify FILE=path/to/your/file.xlsx"; \
		exit 1; \
	fi
	@echo "Listing sheets in $(FILE)..."
	python -m excel_pdf_converter "$(FILE)" --list-sheets

# Docker build
docker/build:
	@echo "Building Docker image..."
	docker build -t excel-pdf-converter .

# Docker run
docker/run:
	@echo "Running Docker container..."
	docker run -p 8501:8501 -v $(PWD)/output:/app/output excel-pdf-converter

# Clean up
clean:
	@echo "Cleaning temporary files..."
	find . -type f -name "*.pyc" -delete
	find . -type d -name "__pycache__" -delete
	find . -type d -name "*.egg-info" -exec rm -rf {} +
	rm -rf build/
	rm -rf dist/
	rm -rf .pytest_cache/
	rm -rf .ruff_cache/
	rm -rf .mypy_cache/
	rm -rf output/*.pdf
	@echo "Cleanup complete!"

# Quick start for new users
quick-start: setup
	@echo ""
	@echo "🚀 Quick Start Guide:"
	@echo "1. Generate PDF from project file: make generate-pdf"
	@echo "2. Or start the web app: make dev"
	@echo "3. Open browser: http://localhost:8501"
	@echo ""
	@echo "Command line options:"
	@echo "make generate-pdf                           # Use project Excel file"
	@echo "make convert FILE=path/to/your/file.xlsx    # Use any Excel file"
