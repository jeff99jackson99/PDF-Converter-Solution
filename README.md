# Excel to PDF Converter

A powerful Python application that converts Excel files to professional PDF documents, specifically designed to solve printing issues with proforma financial documents.

## 🚀 Features

- **Excel to PDF Conversion**: Convert any Excel file to a professional PDF
- **Proforma Sheet Support**: Special handling for Assumptions, Proforma, Proforma Condensed, and Calculations sheets
- **Web Interface**: Beautiful Streamlit web app for easy file upload and conversion
- **Command Line Interface**: Batch processing and automation support
- **Professional Formatting**: Clean, readable PDF layouts with proper styling
- **Sheet Validation**: Automatic detection and validation of Excel sheets
- **Customizable Output**: Control rows, columns, and formatting options

## 🎯 Perfect For

- ✅ Proforma financial models
- ✅ Assumptions documentation  
- ✅ Calculation summaries
- ✅ Professional reports
- ✅ Solving Excel PDF printing issues

## 📋 Requirements

- Python 3.11+
- Excel files (.xlsx, .xls)

## 🛠️ Installation

### Option 1: Quick Setup

```bash
# Clone the repository
git clone https://github.com/yourusername/excel-pdf-converter.git
cd excel-pdf-converter

# Install dependencies
make setup

# Start the web app
make dev
```

### Option 2: Manual Installation

```bash
# Install dependencies
pip install -r requirements.txt

# Or install in development mode
pip install -e .
```

## 🖥️ Usage

### Web Interface (Recommended)

1. **Start the application:**
   ```bash
   make dev
   # or
   streamlit run streamlit_app.py
   ```

2. **Open your browser:**
   - Navigate to `http://localhost:8501`
   - Upload your Excel file
   - Configure settings
   - Click "Generate PDF"
   - Download your PDF

### Command Line Interface

#### Convert All Sheets
```bash
make convert FILE=path/to/your/file.xlsx
# or
python -m excel_pdf_converter path/to/your/file.xlsx
```

#### Convert Proforma Sheets Only
```bash
make convert-proforma FILE=path/to/your/file.xlsx
# or
python -m excel_pdf_converter path/to/your/file.xlsx --proforma-only
```

#### List Available Sheets
```bash
make list-sheets FILE=path/to/your/file.xlsx
# or
python -m excel_pdf_converter path/to/your/file.xlsx --list-sheets
```

#### Advanced Options
```bash
python -m excel_pdf_converter input.xlsx \
  --output "my_report.pdf" \
  --sheets "Assumptions" "Proforma" \
  --max-rows 25 \
  --max-cols 8 \
  --verbose
```

## 📁 Project Structure

```
excel-pdf-converter/
├── src/
│   └── excel_pdf_converter/
│       ├── __init__.py
│       ├── excel_reader.py      # Excel file reading
│       ├── pdf_generator.py     # PDF creation
│       ├── converter.py         # Main conversion logic
│       └── __main__.py          # CLI interface
├── streamlit_app.py             # Web interface
├── Makefile                     # Build commands
├── pyproject.toml               # Project configuration
├── requirements.txt             # Dependencies
└── README.md                    # This file
```

## 🔧 Development

### Setup Development Environment
```bash
make setup
```

### Run Tests
```bash
make test
```

### Code Formatting
```bash
make fmt
```

### Linting
```bash
make lint
```

### Clean Up
```bash
make clean
```

## 🐳 Docker Support

### Build Docker Image
```bash
make docker/build
```

### Run Docker Container
```bash
make docker/run
```

## ☁️ Deployment

### Streamlit Cloud
1. Push your code to GitHub
2. Connect your repository to [Streamlit Cloud](https://streamlit.io/cloud)
3. Deploy automatically with GitHub Actions

### GitHub Actions
The repository includes GitHub Actions workflows for:
- CI/CD pipeline with linting, type checking, and testing
- Automatic deployment to Streamlit Cloud

## 📊 Supported Excel Features

- ✅ Multiple sheets
- ✅ Formulas (displayed as values)
- ✅ Number formatting
- ✅ Text formatting
- ✅ Tables and data ranges
- ✅ Large datasets (with row/column limits)

## 🚨 Troubleshooting

### Common Issues

**"No sheets found"**
- Ensure your Excel file contains the expected sheet names
- Use `--list-sheets` to see available sheets

**"PDF generation failed"**
- Check file permissions
- Ensure sufficient disk space
- Verify Excel file is not corrupted

**"Memory error with large files"**
- Reduce `--max-rows` and `--max-cols` parameters
- Process sheets individually

### Getting Help

1. Check the logs with `--verbose` flag
2. Ensure all dependencies are installed
3. Verify Excel file format (.xlsx or .xls)

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Run tests and linting
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- Built with [Streamlit](https://streamlit.io/) for the web interface
- Uses [ReportLab](https://www.reportlab.com/) for PDF generation
- Powered by [pandas](https://pandas.pydata.org/) for data processing

## 📞 Support

For issues and questions:
- Open an issue on GitHub
- Check the troubleshooting section
- Review the command line help: `python -m excel_pdf_converter --help`
