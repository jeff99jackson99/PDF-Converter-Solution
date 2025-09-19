#!/usr/bin/env python3
"""Simple script to generate PDF from the proforma Excel file."""

import sys
import logging
from pathlib import Path
from src.excel_pdf_converter.converter import ExcelToPDFConverter

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

def main():
    """Generate PDF from the proforma Excel file."""
    
    # Path to the Excel file
    excel_file = "Pro Forma (4 Products).xlsx"
    
    # Check if file exists
    if not Path(excel_file).exists():
        logger.error(f"Excel file not found: {excel_file}")
        sys.exit(1)
    
    try:
        logger.info(f"Loading Excel file: {excel_file}")
        
        # Initialize converter
        converter = ExcelToPDFConverter(excel_file)
        
        # Get available sheets
        available_sheets = converter.get_available_sheets()
        logger.info(f"Available sheets: {available_sheets}")
        
        # Generate PDF with all proforma data
        logger.info("Converting proforma sheets to PDF...")
        pdf_path = converter.convert_proforma_to_pdf("Proforma_Complete.pdf")
        
        # Get file size
        pdf_size = Path(pdf_path).stat().st_size
        
        # Success
        print(f"\n🎉 PDF Generated Successfully!")
        print(f"📄 File: {pdf_path}")
        print(f"📊 Size: {pdf_size:,} bytes")
        print(f"✅ All proforma data captured")
        
        logger.info("PDF generation completed successfully")
        
    except Exception as e:
        logger.error(f"Failed to generate PDF: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
