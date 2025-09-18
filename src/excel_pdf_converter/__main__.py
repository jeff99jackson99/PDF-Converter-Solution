"""Command-line interface for Excel to PDF converter."""

import argparse
import logging
import sys
from pathlib import Path
from typing import List, Optional

from .converter import ExcelToPDFConverter

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def main() -> None:
    """Main entry point for the command-line interface."""
    parser = argparse.ArgumentParser(
        description="Convert Excel files to PDF, specifically designed for proforma documents"
    )
    
    parser.add_argument(
        "excel_file",
        help="Path to the Excel file to convert"
    )
    
    parser.add_argument(
        "-o", "--output",
        help="Output PDF filename (default: based on Excel filename)",
        default=None
    )
    
    parser.add_argument(
        "-d", "--output-dir",
        help="Output directory for PDF (default: ./output)",
        default="output"
    )
    
    parser.add_argument(
        "-s", "--sheets",
        nargs="+",
        help="Specific sheets to convert (default: all sheets)",
        default=None
    )
    
    parser.add_argument(
        "--proforma-only",
        action="store_true",
        help="Convert only proforma sheets (Assumptions, Proforma, Proforma Condensed, Calculations)"
    )
    
    parser.add_argument(
        "--max-rows",
        type=int,
        default=30,
        help="Maximum rows per sheet in PDF (default: 30)"
    )
    
    parser.add_argument(
        "--max-cols",
        type=int,
        default=10,
        help="Maximum columns per sheet in PDF (default: 10)"
    )
    
    parser.add_argument(
        "--no-summaries",
        action="store_true",
        help="Skip sheet summaries in PDF"
    )
    
    parser.add_argument(
        "--list-sheets",
        action="store_true",
        help="List available sheets and exit"
    )
    
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose logging"
    )
    
    args = parser.parse_args()
    
    # Set logging level
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    # Validate input file
    excel_path = Path(args.excel_file)
    if not excel_path.exists():
        logger.error(f"Excel file not found: {excel_path}")
        sys.exit(1)
    
    if not excel_path.suffix.lower() in ['.xlsx', '.xls']:
        logger.error(f"File must be an Excel file (.xlsx or .xls): {excel_path}")
        sys.exit(1)
    
    try:
        # Initialize converter
        converter = ExcelToPDFConverter(str(excel_path), args.output_dir)
        
        # List sheets if requested
        if args.list_sheets:
            sheets = converter.get_available_sheets()
            print(f"Available sheets in {excel_path.name}:")
            for i, sheet in enumerate(sheets, 1):
                print(f"  {i}. {sheet}")
            sys.exit(0)
        
        # Load sheets
        if args.proforma_only:
            logger.info("Loading proforma sheets...")
            sheets_data = converter.load_proforma_sheets()
            if not sheets_data:
                logger.error("No proforma sheets found in the file")
                sys.exit(1)
        else:
            logger.info("Loading sheets...")
            sheets_data = converter.load_sheets(args.sheets)
        
        # Show loaded sheets
        logger.info(f"Loaded {len(sheets_data)} sheets: {list(sheets_data.keys())}")
        
        # Convert to PDF
        logger.info("Converting to PDF...")
        pdf_path = converter.convert_to_pdf(
            pdf_filename=args.output,
            include_sheet_summaries=not args.no_summaries,
            max_rows_per_sheet=args.max_rows,
            max_cols_per_sheet=args.max_cols
        )
        
        # Success message
        pdf_size = Path(pdf_path).stat().st_size
        logger.info(f"✅ PDF generated successfully!")
        logger.info(f"📄 File: {pdf_path}")
        logger.info(f"📊 Size: {pdf_size:,} bytes")
        logger.info(f"📋 Sheets: {len(sheets_data)}")
        
        print(f"\n🎉 Conversion completed successfully!")
        print(f"📄 PDF saved as: {pdf_path}")
        print(f"📊 File size: {pdf_size:,} bytes")
        print(f"📋 Sheets converted: {len(sheets_data)}")
        
    except KeyboardInterrupt:
        logger.info("Conversion cancelled by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Conversion failed: {e}")
        if args.verbose:
            logger.exception("Full error details:")
        sys.exit(1)


if __name__ == "__main__":
    main()
