"""Main converter class that orchestrates Excel to PDF conversion."""

import logging
from pathlib import Path
from typing import Dict, List, Optional
import pandas as pd

from .excel_reader import ExcelReader
from .pdf_generator import PDFGenerator

logger = logging.getLogger(__name__)


class ExcelToPDFConverter:
    """Main class that handles the complete Excel to PDF conversion process."""
    
    def __init__(self, excel_file_path: str, output_dir: str = "output") -> None:
        """Initialize the converter.
        
        Args:
            excel_file_path: Path to the Excel file to convert
            output_dir: Directory where PDF will be saved
        """
        self.excel_file_path = Path(excel_file_path)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)
        
        self.excel_reader = ExcelReader(str(self.excel_file_path))
        self.sheets_data = {}
        
        logger.info(f"Converter initialized for: {self.excel_file_path}")
        logger.info(f"Output directory: {self.output_dir}")
    
    def load_sheets(self, sheet_names: Optional[List[str]] = None) -> Dict[str, pd.DataFrame]:
        """Load sheets from the Excel file.
        
        Args:
            sheet_names: List of specific sheet names to load. If None, loads all sheets.
            
        Returns:
            Dictionary mapping sheet names to DataFrames
        """
        if sheet_names is None:
            sheet_names = self.excel_reader.get_sheet_names()
        
        self.sheets_data = self.excel_reader.read_multiple_sheets(sheet_names)
        logger.info(f"Loaded {len(self.sheets_data)} sheets")
        return self.sheets_data
    
    def load_proforma_sheets(self) -> Dict[str, pd.DataFrame]:
        """Load the specific proforma sheets that were having printing issues.
        
        Returns:
            Dictionary with proforma sheet data
        """
        self.sheets_data = self.excel_reader.read_proforma_sheets()
        logger.info(f"Loaded proforma sheets: {list(self.sheets_data.keys())}")
        return self.sheets_data
    
    def get_sheet_info(self) -> Dict[str, Dict[str, any]]:
        """Get information about all loaded sheets.
        
        Returns:
            Dictionary with sheet information
        """
        sheet_info = {}
        for sheet_name in self.sheets_data.keys():
            sheet_info[sheet_name] = self.excel_reader.get_sheet_info(sheet_name)
        return sheet_info
    
    def convert_to_pdf(self, pdf_filename: Optional[str] = None, 
                      include_sheet_summaries: bool = True,
                      max_rows_per_sheet: int = 1000,
                      max_cols_per_sheet: int = 50) -> str:
        """Convert loaded sheets to PDF.
        
        Args:
            pdf_filename: Name of the output PDF file. If None, uses Excel filename.
            include_sheet_summaries: Whether to include sheet summaries
            max_rows_per_sheet: Maximum rows to display per sheet
            max_cols_per_sheet: Maximum columns to display per sheet
            
        Returns:
            Path to the generated PDF file
        """
        if not self.sheets_data:
            raise ValueError("No sheets loaded. Call load_sheets() or load_proforma_sheets() first.")
        
        # Generate PDF filename if not provided
        if pdf_filename is None:
            pdf_filename = f"{self.excel_file_path.stem}_converted.pdf"
        
        pdf_path = self.output_dir / pdf_filename
        
        # Initialize PDF generator
        pdf_gen = PDFGenerator(str(pdf_path))
        
        # Add title page
        title = f"Proforma Analysis - {self.excel_file_path.stem}"
        pdf_gen.add_title_page(title)
        
        # Add each sheet
        sheet_names = list(self.sheets_data.keys())
        for i, (sheet_name, df) in enumerate(self.sheets_data.items()):
            try:
                # Add sheet summary if requested
                if include_sheet_summaries:
                    pdf_gen.add_sheet_summary(sheet_name, df)
                
                # Add sheet data
                pdf_gen.add_sheet_data(
                    sheet_name, 
                    df, 
                    max_rows=max_rows_per_sheet,
                    max_cols=max_cols_per_sheet
                )
                
                # Add page break except for last sheet
                if i < len(sheet_names) - 1:
                    pdf_gen.add_page_break()
                
                logger.info(f"Added sheet '{sheet_name}' to PDF")
                
            except Exception as e:
                logger.error(f"Error processing sheet '{sheet_name}': {e}")
                # Continue with other sheets
                continue
        
        # Add notes section
        notes = [
            f"Generated from: {self.excel_file_path.name}",
            f"Total sheets processed: {len(self.sheets_data)}",
            f"Maximum rows per sheet: {max_rows_per_sheet}",
            f"Maximum columns per sheet: {max_cols_per_sheet}"
        ]
        pdf_gen.add_notes_section(notes)
        
        # Generate PDF
        final_path = pdf_gen.generate_pdf()
        logger.info(f"PDF conversion completed: {final_path}")
        
        return final_path
    
    def convert_proforma_to_pdf(self, pdf_filename: Optional[str] = None) -> str:
        """Convert the specific proforma sheets to PDF.
        
        Args:
            pdf_filename: Name of the output PDF file
            
        Returns:
            Path to the generated PDF file
        """
        # Load proforma sheets
        self.load_proforma_sheets()
        
        # Set default filename if not provided
        if pdf_filename is None:
            pdf_filename = f"{self.excel_file_path.stem}_proforma.pdf"
        
        return self.convert_to_pdf(
            pdf_filename=pdf_filename,
            include_sheet_summaries=True,
            max_rows_per_sheet=1000,  # No practical limit
            max_cols_per_sheet=50     # No practical limit
        )
    
    def get_available_sheets(self) -> List[str]:
        """Get list of available sheet names in the Excel file.
        
        Returns:
            List of sheet names
        """
        return self.excel_reader.get_sheet_names()
    
    def validate_sheets(self, sheet_names: List[str]) -> Dict[str, bool]:
        """Validate that specified sheets exist and have data.
        
        Args:
            sheet_names: List of sheet names to validate
            
        Returns:
            Dictionary mapping sheet names to validation status
        """
        available_sheets = self.get_available_sheets()
        validation_results = {}
        
        for sheet_name in sheet_names:
            exists = sheet_name in available_sheets
            has_data = False
            
            if exists:
                try:
                    df = self.excel_reader.read_sheet(sheet_name)
                    has_data = not df.empty
                except Exception as e:
                    logger.error(f"Error reading sheet {sheet_name}: {e}")
                    has_data = False
            
            validation_results[sheet_name] = exists and has_data
        
        return validation_results
