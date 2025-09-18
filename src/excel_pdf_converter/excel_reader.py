"""Excel file reader for proforma documents."""

import pandas as pd
from typing import Dict, List, Optional, Tuple
import logging

logger = logging.getLogger(__name__)


class ExcelReader:
    """Reads Excel files and extracts data from specific sheets."""
    
    def __init__(self, file_path: str) -> None:
        """Initialize with Excel file path.
        
        Args:
            file_path: Path to the Excel file
        """
        self.file_path = file_path
        self.workbook = None
        self._load_workbook()
    
    def _load_workbook(self) -> None:
        """Load the Excel workbook."""
        try:
            self.workbook = pd.ExcelFile(self.file_path)
            logger.info(f"Successfully loaded Excel file: {self.file_path}")
        except Exception as e:
            logger.error(f"Failed to load Excel file {self.file_path}: {e}")
            raise
    
    def get_sheet_names(self) -> List[str]:
        """Get all sheet names from the workbook.
        
        Returns:
            List of sheet names
        """
        return self.workbook.sheet_names
    
    def read_sheet(self, sheet_name: str, header_row: Optional[int] = None) -> pd.DataFrame:
        """Read a specific sheet from the workbook.
        
        Args:
            sheet_name: Name of the sheet to read
            header_row: Row number to use as header (0-indexed). If None, auto-detect.
            
        Returns:
            DataFrame containing the sheet data
        """
        try:
            # First, read without headers to inspect the structure
            df_raw = pd.read_excel(
                self.file_path,
                sheet_name=sheet_name,
                header=None,
                engine='openpyxl'
            )
            
            # Auto-detect header row if not specified
            if header_row is None:
                header_row = self._find_header_row(df_raw)
            
            # Read with proper headers - read ALL data
            df = pd.read_excel(
                self.file_path,
                sheet_name=sheet_name,
                header=header_row,
                engine='openpyxl',
                # Don't skip any rows or columns
                skiprows=None,
                skipcols=None,
                # Read all data
                nrows=None,
                usecols=None
            )
            
            # Clean up column names
            df = self._clean_dataframe_headers(df)
            
            logger.info(f"Successfully read sheet: {sheet_name} ({len(df)} rows, {len(df.columns)} columns)")
            return df
        except Exception as e:
            logger.error(f"Failed to read sheet {sheet_name}: {e}")
            raise
    
    def _find_header_row(self, df: pd.DataFrame) -> int:
        """Find the best row to use as headers by looking for the row with the most text content.
        
        Args:
            df: Raw DataFrame without headers
            
        Returns:
            Row index to use as header
        """
        best_row = 0
        max_text_cells = 0
        
        # Check first 10 rows
        for i in range(min(10, len(df))):
            row = df.iloc[i]
            text_cells = sum(1 for cell in row if pd.notna(cell) and str(cell).strip())
            
            if text_cells > max_text_cells:
                max_text_cells = text_cells
                best_row = i
        
        return best_row
    
    def _clean_dataframe_headers(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean up DataFrame column headers and remove empty rows/columns.
        
        Args:
            df: Input DataFrame
            
        Returns:
            Cleaned DataFrame
        """
        # Clean column names
        df.columns = [str(col).strip() if pd.notna(col) else f"Column_{i}" 
                     for i, col in enumerate(df.columns)]
        
        # Remove completely empty rows and columns
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # Reset index
        df = df.reset_index(drop=True)
        
        return df
    
    def read_multiple_sheets(self, sheet_names: List[str]) -> Dict[str, pd.DataFrame]:
        """Read multiple sheets at once.
        
        Args:
            sheet_names: List of sheet names to read
            
        Returns:
            Dictionary mapping sheet names to DataFrames
        """
        sheets_data = {}
        for sheet_name in sheet_names:
            sheets_data[sheet_name] = self.read_sheet(sheet_name)
        return sheets_data
    
    def get_sheet_info(self, sheet_name: str) -> Dict[str, any]:
        """Get information about a specific sheet.
        
        Args:
            sheet_name: Name of the sheet
            
        Returns:
            Dictionary with sheet information
        """
        try:
            df = self.read_sheet(sheet_name)
            return {
                'name': sheet_name,
                'rows': len(df),
                'columns': len(df.columns),
                'column_names': df.columns.tolist(),
                'has_data': not df.empty,
                'data_types': df.dtypes.to_dict()
            }
        except Exception as e:
            logger.error(f"Failed to get info for sheet {sheet_name}: {e}")
            return {'name': sheet_name, 'error': str(e)}
    
    def find_data_range(self, sheet_name: str) -> Tuple[int, int, int, int]:
        """Find the actual data range in a sheet (excluding empty rows/columns).
        
        Args:
            sheet_name: Name of the sheet
            
        Returns:
            Tuple of (start_row, end_row, start_col, end_col)
        """
        df = self.read_sheet(sheet_name)
        
        # Remove completely empty rows and columns
        df_clean = df.dropna(how='all').dropna(axis=1, how='all')
        
        if df_clean.empty:
            return (0, 0, 0, 0)
        
        # Find non-empty cells
        mask = df_clean.notna()
        
        # Get the bounds of non-empty data
        rows_with_data = mask.any(axis=1)
        cols_with_data = mask.any(axis=0)
        
        start_row = rows_with_data.idxmax() if rows_with_data.any() else 0
        end_row = rows_with_data[::-1].idxmax() if rows_with_data.any() else 0
        start_col = cols_with_data.idxmax() if cols_with_data.any() else 0
        end_col = cols_with_data[::-1].idxmax() if cols_with_data.any() else 0
        
        return (start_row, end_row, start_col, end_col)
    
    def read_proforma_sheets(self) -> Dict[str, pd.DataFrame]:
        """Read the specific proforma sheets that are having printing issues.
        
        Returns:
            Dictionary with the four key sheets
        """
        target_sheets = [
            'Assumptions',
            'Proforma', 
            'Proforma Condensed',
            'Calculations'
        ]
        
        available_sheets = self.get_sheet_names()
        found_sheets = [sheet for sheet in target_sheets if sheet in available_sheets]
        
        if not found_sheets:
            logger.warning("None of the target proforma sheets found. Available sheets:")
            logger.warning(f"Available sheets: {available_sheets}")
            return {}
        
        logger.info(f"Found proforma sheets: {found_sheets}")
        return self.read_multiple_sheets(found_sheets)
