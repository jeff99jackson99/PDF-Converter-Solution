"""PDF generator for converting Excel data to professional PDF format."""

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.platypus.flowables import KeepTogether
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import pandas as pd
from typing import Dict, List, Optional, Tuple, Any
import logging
from datetime import datetime

logger = logging.getLogger(__name__)


class PDFGenerator:
    """Generates professional PDF documents from Excel data."""
    
    def __init__(self, output_path: str, page_size: Tuple[int, int] = None) -> None:
        """Initialize PDF generator.
        
        Args:
            output_path: Path where PDF will be saved
            page_size: Page size tuple (width, height) in points. If None, auto-detect based on content.
        """
        self.output_path = output_path
        self.page_size = page_size or landscape(letter)  # Default to landscape for better table display
        self.doc = SimpleDocTemplate(
            output_path,
            pagesize=self.page_size,
            rightMargin=0.3*inch,  # Smaller margins for more space
            leftMargin=0.3*inch,
            topMargin=0.5*inch,
            bottomMargin=0.3*inch
        )
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()
        self.story = []
    
    def _setup_custom_styles(self) -> None:
        """Set up custom paragraph styles for the PDF."""
        # Title style
        self.title_style = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            spaceAfter=12,
            alignment=TA_CENTER,
            textColor=colors.darkblue
        )
        
        # Sheet header style
        self.sheet_header_style = ParagraphStyle(
            'SheetHeader',
            parent=self.styles['Heading2'],
            fontSize=14,
            spaceAfter=8,
            spaceBefore=12,
            textColor=colors.darkblue,
            borderWidth=1,
            borderColor=colors.grey,
            borderPadding=6
        )
        
        # Data cell style
        self.cell_style = ParagraphStyle(
            'CellStyle',
            parent=self.styles['Normal'],
            fontSize=8,
            alignment=TA_LEFT
        )
        
        # Number cell style
        self.number_style = ParagraphStyle(
            'NumberStyle',
            parent=self.styles['Normal'],
            fontSize=8,
            alignment=TA_RIGHT
        )
    
    def add_title_page(self, title: str = "Proforma Financial Analysis") -> None:
        """Add a title page to the PDF.
        
        Args:
            title: Title of the document
        """
        title_para = Paragraph(title, self.title_style)
        self.story.append(title_para)
        self.story.append(Spacer(1, 0.3*inch))
        
        # Add generation timestamp
        timestamp = datetime.now().strftime("%B %d, %Y at %I:%M %p")
        timestamp_para = Paragraph(f"Generated on: {timestamp}", self.styles['Normal'])
        self.story.append(timestamp_para)
        self.story.append(Spacer(1, 0.5*inch))
    
    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean and prepare DataFrame for PDF display.
        
        Args:
            df: Input DataFrame
            
        Returns:
            Cleaned DataFrame
        """
        # Replace NaN values with empty strings
        df_clean = df.fillna('')
        
        # Convert all values to strings for consistent formatting
        df_clean = df_clean.astype(str)
        
        # Clean up cell values
        for col in df_clean.columns:
            df_clean[col] = df_clean[col].apply(self._clean_cell_value)
        
        return df_clean
    
    def _clean_cell_value(self, value: str) -> str:
        """Clean individual cell values for better PDF display.
        
        Args:
            value: Cell value as string
            
        Returns:
            Cleaned cell value
        """
        if not value or value == 'nan':
            return ''
        
        # Remove extra whitespace
        value = str(value).strip()
        
        # Handle numbers with proper formatting
        try:
            # Try to format as number if it's numeric
            if '.' in value:
                num = float(value)
                if num == int(num):
                    return str(int(num))
                return f"{num:.2f}"
            else:
                num = int(value)
                return str(num)
        except (ValueError, TypeError):
            # Not a number, return as string with length limit
            return value[:40] + '...' if len(value) > 40 else value
    
    def _create_table_from_dataframe(self, df: pd.DataFrame, max_rows: int = 100) -> Table:
        """Create a ReportLab Table from a pandas DataFrame.
        
        Args:
            df: Input DataFrame
            max_rows: Maximum number of rows to display
            
        Returns:
            ReportLab Table object
        """
        df_clean = self._clean_dataframe(df)
        
        # Log original size
        original_rows = len(df_clean)
        logger.info(f"Creating table from {original_rows} rows, {len(df_clean.columns)} columns")
        
        # Limit rows if DataFrame is too large
        if len(df_clean) > max_rows:
            df_clean = df_clean.head(max_rows)
            logger.warning(f"DataFrame truncated to {max_rows} rows for PDF display (was {original_rows} rows)")
        
        # Prepare table data
        table_data = []
        
        # Add header row
        if not df_clean.empty:
            headers = [str(col) for col in df_clean.columns]
            table_data.append(headers)
            
            # Add data rows
            for _, row in df_clean.iterrows():
                row_data = [str(cell) for cell in row.values]
                table_data.append(row_data)
        
        # Create table with repeatRows for headers
        table = Table(table_data, repeatRows=1)
        
        # Calculate column widths based on content
        col_widths = self._calculate_column_widths(table_data)
        
        # Style the table
        table_style = TableStyle([
            # Header row styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            
            # Data rows styling
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
            
            # Borders
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            
            # Alternating row colors
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            
            # Wrap text in cells
            ('WORDWRAP', (0, 0), (-1, -1), 'CJK'),
        ])
        
        table.setStyle(table_style)
        
        # Apply column widths if calculated
        if col_widths:
            table._colWidths = col_widths
        
        return table
    
    def _calculate_column_widths(self, table_data: List[List[str]]) -> List[float]:
        """Calculate appropriate column widths based on content.
        
        Args:
            table_data: Table data as list of rows
            
        Returns:
            List of column widths in points
        """
        if not table_data:
            return []
        
        num_cols = len(table_data[0])
        max_widths = [0] * num_cols
        
        # Find maximum width for each column
        for row in table_data:
            for i, cell in enumerate(row):
                if i < num_cols:
                    cell_width = len(str(cell)) * 6  # Approximate character width
                    max_widths[i] = max(max_widths[i], cell_width)
        
        # Set reasonable limits for landscape mode
        min_width = 80  # Minimum column width
        max_width = 200  # Maximum column width
        
        # Apply limits and convert to points
        col_widths = []
        for width in max_widths:
            width = max(min_width, min(width, max_width))
            col_widths.append(width)
        
        return col_widths
    
    def add_sheet_data(self, sheet_name: str, df: pd.DataFrame, 
                      max_rows: int = 30, max_cols: int = 10) -> None:
        """Add a sheet's data to the PDF.
        
        Args:
            sheet_name: Name of the sheet
            df: DataFrame containing the sheet data
            max_rows: Maximum number of rows to display
            max_cols: Maximum number of columns to display
        """
        # Add sheet header
        sheet_header = Paragraph(f"<b>{sheet_name}</b>", self.sheet_header_style)
        self.story.append(sheet_header)
        
        # Check if DataFrame is empty
        if df.empty:
            empty_msg = Paragraph("No data available in this sheet.", self.styles['Normal'])
            self.story.append(empty_msg)
            self.story.append(Spacer(1, 0.2*inch))
            return
        
        # Limit columns if necessary
        if len(df.columns) > max_cols:
            df = df.iloc[:, :max_cols]
            logger.warning(f"DataFrame truncated to {max_cols} columns for PDF display")
        
        # Create and add table
        try:
            table = self._create_table_from_dataframe(df, max_rows)
            self.story.append(table)
            self.story.append(Spacer(1, 0.2*inch))
        except Exception as e:
            logger.error(f"Error creating table for sheet {sheet_name}: {e}")
            error_msg = Paragraph(f"Error displaying data for {sheet_name}: {str(e)}", 
                                self.styles['Normal'])
            self.story.append(error_msg)
    
    def add_sheet_summary(self, sheet_name: str, df: pd.DataFrame) -> None:
        """Add a summary of the sheet data.
        
        Args:
            sheet_name: Name of the sheet
            df: DataFrame containing the sheet data
        """
        if df.empty:
            return
        
        summary_text = f"""
        <b>{sheet_name} Summary:</b><br/>
        • Rows: {len(df)}<br/>
        • Columns: {len(df.columns)}<br/>
        • Data Types: {', '.join(df.dtypes.astype(str).unique())}<br/>
        """
        
        summary_para = Paragraph(summary_text, self.styles['Normal'])
        self.story.append(summary_para)
        self.story.append(Spacer(1, 0.1*inch))
    
    def add_page_break(self) -> None:
        """Add a page break to the PDF."""
        self.story.append(PageBreak())
    
    def generate_pdf(self) -> str:
        """Generate the final PDF document.
        
        Returns:
            Path to the generated PDF file
        """
        try:
            self.doc.build(self.story)
            logger.info(f"PDF successfully generated: {self.output_path}")
            return self.output_path
        except Exception as e:
            logger.error(f"Failed to generate PDF: {e}")
            raise
    
    def add_notes_section(self, notes: List[str]) -> None:
        """Add a notes section to the PDF.
        
        Args:
            notes: List of note strings
        """
        if not notes:
            return
        
        notes_header = Paragraph("<b>Notes:</b>", self.styles['Heading3'])
        self.story.append(notes_header)
        
        for note in notes:
            note_para = Paragraph(f"• {note}", self.styles['Normal'])
            self.story.append(note_para)
        
        self.story.append(Spacer(1, 0.2*inch))
