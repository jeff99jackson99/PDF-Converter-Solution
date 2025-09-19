"""Streamlit web application for Excel to PDF conversion."""

import streamlit as st
import pandas as pd
from pathlib import Path
import logging
from typing import Dict, List, Optional
import os

# Import our converter modules
from src.excel_pdf_converter.converter import ExcelToPDFConverter
from src.excel_pdf_converter.excel_reader import ExcelReader

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Path to the Excel file in the project folder
EXCEL_FILE_PATH = "Pro Forma (4 Products).xlsx"

# Page configuration
st.set_page_config(
    page_title="Excel to PDF Converter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-header {
        font-size: 1.5rem;
        color: #2c3e50;
        margin-top: 2rem;
        margin-bottom: 1rem;
    }
    .success-message {
        background-color: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #c3e6cb;
    }
    .error-message {
        background-color: #f8d7da;
        color: #721c24;
        padding: 1rem;
        border-radius: 0.5rem;
        border: 1px solid #f5c6cb;
    }
</style>
""", unsafe_allow_html=True)

def main():
    """Main Streamlit application."""
    
    # Header
    st.markdown('<h1 class="main-header">📊 Proforma PDF Generator</h1>', unsafe_allow_html=True)
    st.markdown("""
    Generate professional PDF from your proforma Excel file. 
    Optimized for Assumptions, Proforma, Proforma Condensed, and Calculations sheets.
    """)
    
    # Check if Excel file exists
    if not os.path.exists(EXCEL_FILE_PATH):
        st.error(f"❌ Excel file not found: {EXCEL_FILE_PATH}")
        st.info("Please ensure 'Pro Forma (4 Products).xlsx' is in the project folder.")
        return
    
    # Show file info
    file_size = os.path.getsize(EXCEL_FILE_PATH)
    st.success(f"✅ Found Excel file: {EXCEL_FILE_PATH} ({file_size:,} bytes)")
    
    # Sidebar for settings
    with st.sidebar:
        st.header("⚙️ PDF Generation Settings")
        
        # Display settings
        st.subheader("Data Capture Settings")
        max_rows = st.slider(
            "Maximum rows per sheet",
            min_value=100,
            max_value=1000,
            value=500,
            help="Capture all your Excel data - set high to include everything"
        )
        
        max_cols = st.slider(
            "Maximum columns per sheet",
            min_value=20,
            max_value=50,
            value=30,
            help="Capture all columns from your Excel sheets"
        )
        
        include_summaries = st.checkbox(
            "Include sheet summaries",
            value=True,
            help="Add summary information for each sheet"
        )
        
        st.subheader("Output Settings")
        pdf_filename = st.text_input(
            "PDF filename",
            value="Proforma_Analysis.pdf",
            help="Name for the generated PDF file"
        )
    
    # Initialize converter with the project Excel file
    try:
        converter = ExcelToPDFConverter(EXCEL_FILE_PATH)
        
        # Get available sheets
        available_sheets = converter.get_available_sheets()
        
        # Display file information
        st.markdown('<h2 class="section-header">📋 Excel File Analysis</h2>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.info(f"**File Name:** {EXCEL_FILE_PATH}")
            st.info(f"**File Size:** {file_size:,} bytes")
            st.info(f"**Total Sheets:** {len(available_sheets)}")
        
        with col2:
            st.info(f"**Available Sheets:**")
            for sheet in available_sheets:
                st.write(f"• {sheet}")
        
        # Auto-select proforma sheets
        proforma_sheets = ['Assumptions', 'Proforma', 'Proforma Condensed', 'Calculations']
        selected_sheets = [sheet for sheet in proforma_sheets if sheet in available_sheets]
        
        if not selected_sheets:
            st.error("❌ No proforma sheets found in the Excel file.")
            st.info("Available sheets: " + ", ".join(available_sheets))
            return
        
        st.success(f"✅ Found proforma sheets: {', '.join(selected_sheets)}")
        
        # Sheet validation and preview
        st.markdown('<h2 class="section-header">🔍 Sheet Data Preview</h2>', unsafe_allow_html=True)
        
        validation_results = converter.validate_sheets(selected_sheets)
        
        for sheet_name in selected_sheets:
            with st.expander(f"📊 {sheet_name}", expanded=False):
                if validation_results.get(sheet_name, False):
                    try:
                        # Load sheet data
                        df = converter.excel_reader.read_sheet(sheet_name)
                        
                        st.write(f"**Dimensions:** {df.shape[0]} rows × {df.shape[1]} columns")
                        st.write(f"**Data Types:** {', '.join(df.dtypes.astype(str).unique())}")
                        
                        # Show preview
                        if not df.empty:
                            st.write("**Preview (first 10 rows):**")
                            st.dataframe(df.head(10), use_container_width=True)
                        else:
                            st.warning("This sheet is empty.")
                            
                    except Exception as e:
                        st.error(f"Error reading sheet: {e}")
                else:
                    st.error("Sheet not found or has no data.")
        
        # Conversion button
        st.markdown('<h2 class="section-header">🔄 Generate PDF</h2>', unsafe_allow_html=True)
        
        if st.button("🚀 Generate Professional PDF", type="primary", use_container_width=True):
            with st.spinner("Converting your proforma to PDF..."):
                try:
                    # Load proforma sheets
                    converter.load_proforma_sheets()
                    
                    # Generate PDF with custom filename
                    pdf_path = converter.convert_to_pdf(
                        pdf_filename=pdf_filename,
                        include_sheet_summaries=include_summaries,
                        max_rows_per_sheet=max_rows,
                        max_cols_per_sheet=max_cols
                    )
                    
                    # Read the generated PDF
                    with open(pdf_path, "rb") as pdf_file:
                        pdf_bytes = pdf_file.read()
                    
                    # Success message
                    st.markdown('<div class="success-message">✅ PDF generated successfully!</div>', 
                              unsafe_allow_html=True)
                    
                    # Download button
                    st.download_button(
                        label="📥 Download PDF",
                        data=pdf_bytes,
                        file_name=pdf_filename,
                        mime="application/pdf",
                        use_container_width=True
                    )
                    
                    # Show file info
                    st.info(f"**PDF generated:** {pdf_filename}")
                    st.info(f"**File size:** {len(pdf_bytes):,} bytes")
                    st.info(f"**Sheets included:** {len(selected_sheets)}")
                    
                except Exception as e:
                    logger.error(f"Conversion error: {e}")
                    st.markdown(f'<div class="error-message">❌ Error during conversion: {str(e)}</div>', 
                              unsafe_allow_html=True)
        
    except Exception as e:
        logger.error(f"Application error: {e}")
        st.markdown(f'<div class="error-message">❌ Error loading Excel file: {str(e)}</div>', 
                  unsafe_allow_html=True)

if __name__ == "__main__":
    main()
