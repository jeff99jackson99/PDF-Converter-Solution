"""Streamlit web application for Excel to PDF conversion."""

import streamlit as st
import pandas as pd
from pathlib import Path
import tempfile
import logging
from typing import Dict, List, Optional
import io

# Import our converter modules
from src.excel_pdf_converter.converter import ExcelToPDFConverter
from src.excel_pdf_converter.excel_reader import ExcelReader

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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
    st.markdown('<h1 class="main-header">📊 Excel to PDF Converter</h1>', unsafe_allow_html=True)
    st.markdown("""
    Convert your Excel proforma files to professional PDF documents. 
    Perfect for solving Excel printing issues with Assumptions, Proforma, Proforma Condensed, and Calculations sheets.
    """)
    
    # Sidebar for file upload and settings
    with st.sidebar:
        st.header("📁 File Upload")
        
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload your Excel file containing the proforma data"
        )
        
        st.header("⚙️ Conversion Settings")
        
        # Sheet selection
        st.subheader("Select Sheets")
        convert_proforma_only = st.checkbox(
            "Convert Proforma Sheets Only", 
            value=True,
            help="Automatically select Assumptions, Proforma, Proforma Condensed, and Calculations sheets"
        )
        
        # Display settings
        st.subheader("Display Settings")
        max_rows = st.slider(
            "Maximum rows per sheet",
            min_value=10,
            max_value=50,
            value=25,
            help="Limit the number of rows displayed to keep PDF manageable"
        )
        
        max_cols = st.slider(
            "Maximum columns per sheet",
            min_value=5,
            max_value=15,
            value=8,
            help="Limit the number of columns displayed"
        )
        
        include_summaries = st.checkbox(
            "Include sheet summaries",
            value=True,
            help="Add summary information for each sheet"
        )
    
    # Main content area
    if uploaded_file is not None:
        try:
            # Save uploaded file temporarily
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name
            
            # Initialize converter
            converter = ExcelToPDFConverter(tmp_file_path)
            
            # Get available sheets
            available_sheets = converter.get_available_sheets()
            
            # Display file information
            st.markdown('<h2 class="section-header">📋 File Information</h2>', unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.info(f"**File Name:** {uploaded_file.name}")
                st.info(f"**File Size:** {uploaded_file.size:,} bytes")
                st.info(f"**Total Sheets:** {len(available_sheets)}")
            
            with col2:
                st.info(f"**Available Sheets:**")
                for sheet in available_sheets:
                    st.write(f"• {sheet}")
            
            # Sheet selection
            if not convert_proforma_only:
                st.markdown('<h2 class="section-header">🎯 Sheet Selection</h2>', unsafe_allow_html=True)
                
                selected_sheets = st.multiselect(
                    "Choose sheets to convert:",
                    options=available_sheets,
                    default=available_sheets,
                    help="Select which sheets you want to include in the PDF"
                )
                
                if not selected_sheets:
                    st.warning("Please select at least one sheet to convert.")
                    return
            else:
                # Auto-select proforma sheets
                proforma_sheets = ['Assumptions', 'Proforma', 'Proforma Condensed', 'Calculations']
                selected_sheets = [sheet for sheet in proforma_sheets if sheet in available_sheets]
                
                if not selected_sheets:
                    st.error("No proforma sheets found in the uploaded file.")
                    st.info("Available sheets: " + ", ".join(available_sheets))
                    return
                
                st.success(f"Auto-selected proforma sheets: {', '.join(selected_sheets)}")
            
            # Sheet validation and preview
            st.markdown('<h2 class="section-header">🔍 Sheet Preview</h2>', unsafe_allow_html=True)
            
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
                                st.write("**Preview:**")
                                st.dataframe(df.head(10), use_container_width=True)
                            else:
                                st.warning("This sheet is empty.")
                                
                        except Exception as e:
                            st.error(f"Error reading sheet: {e}")
                    else:
                        st.error("Sheet not found or has no data.")
            
            # Conversion button
            st.markdown('<h2 class="section-header">🔄 Convert to PDF</h2>', unsafe_allow_html=True)
            
            if st.button("🚀 Generate PDF", type="primary", use_container_width=True):
                with st.spinner("Converting Excel to PDF..."):
                    try:
                        # Load selected sheets
                        if convert_proforma_only:
                            converter.load_proforma_sheets()
                        else:
                            converter.load_sheets(selected_sheets)
                        
                        # Generate PDF
                        pdf_path = converter.convert_to_pdf(
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
                            file_name=f"{Path(uploaded_file.name).stem}_converted.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                        
                        # Show file info
                        st.info(f"**PDF saved as:** {Path(pdf_path).name}")
                        st.info(f"**File size:** {len(pdf_bytes):,} bytes")
                        
                    except Exception as e:
                        logger.error(f"Conversion error: {e}")
                        st.markdown(f'<div class="error-message">❌ Error during conversion: {str(e)}</div>', 
                                  unsafe_allow_html=True)
            
            # Cleanup
            Path(tmp_file_path).unlink(missing_ok=True)
            
        except Exception as e:
            logger.error(f"Application error: {e}")
            st.markdown(f'<div class="error-message">❌ Error processing file: {str(e)}</div>', 
                      unsafe_allow_html=True)
    
    else:
        # Instructions when no file is uploaded
        st.markdown('<h2 class="section-header">📖 How to Use</h2>', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            **Step 1:** Upload your Excel file using the sidebar
            
            **Step 2:** Choose your conversion settings:
            - Select which sheets to convert
            - Adjust display limits (rows/columns)
            - Choose whether to include summaries
            
            **Step 3:** Click "Generate PDF" to convert
            
            **Step 4:** Download your professional PDF document
            """)
        
        with col2:
            st.markdown("""
            **Perfect for:**
            - ✅ Proforma financial models
            - ✅ Assumptions documentation
            - ✅ Calculation summaries
            - ✅ Professional reports
            
            **Solves Excel Issues:**
            - ❌ Sheets won't print to PDF
            - ❌ Inconsistent page setups
            - ❌ Hidden data problems
            - ❌ Print area issues
            """)
        
        # Example features
        st.markdown('<h2 class="section-header">✨ Features</h2>', unsafe_allow_html=True)
        
        features = [
            "🎯 Automatic proforma sheet detection",
            "📊 Professional PDF formatting",
            "🔍 Sheet validation and preview",
            "⚙️ Customizable display settings",
            "📱 Responsive web interface",
            "🚀 Fast conversion process",
            "💾 Direct download capability"
        ]
        
        cols = st.columns(2)
        for i, feature in enumerate(features):
            cols[i % 2].write(feature)

if __name__ == "__main__":
    main()
