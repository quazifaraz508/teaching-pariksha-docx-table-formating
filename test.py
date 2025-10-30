import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import io
import os

# Set page config
st.set_page_config(
    page_title="Teaching Pariksha Word Formatter",
    page_icon="📊",
    layout="centered"
)

def fix_table_background(doc):
    """
    Convert tables with black background to normal tables in a Word document
    """
    for table in doc.tables:
        # Remove any shading from the entire table
        for row in table.rows:
            for cell in row.cells:
                # Remove cell shading
                try:
                    tc_pr = cell._element.tcPr
                    shd = tc_pr.find(qn('w:shd'))
                    if shd is not None:
                        shd.set(qn('w:fill'), 'auto')
                        shd.set(qn('w:val'), 'clear')
                except:
                    pass
                
                # Also check and fix paragraph formatting in cells
                for paragraph in cell.paragraphs:
                    # Reset paragraph formatting
                    try:
                        paragraph.style = doc.styles['Normal']
                    except:
                        pass
                    
                    # Reset run formatting
                    for run in paragraph.runs:
                        try:
                            run.font.color.rgb = RGBColor(0, 0, 0)  # Set text to black
                            run.font.bold = False
                            run.font.italic = False
                        except:
                            pass
    return doc

def main():
    st.title("📊 Teaching Pariksha Word Formatter")
    st.markdown("Fix black background tables in Word documents")
    
    # File upload section
    st.header("1. Upload Your Word Document")
    uploaded_file = st.file_uploader(
        "Choose a Word document (.docx)", 
        type=['docx'],
        help="Upload a Word file with tables that have black backgrounds"
    )
    
    if uploaded_file is not None:
        # Display file info
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        
        # Preview section
        st.header("2. Process Document")
        
        if st.button("🚀 Fix Table Formatting", type="primary"):
            with st.spinner("Processing your document..."):
                try:
                    # Read the uploaded file
                    doc = Document(uploaded_file)
                    
                    # Fix tables
                    fixed_doc = fix_table_background(doc)
                    
                    # Save to bytes
                    output = io.BytesIO()
                    fixed_doc.save(output)
                    output.seek(0)
                    
                    # Success message
                    st.success("✅ Table formatting fixed successfully!")
                    
                    # Download section
                    st.header("3. Download Fixed Document")
                    
                    # Get original filename and create new one
                    original_name = uploaded_file.name
                    if original_name.endswith('.docx'):
                        new_name = original_name.replace('.docx', '_fixed.docx')
                    else:
                        new_name = original_name + '_fixed.docx'
                    
                    # Download button
                    st.download_button(
                        label="📥 Download Fixed Document",
                        data=output,
                        file_name=new_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        help="Click to download the formatted document"
                    )
                    
                    st.info("💡 The downloaded file will have normal table formatting with white backgrounds and black text.")
                    
                except Exception as e:
                    st.error(f"❌ Error processing document: {str(e)}")
                    st.info("Please make sure you uploaded a valid Word document (.docx file)")
    
    # Instructions section
    with st.expander("📖 How to Use"):
        st.markdown("""
        **Simple Steps:**
        1. **Upload** - Select your Word document with black background tables
        2. **Process** - Click the 'Fix Table Formatting' button
        3. **Download** - Save the fixed document with normal tables
        
        **What this tool fixes:**
        - Removes black/colored table backgrounds
        - Converts text to black color
        - Removes bold/italic formatting from table text
        - Makes tables readable and printable
        """)
    
    # Features section
    with st.expander("✨ Features"):
        st.markdown("""
        - ✅ Fixes multiple tables in one document
        - ✅ Preserves table structure and content
        - ✅ Converts black backgrounds to white
        - ✅ Sets text color to black for readability
        - ✅ Simple one-click operation
        - ✅ No installation required - works in browser
        """)

if __name__ == "__main__":

    main()


