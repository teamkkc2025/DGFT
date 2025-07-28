import streamlit as st
import PyPDF2
import re
import json
from io import BytesIO
import pandas as pd
from datetime import datetime
import zipfile
import os
 
# Try to import pdfplumber for better text extraction
try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False
    st.warning("⚠️ pdfplumber not available. Install with: pip install pdfplumber for better text extraction.")
 
# Try to import openpyxl for Excel support
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    st.error("❌ openpyxl not available. Install with: pip install openpyxl for Excel support.")
 
def extract_text_from_pdf(pdf_file):
    """Extract text from uploaded PDF file with multiple methods"""
    text = ""
   
    # Method 1: Try pdfplumber first (usually better)
    if PDFPLUMBER_AVAILABLE:
        try:
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
           
            if len(text.strip()) > 100:  # If we got good text, return it
                return text, "pdfplumber"
        except Exception as e:
            st.warning(f"pdfplumber extraction failed for {pdf_file.name}: {str(e)}, trying PyPDF2...")
   
    # Method 2: Fallback to PyPDF2
    try:
        pdf_file.seek(0)  # Reset file pointer
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = ""
       
        for page in pdf_reader.pages:
            page_text = page.extract_text()
            text += page_text + "\n"
       
        # Clean up the text
        text = text.replace('\x00', '')  # Remove null characters
        text = re.sub(r'\n\s*\n', '\n', text)  # Remove excessive newlines
       
        if len(text.strip()) > 50:
            return text, "PyPDF2"
        else:
            return text, "PyPDF2 (limited)"
           
    except Exception as e:
        st.error(f"Error reading PDF {pdf_file.name} with PyPDF2: {str(e)}")
        return None, "Error"
 
def extract_brc_data(text):
    """Extract BRC data from the text using improved regex patterns"""
    data = {}
   
    # Clean the text - remove extra spaces and normalize
    text = re.sub(r'\s+', ' ', text.strip())
   
    # More flexible extraction patterns that handle various formats
    patterns = {
        "Firm Name": [
            r"1\s+Firm Name\s+([^\n\r2]+?)(?=\s*2\s+Address)",
            r"Firm Name\s+([A-Z\s&.,-]+?)(?=\s*(?:Address|2\s))",
            r"1\s*Firm Name\s*([A-Z\s&.,-]+)"
        ],
        "Address/GSTIN": [
            r"2\s+Address/GSTIN\s+(.+?)(?=\s*3\s+IEC)",
            r"Address/GSTIN\s+([^3]+?)(?=\s*3\s+IEC)",
            r"2\s*Address/GSTIN\s*([A-Z0-9\s,.-/&]+?)(?=\s*3)"
        ],
        "IEC": [
            r"3\s+IEC\s+(\d+)",
            r"IEC\s+(\d{10})",
            r"3\s*IEC\s*(\d+)"
        ],
        "Shipping Bill / Invoice No.": [
            r"4\s+Shipping Bill / Invoice No\.\s+(\d+)",
            r"Shipping Bill / Invoice No\.\s+(\d+)",
            r"4\s*Shipping Bill.*?No\.?\s*(\d+)"
        ],
        "Shipping Bill / Invoice Date": [
            r"5\s+Shipping Bill / Invoice Date\s+([\d-]+)",
            r"Shipping Bill / Invoice Date\s+([\d-]+)",
            r"5\s*Shipping Bill.*?Date\s*([\d-]+)"
        ],
        "Shipping Bill Port": [
            r"6\s+Shipping Bill Port\s+([A-Z0-9]+)",
            r"Shipping Bill Port\s+([A-Z0-9]+)",
            r"6\s*Shipping Bill Port\s*([A-Z0-9]+)"
        ],
        "Bank Name": [
            r"7\s+Bank Name\s+([A-Z\s&]+?)(?=\s*8\s+Bill)",
            r"Bank Name\s+([A-Z\s&]+?)(?=\s*(?:Bill|8\s))",
            r"7\s*Bank Name\s*([A-Z\s&]+)"
        ],
        "Bill ID No.": [
            r"8\s+Bill ID No\.\s+([A-Z0-9]+)",
            r"Bill ID No\.\s+([A-Z0-9]+)",
            r"8\s*Bill ID No\.?\s*([A-Z0-9]+)"
        ],
        "Bank Realisation Certificate No.": [
            r"(?:9\s+)?Bank\s+Realisation\s+Certificate\s+No\.\s+([A-Z0-9]+)\s+Dated\s+([\d-]+)",
            r"Certificate\s+No\.\s+([A-Z0-9]+)\s+Dated\s+([\d-]+)",
            r"([A-Z0-9]+)\s+Dated\s+([\d-]+)"
        ],
        "Date of Realisation of Money by Bank": [
            r"10\s+Date\s+of\s+Realisation\s+of\s+Money\s+by\s+Bank\s+([\d-]+)",
            r"Date\s+of\s+Realisation\s+of\s+Money\s+by\s+Bank\s+([\d-]+)",
            r"Realisation\s+of\s+Money\s+by\s+Bank\s+([\d-]+)",
            r"by\s+Bank\s+([\d-]+)(?=\s+11)",
            r"Money\s+by\s+Bank\s+([\d-]+)"
        ],
        "Total Realised Value": [
            r"11\s+Total Realised Value\s+([\d,\.]+)",
            r"Total Realised Value\s+([\d,\.]+)",
            r"11\s*Total.*?Value\s*([\d,\.]+)"
        ],
        "Net Realised Value": [
            r"13\s+Net Realised Value\s+([\d,\.]+)",
            r"Net Realised Value\s+([\d,\.]+)",
            r"13\s*Net.*?Value\s*([\d,\.]+)"
        ],
        "Currency of Realization": [
            r"14\s+Currency of Realization\s+([A-Z]+)",
            r"Currency of Realization\s+([A-Z]+)",
            r"14\s*Currency.*?Realization\s*([A-Z]+)"
        ],
        "Date and Time of Printing": [
            r"15\s+Date and Time of Printing\s+([\d-]+\s+[\d:]+\s+[AP]M)",
            r"Date and Time of Printing\s+([\d-]+\s+[\d:]+\s+[AP]M)",
            r"15\s*Date and Time.*?Printing\s*([\d-]+\s+[\d:]+\s+[AP]M)"
        ],
        "Source": [
            r"17\s+Source \(Bank /\s+Exporter\)\s+([A-Za-z]+)",
            r"Source.*?\(Bank.*?Exporter\)\s*([A-Za-z]+)",
            r"17\s*Source.*?([A-Za-z]+)$"
        ]
    }
   
    # Extract deductions with multiple pattern attempts
    deductions_patterns = [
        r"Commission\s+Discount\s+Insurance\s+Freight\s+Other\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)",
        r"Commission\s*Discount\s*Insurance\s*Freight\s*Other.*?([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)",
        r"12.*?Commission.*?Discount.*?Insurance.*?Freight.*?Other.*?([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)\s+([\d\.]+)"
    ]
   
    deductions_found = False
    for pattern in deductions_patterns:
        deductions_match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if deductions_match:
            data["Commission"] = deductions_match.group(1)
            data["Discount"] = deductions_match.group(2)
            data["Insurance"] = deductions_match.group(3)
            data["Freight"] = deductions_match.group(4)
            data["Other Deductions"] = deductions_match.group(5)
            deductions_found = True
            break
   
    if not deductions_found:
        # Default values if not found
        data["Commission"] = "0.00"
        data["Discount"] = "0.00"
        data["Insurance"] = "0.00"
        data["Freight"] = "0.00"
        data["Other Deductions"] = "0.00"
   
    # Extract other fields using multiple patterns
    for field, pattern_list in patterns.items():
        found = False
        for pattern in pattern_list:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                if field == "Bank Realisation Certificate No." and match.lastindex >= 2:
                    # Special handling for certificate number with date
                    extracted_value = f"{match.group(1)} Dated {match.group(2)}"
                else:
                    extracted_value = match.group(1).strip()
                # Clean up the extracted value
                extracted_value = re.sub(r'\s+', ' ', extracted_value)
                data[field] = extracted_value
                found = True
                break
       
        if not found:
            data[field] = ""
   
    return data
 
def validate_extracted_data(data):
    """Validate the extracted data and highlight missing fields"""
    required_fields = [
        "Firm Name", "Address/GSTIN", "IEC", "Shipping Bill / Invoice No.",
        "Shipping Bill / Invoice Date", "Bank Name", "Total Realised Value",
        "Net Realised Value", "Currency of Realization"
    ]
   
    missing_fields = []
    for field in required_fields:
        if not data.get(field) or data[field].strip() == "":
            missing_fields.append(field)
   
    return missing_fields
 
def process_multiple_files(uploaded_files):
    """Process multiple PDF files and return consolidated data"""
    all_extracted_data = []
    processing_summary = []
   
    progress_bar = st.progress(0)
    status_text = st.empty()
   
    for i, uploaded_file in enumerate(uploaded_files):
        status_text.text(f"Processing {uploaded_file.name}...")
        progress_bar.progress((i + 1) / len(uploaded_files))
       
        # Extract text from PDF
        pdf_text, extraction_method = extract_text_from_pdf(uploaded_file)
       
        if pdf_text:
            # Extract structured data
            extracted_data = extract_brc_data(pdf_text)
           
            # Add metadata
            extracted_data["File Name"] = uploaded_file.name
            extracted_data["Extraction Method"] = extraction_method
            extracted_data["Processing Date"] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
           
            # Validate data
            missing_fields = validate_extracted_data(extracted_data)
            extracted_data["Missing Fields Count"] = len(missing_fields)
            extracted_data["Missing Fields"] = ", ".join(missing_fields) if missing_fields else "None"
           
            all_extracted_data.append(extracted_data)
           
            # Add to summary
            processing_summary.append({
                "File Name": uploaded_file.name,
                "Status": "Success" if len(missing_fields) < 5 else "Partial",
                "Missing Fields": len(missing_fields),
                "Extraction Method": extraction_method
            })
        else:
            # Handle failed extraction
            failed_data = {
                "File Name": uploaded_file.name,
                "Extraction Method": "Failed",
                "Processing Date": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "Missing Fields Count": "All",
                "Missing Fields": "File processing failed"
            }
            all_extracted_data.append(failed_data)
           
            processing_summary.append({
                "File Name": uploaded_file.name,
                "Status": "Failed",
                "Missing Fields": "All",
                "Extraction Method": "Failed"
            })
   
    status_text.text("Processing complete!")
    progress_bar.progress(1.0)
   
    return all_extracted_data, processing_summary
 
def create_consolidated_excel(all_data):
    """Create consolidated Excel file with multiple sheets"""
    excel_buffer = BytesIO()
   
    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
        # Main data sheet
        df_main = pd.DataFrame(all_data)
        df_main.to_excel(writer, sheet_name='All_BRC_Data', index=False)
       
        # Summary sheet
        summary_data = []
        for data in all_data:
            summary_data.append({
                "File Name": data.get("File Name", ""),
                "Firm Name": data.get("Firm Name", ""),
                "IEC": data.get("IEC", ""),
                "Total Realised Value": data.get("Total Realised Value", ""),
                "Currency": data.get("Currency of Realization", ""),
                "Missing Fields": data.get("Missing Fields Count", ""),
                "Status": "Complete" if data.get("Missing Fields Count", 0) == 0 else "Partial"
            })
       
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name='Summary', index=False)
       
        # Statistics sheet
        total_files = len(all_data)
        successful_files = len([d for d in all_data if d.get("Missing Fields Count", 0) == 0])
        partial_files = len([d for d in all_data if isinstance(d.get("Missing Fields Count", 0), int) and d.get("Missing Fields Count", 0) > 0])
        failed_files = total_files - successful_files - partial_files
       
        stats_data = [
            ["Total Files Processed", total_files],
            ["Fully Successful", successful_files],
            ["Partially Successful", partial_files],
            ["Failed", failed_files],
            ["Success Rate", f"{(successful_files/total_files)*100:.1f}%" if total_files > 0 else "0%"],
            ["Processing Date", datetime.now().strftime('%Y-%m-%d %H:%M:%S')]
        ]
       
        df_stats = pd.DataFrame(stats_data, columns=['Metric', 'Value'])
        df_stats.to_excel(writer, sheet_name='Statistics', index=False)
   
    excel_buffer.seek(0)
    return excel_buffer
 
def main():
    st.set_page_config(
        page_title="DGFT BRC Data Extractor",
        page_icon="📋",
        layout="wide"
    )
   
    st.title("📋 DGFT Bank Realisation Certificate Data Extractor")
    st.markdown("Upload multiple DGFT Bank Realisation Certificate PDFs to extract structured data")
   
    # Sidebar for instructions
    with st.sidebar:
        st.header("⚙️ Requirements")
        st.markdown("""
        **Required packages:**
        ```bash
        pip install streamlit PyPDF2 pandas pdfplumber openpyxl
        ```
       
        **Note:**
        - pdfplumber provides better text extraction
        - openpyxl is required for Excel file downloads
        """)
       
        st.header("📖 Instructions")
        st.markdown("""
        1. Upload multiple DGFT Bank Realisation Certificate PDFs
        2. The tool will process all files automatically
        3. Review the consolidated results
        4. Download results as Excel file with multiple sheets
       
        **Features:**
        - Batch processing of multiple files
        - Consolidated Excel output
        - Processing summary and statistics
        - Individual file status tracking
        """)
       
        st.header("📊 Output Sheets")
        st.markdown("""
        **Excel file contains:**
        - **All_BRC_Data**: Complete extracted data
        - **Summary**: Overview of each file
        - **Statistics**: Processing statistics
        """)
   
    # File upload - now accepts multiple files
    uploaded_files = st.file_uploader(
        "Choose PDF files",
        type="pdf",
        accept_multiple_files=True,
        help="Upload multiple DGFT Bank Realisation Certificate PDF files"
    )
   
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully!")
       
        # Show uploaded files
        with st.expander("📁 Uploaded Files", expanded=True):
            for i, file in enumerate(uploaded_files, 1):
                st.write(f"{i}. {file.name} ({file.size} bytes)")
       
        # Process files button
        if st.button("🚀 Process All Files", type="primary"):
            with st.spinner("Processing all files..."):
                all_extracted_data, processing_summary = process_multiple_files(uploaded_files)
           
            # Display processing summary
            st.header("📊 Processing Summary")
            df_summary = pd.DataFrame(processing_summary)
            st.dataframe(df_summary, use_container_width=True, hide_index=True)
           
            # Show statistics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Files", len(uploaded_files))
            with col2:
                successful = len([s for s in processing_summary if s["Status"] == "Success"])
                st.metric("Successful", successful)
            with col3:
                partial = len([s for s in processing_summary if s["Status"] == "Partial"])
                st.metric("Partial", partial)
            with col4:
                failed = len([s for s in processing_summary if s["Status"] == "Failed"])
                st.metric("Failed", failed)
           
            # Display consolidated results
            st.header("📋 Consolidated Extracted Data")
            df_consolidated = pd.DataFrame(all_extracted_data)
            st.dataframe(df_consolidated, use_container_width=True, hide_index=True)
           
            # Download section
            st.header("💾 Download Results")
           
            if OPENPYXL_AVAILABLE and all_extracted_data:
                # Create consolidated Excel file
                excel_buffer = create_consolidated_excel(all_extracted_data)
               
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"consolidated_brc_data_{timestamp}.xlsx"
               
                st.download_button(
                    label="📥 Download Consolidated Excel File",
                    data=excel_buffer.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Downloads Excel file with multiple sheets: All Data, Summary, and Statistics"
                )
               
                # Also provide JSON download option
                json_data = json.dumps(all_extracted_data, indent=2, default=str)
                st.download_button(
                    label="📥 Download JSON File",
                    data=json_data,
                    file_name=f"consolidated_brc_data_{timestamp}.json",
                    mime="application/json"
                )
               
                # Individual CSV download
                csv_data = df_consolidated.to_csv(index=False)
                st.download_button(
                    label="📥 Download CSV File",
                    data=csv_data,
                    file_name=f"consolidated_brc_data_{timestamp}.csv",
                    mime="text/csv"
                )
           
            # Show detailed results for individual files
            if st.checkbox("📄 Show Individual File Details"):
                for i, data in enumerate(all_extracted_data):
                    with st.expander(f"File {i+1}: {data.get('File Name', 'Unknown')}", expanded=False):
                        # Create individual file DataFrame
                        file_data = {k: v for k, v in data.items() if k not in ['File Name', 'Processing Date', 'Extraction Method']}
                        df_individual = pd.DataFrame([file_data])
                        st.dataframe(df_individual, use_container_width=True, hide_index=True)
                       
                        # Show missing fields if any
                        missing_fields = data.get('Missing Fields', '')
                        if missing_fields and missing_fields != 'None':
                            st.warning(f"Missing fields: {missing_fields}")
 
if __name__ == "__main__":
    main()
 
 
