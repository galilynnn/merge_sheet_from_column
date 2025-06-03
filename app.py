import streamlit as st
import pandas as pd
import io
from collections import defaultdict
import tempfile
import os

# Page configuration
st.set_page_config(
    page_title="Excel Sheet Merger",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state
if 'merged_data' not in st.session_state:
    st.session_state.merged_data = defaultdict(list)
if 'processed_files' not in st.session_state:
    st.session_state.processed_files = set()
if 'processing_log' not in st.session_state:
    st.session_state.processing_log = []

def normalize_sheet_name(name):
    """Normalize sheet name for case and space insensitive comparison"""
    return str(name).strip().lower().replace(' ', '').replace('_', '')

def log_message(message, msg_type="info"):
    """Add message to processing log"""
    st.session_state.processing_log.append({
        'message': message,
        'type': msg_type
    })

def process_excel_file(file, target_sheets):
    """Process a single Excel file and extract matching sheets"""
    results = defaultdict(list)
    file_log = []
    
    try:
        # Create temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(file.getvalue())
            tmp_file_path = tmp_file.name
        
        # Read Excel file
        xls = pd.ExcelFile(tmp_file_path)
        file_log.append(f"📁 Processing file: {file.name}")
        file_log.append(f"📑 Found {len(xls.sheet_names)} sheets: {', '.join(xls.sheet_names)}")
        
        # Create sheet mapping for case-insensitive comparison
        sheet_mapping = {}
        for sheet_name in xls.sheet_names:
            normalized = normalize_sheet_name(sheet_name)
            sheet_mapping[normalized] = sheet_name
        
        # Check for target sheets
        matches_found = 0
        for target_sheet in target_sheets:
            normalized_target = normalize_sheet_name(target_sheet)
            
            if normalized_target in sheet_mapping:
                original_sheet = sheet_mapping[normalized_target]
                
                try:
                    df = pd.read_excel(tmp_file_path, sheet_name=original_sheet)
                    
                    if df.empty:
                        file_log.append(f"   ⚠️ Sheet '{original_sheet}' is empty")
                        continue
                    
                    # Add all data from this sheet with metadata
                    for index, row in df.iterrows():
                        row_data = row.to_dict()
                        results[target_sheet].append({
                            'data': row_data,
                            'source_file': file.name,
                            'source_sheet': original_sheet,
                            'row_index': index
                        })
                    
                    matches_found += 1
                    file_log.append(f"   ✅ Found sheet '{original_sheet}' ({len(df)} rows, {len(df.columns)} columns)")
                    
                except Exception as e:
                    file_log.append(f"   ❌ Error reading sheet '{original_sheet}': {str(e)}")
            else:
                file_log.append(f"   ❌ Sheet '{target_sheet}' not found")
        
        if matches_found == 0:
            file_log.append(f"   ❌ No matching sheets found in file")
        else:
            file_log.append(f"   ✅ Successfully processed {matches_found} matching sheets")
        
        # Clean up temporary file
        os.unlink(tmp_file_path)
        
    except Exception as e:
        file_log.append(f"❌ Error processing file {file.name}: {str(e)}")
    
    return results, file_log

def create_individual_excel_file(sheet_name, data_list):
    """Create individual Excel file for a specific sheet type"""
    if not data_list:
        return None
    
    # Create Excel in memory
    output = io.BytesIO()
    
    # Extract all data and create combined DataFrame
    all_rows = []
    
    for item in data_list:
        # Add the actual data with source metadata
        row_data = item['data'].copy()
        row_data['_Source_File'] = item['source_file']
        row_data['_Source_Sheet'] = item['source_sheet']
        row_data['_Original_Row'] = item['row_index']
        all_rows.append(row_data)
    
    # Create DataFrame
    df_combined = pd.DataFrame(all_rows)
    
    # Write to Excel
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_combined.to_excel(writer, sheet_name='Data', index=False)
    
    output.seek(0)
    return output

def reset_all_data():
    """Reset all session data"""
    st.session_state.merged_data = defaultdict(list)
    st.session_state.processed_files = set()
    st.session_state.processing_log = []

# Main UI
st.title("📊 Excel Sheet Merger Tool")
st.markdown("**Merge data from sheets with matching names across multiple Excel files**")
st.markdown("---")

# Sidebar controls
with st.sidebar:
    st.header("🔧 Controls")
    
    # Preview settings
    st.subheader("Preview Settings")
    preview_limit = st.selectbox(
        "Process file limit (for testing):",
        [10, 50, 100, 500, "All Files"],
        index=4
    )
    
    if preview_limit == "All Files":
        preview_limit = float('inf')
    
    # Reset controls
    st.subheader("Reset Options")
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🗑️ Reset All", type="secondary"):
            reset_all_data()
            st.success("All data reset!")
    
    with col2:
        if st.button("📝 Clear Log"):
            st.session_state.processing_log = []
            st.success("Log cleared!")
    
    # Statistics
    st.subheader("📈 Statistics")
    st.metric("Files Processed", len(st.session_state.processed_files))
    st.metric("Sheets Merged", len(st.session_state.merged_data))
    
    if st.session_state.merged_data:
        total_rows = sum(len(data) for data in st.session_state.merged_data.values())
        st.metric("Total Rows", total_rows)

# Main content area
col1, col2 = st.columns([2, 1])

with col1:
    st.header("📁 File Upload")
    
    # File uploader
    uploaded_files = st.file_uploader(
        "Upload Excel files (.xlsx)",
        type=['xlsx'],
        accept_multiple_files=True,
        help="Upload as many Excel files as you need - no limit!"
    )
    
    # Show file count info
    if uploaded_files:
        st.info(f"📁 {len(uploaded_files)} files selected for processing")

with col2:
    st.header("📋 Sheet Names")
    
    # Sheet name input
    sheet_input = st.text_area(
        "Sheet names (comma-separated):",
        placeholder="Final MR_AC, final MR, Summary",
        help="Enter sheet names separated by commas. Matching is case and space insensitive."
    )
    
    # Parse sheet names
    target_sheets = []
    if sheet_input.strip():
        target_sheets = [sheet.strip() for sheet in sheet_input.split(',') if sheet.strip()]
    
    if target_sheets:
        st.success(f"Looking for {len(target_sheets)} sheets:")
        for sheet in target_sheets:
            st.write(f"• {sheet}")

# Processing section
if uploaded_files and target_sheets:
    st.header("⚙️ Processing")
    
    # Filter new files
    new_files = [f for f in uploaded_files if f.name not in st.session_state.processed_files]
    
    if new_files:
        st.info(f"Found {len(new_files)} new files to process")
        
        if st.button("🚀 Process Files", type="primary"):
            # Determine how many files to process
            files_to_process = new_files if preview_limit == float('inf') else new_files[:int(preview_limit)]
            
            # Create progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Process files
            for i, file in enumerate(files_to_process):
                status_text.text(f"Processing {i+1}/{len(files_to_process)}: {file.name}...")
                
                # Process file
                file_results, file_log = process_excel_file(file, target_sheets)
                
                # Add to session data
                for sheet_name, data_list in file_results.items():
                    st.session_state.merged_data[sheet_name].extend(data_list)
                
                # Add to processed files
                st.session_state.processed_files.add(file.name)
                
                # Add to log
                for log_entry in file_log:
                    log_message(log_entry)
                
                # Update progress
                progress_bar.progress((i + 1) / len(files_to_process))
            
            status_text.text("✅ Processing complete!")
            if len(files_to_process) < len(new_files):
                st.success(f"Processed {len(files_to_process)} files (limited by preview setting). {len(new_files) - len(files_to_process)} files remaining.")
            else:
                st.success(f"Processed all {len(files_to_process)} files successfully!")
    else:
        st.info("All uploaded files have already been processed.")

# Results section
if st.session_state.merged_data:
    st.header("📥 Download Separate Excel Files")
    
    # Summary statistics
    col1, col2, col3, col4 = st.columns(4)
    
    total_sheets = len(st.session_state.merged_data)
    total_rows = sum(len(data) for data in st.session_state.merged_data.values())
    total_files = len(st.session_state.processed_files)
    
    with col1:
        st.metric("Sheet Types Found", total_sheets)
    with col2:
        st.metric("Total Rows Merged", total_rows)
    with col3:
        st.metric("Files Processed", total_files)
    with col4:
        processed_percentage = (len(st.session_state.processed_files) / len(uploaded_files) * 100) if uploaded_files else 0
        st.metric("Processing Progress", f"{processed_percentage:.1f}%")
    
    # Download separate Excel files for each sheet type
    st.subheader("📥 Download separate Excel files for each sheet type:")
    
    for sheet_name, data_list in st.session_state.merged_data.items():
        if data_list:
            # Count files contributing to this sheet
            file_counts = {}
            for item in data_list:
                file_name = item['source_file']
                file_counts[file_name] = file_counts.get(file_name, 0) + 1
            
            # Create download section for this sheet
            col1, col2 = st.columns([3, 1])
            
            with col1:
                st.write(f"**{sheet_name}**")
                st.caption(f"{len(data_list)} rows from {len(file_counts)} files")
            
            with col2:
                # Create individual Excel file
                excel_file = create_individual_excel_file(sheet_name, data_list)
                if excel_file:
                    file_name = f"{sheet_name}.xlsx"
                    st.download_button(
                        label=f"📥 Download",
                        data=excel_file,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{sheet_name}"
                    )
    
    # Detailed breakdown
    st.subheader("📋 Sheet Breakdown")
    
    breakdown_data = []
    for sheet_name, data_list in st.session_state.merged_data.items():
        file_counts = {}
        for item in data_list:
            file_name = item['source_file']
            file_counts[file_name] = file_counts.get(file_name, 0) + 1
        
        breakdown_data.append({
            'Sheet Type': sheet_name,
            'Total Rows': len(data_list),
            'Files Contributing': len(file_counts),
            'Sample Files': ', '.join(list(file_counts.keys())[:3]) + ('...' if len(file_counts) > 3 else '')
        })
    
    if breakdown_data:
        df_breakdown = pd.DataFrame(breakdown_data)
        st.dataframe(df_breakdown, use_container_width=True)
    
    # Preview section
    st.subheader("👀 Data Preview")
    
    for sheet_name, data_list in st.session_state.merged_data.items():
        with st.expander(f"Preview: {sheet_name} ({len(data_list)} rows)"):
            if data_list:
                # Show first few rows of actual data
                preview_data = []
                for item in data_list[:5]:  # Show first 5 rows
                    row_data = item['data'].copy()
                    row_data['_Source_File'] = item['source_file']
                    row_data['_Source_Sheet'] = item['source_sheet']
                    preview_data.append(row_data)
                
                if preview_data:
                    df_preview = pd.DataFrame(preview_data)
                    st.dataframe(df_preview, use_container_width=True)
                    if len(data_list) > 5:
                        st.caption(f"Showing first 5 of {len(data_list)} rows...")

# Processing log
if st.session_state.processing_log:
    with st.expander("📋 Processing Log", expanded=False):
        log_container = st.container()
        with log_container:
            for log_entry in st.session_state.processing_log[-50:]:  # Show last 50 entries
                if log_entry['type'] == 'error':
                    st.error(log_entry['message'])
                elif log_entry['type'] == 'warning':
                    st.warning(log_entry['message'])
                else:
                    st.info(log_entry['message'])

# Footer
st.markdown("---")
st.markdown(
    """
    **Instructions for Processing Large File Sets:**
    1. **Upload All Files**: Select all your Excel files at once (no limit!)
    2. **Enter Sheet Names**: Specify which sheet names to look for
    3. **Set Processing Limit**: Use sidebar to process all files or test with smaller batches
    4. **Process**: Click process to extract matching sheets from all files
    5. **Download**: Get separate Excel files for each sheet type
    """
)