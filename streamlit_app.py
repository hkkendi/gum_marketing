import streamlit as st
import pandas as pd
import io
from datetime import datetime, time
import os
from pathlib import Path

# Required columns
TODO_REQUIRED_COLUMNS = ['Activity Company / ID', 'Assign To (Handler 1)', 'Assign To (Handler 2)']

# Initialize session state
if 'contact_data' not in st.session_state:
    st.session_state.contact_data = None
if 'last_modified_contact' not in st.session_state:
    st.session_state.last_modified_contact = None
if 'last_check_time' not in st.session_state:
    st.session_state.last_check_time = None

def find_contact_file():
    """Find contact file in possible locations"""
    possible_paths = [
        "Contact (res.partner).xlsx",  # Same directory
        "./Contact (res.partner).xlsx",  # Explicit current directory
        "../Contact (res.partner).xlsx",  # Parent directory
        "data/Contact (res.partner).xlsx",  # Data subdirectory
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
            
    return None

def load_contact_file():
    """Load contact file with improved error handling"""
    try:
        # Find the contact file
        contact_path = find_contact_file()
        
        if contact_path:
            st.write(f"Found contact file at: {contact_path}")  # Debug info
            last_modified = datetime.fromtimestamp(os.path.getmtime(contact_path))
            data = pd.read_excel(contact_path)
            return data, last_modified
            
        # Debug information
        st.write("Current working directory:", os.getcwd())
        st.write("Directory contents:", os.listdir())
        
        return None, None
        
    except Exception as e:
        st.error(f"Error loading contact file: {str(e)}")
        return None, None

def check_automatic_contact():
    """Check and load contact file"""
    contact_data, contact_modified = load_contact_file()
    
    if contact_data is not None:
        st.session_state.contact_data = contact_data
        st.session_state.last_modified_contact = contact_modified



def process_data(todo_df, contact_df):
    # Keep only required columns
    todo_df = todo_df[TODO_REQUIRED_COLUMNS]
    todo_df = todo_df.drop_duplicates(subset=['Activity Company / ID'], keep='first')
    
    contact_df = contact_df[['Name', 'ID', 'GUM Reference ID', 'Lead Sales Rep 1', 'Lead Sales Rep 2']]
    
    merged_df = todo_df.merge(contact_df, 
                            left_on='Activity Company / ID', 
                            right_on='ID', 
                            how='left')
    
    def check_handlers_match(row):
        handler1_match = str(row['Assign To (Handler 1)']) == str(row['Lead Sales Rep 1'])
        handler2_match = str(row['Assign To (Handler 2)']) == str(row['Lead Sales Rep 2'])
        return 'YES' if handler1_match and handler2_match else 'NO'
    
    merged_df['Check Handler Match'] = merged_df.apply(check_handlers_match, axis=1)
    
    return merged_df

def main():
    st.title("Excel Data Processor")
    
    # Check for automatic contact file updates at specified times
    if should_check_files():
        check_automatic_contact()
        st.rerun()
    
    # Display automatic contact file loading status
    st.subheader("Contact File Status")
    if st.session_state.last_modified_contact:
        st.write(f"Contact file last modified: {st.session_state.last_modified_contact}")
    else:
        st.warning("""Contact file not found. Please ensure 'Contact (res.partner).xlsx' is in the same directory as the app.
                  Current working directory: """ + os.getcwd())
    
    # Manual refresh button for automatic contact file
    if st.button("Refresh Contact File"):
        check_automatic_contact()
        st.rerun()
    
    # Manual file upload section with expanded help text
    st.subheader("Upload Activity File")
    st.info("Upload any Excel file that contains these required columns:\n" + 
            ", ".join(TODO_REQUIRED_COLUMNS))
    todo_file = st.file_uploader("Choose Excel file", type=["xlsx", "xls"])
    
    # Override contact file section
    st.subheader("Override Contact File (Optional)")
    manual_contact_file = st.file_uploader("Upload Contact (res.partner).xlsx", type=["xlsx", "xls"])
    
    # Process data
    todo_df = None
    contact_df = None
    
    # Get To Do data from manual upload and validate
    if todo_file is not None:
        try:
            temp_df = pd.read_excel(todo_file)
            is_valid, message = validate_todo_file(temp_df)
            if is_valid:
                todo_df = temp_df
                st.success(f"Successfully loaded: {todo_file.name}")
            else:
                st.error(message)
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
    
    # Get Contact data (prioritize manual upload over automatic)
    if manual_contact_file is not None:
        contact_df = pd.read_excel(manual_contact_file)
    elif st.session_state.contact_data is not None:
        contact_df = st.session_state.contact_data
    
    # Process data if both files are available
    if todo_df is not None and contact_df is not None:
        try:
            result_df = process_data(todo_df, contact_df)
            
            st.subheader("Results")
            st.write("Preview of processed data:")
            st.dataframe(result_df.head())
            
            # Create download button
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False, sheet_name='Processed Data')
            
            st.download_button(
                label="Download processed file",
                data=buffer.getvalue(),
                file_name="processed_result.xlsx",
                mime="application/vnd.ms-excel"
            )
            
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            st.error("Please make sure your Excel files have the required columns")
    else:
        if todo_file is None:
            st.info("Please upload an Excel file with the required columns")
        elif contact_df is None:
            st.info("Waiting for Contact file (either automatic or manual upload)")

if __name__ == "__main__":
    main()