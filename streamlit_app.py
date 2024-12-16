import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime, time

# File paths
FILE_PATH = r"C:\Users\KendiNg\Documents\Apps_running_files"
TODO_FILE = "To Do.xlsx"
CONTACT_FILE = "Contact (res.partner).xlsx"

# Initialize session state
if 'todo_data' not in st.session_state:
    st.session_state.todo_data = None
if 'contact_data' not in st.session_state:
    st.session_state.contact_data = None
if 'last_modified_todo' not in st.session_state:
    st.session_state.last_modified_todo = None
if 'last_modified_contact' not in st.session_state:
    st.session_state.last_modified_contact = None
if 'last_check_time' not in st.session_state:
    st.session_state.last_check_time = None

def should_check_files():
    """Determine if files should be checked based on current time"""
    current_time = datetime.now().time()
    check_times = [
        time(10, 0),  # 10:00 AM
        time(12, 30)  # 12:30 PM
    ]
    
    # If it's exactly one of our check times, return True
    if any(current_time.hour == t.hour and current_time.minute == t.minute for t in check_times):
        if (st.session_state.last_check_time is None or 
            st.session_state.last_check_time.hour != current_time.hour or 
            st.session_state.last_check_time.minute != current_time.minute):
            st.session_state.last_check_time = current_time
            return True
    return False

def load_file(file_path):
    """Load file and get its last modified time"""
    if os.path.exists(file_path):
        last_modified = datetime.fromtimestamp(os.path.getmtime(file_path))
        data = pd.read_excel(file_path)
        return data, last_modified
    return None, None

def check_automatic_files():
    """Check and load files"""
    todo_path = os.path.join(FILE_PATH, TODO_FILE)
    contact_path = os.path.join(FILE_PATH, CONTACT_FILE)
    
    # Load files
    todo_data, todo_modified = load_file(todo_path)
    contact_data, contact_modified = load_file(contact_path)
    
    # Update session state
    if todo_data is not None:
        st.session_state.todo_data = todo_data
        st.session_state.last_modified_todo = todo_modified
    
    if contact_data is not None:
        st.session_state.contact_data = contact_data
        st.session_state.last_modified_contact = contact_modified

def process_data(todo_df, contact_df):
    # Keep only required columns
    todo_df = todo_df[['Activity Company / ID', 'Assign To (Handler 1)', 'Assign To (Handler 2)']]
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
    
    # Check for automatic file updates at specified times
    if should_check_files():
        check_automatic_files()
        st.rerun()
    
    # Display automatic file loading status
    st.subheader("Automatic File Loading Status")
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("To Do.xlsx")
        if st.session_state.last_modified_todo:
            st.write(f"Last modified: {st.session_state.last_modified_todo}")
        else:
            st.write("File not found in automatic directory")
            
    with col2:
        st.write("Contact (res.partner).xlsx")
        if st.session_state.last_modified_contact:
            st.write(f"Last modified: {st.session_state.last_modified_contact}")
        else:
            st.write("File not found in automatic directory")
    
    # Manual refresh button for automatic loading
    if st.button("Refresh Automatic Files"):
        check_automatic_files()
        st.rerun()
    
    # Manual file upload section
    st.subheader("Manual File Upload")
    st.write("You can also upload files manually:")
    
    manual_todo_file = st.file_uploader("Upload To Do.xlsx", type="xlsx")
    manual_contact_file = st.file_uploader("Upload Contact (res.partner).xlsx", type="xlsx")
    
    # Process data based on either automatic or manual files
    todo_df = None
    contact_df = None
    
    # Prioritize manual uploads over automatic files
    if manual_todo_file is not None:
        todo_df = pd.read_excel(manual_todo_file)
    elif st.session_state.todo_data is not None:
        todo_df = st.session_state.todo_data
        
    if manual_contact_file is not None:
        contact_df = pd.read_excel(manual_contact_file)
    elif st.session_state.contact_data is not None:
        contact_df = st.session_state.contact_data
    
    # Process data if both files are available from either source
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
        st.info("Waiting for files (either automatic or manual upload)...")

if __name__ == "__main__":
    main()