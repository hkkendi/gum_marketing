# ====== IMPORTS ======
import streamlit as st
import pandas as pd
import io
from datetime import datetime
import os
import pytz

# ====== CONSTANTS ======
# Required columns
TODO_REQUIRED_COLUMNS = ['Activity Company / ID', 'Assign To (Handler 1)', 'Assign To (Handler 2)']
GUM_CONTACT_COLUMNS = ['Email*', 'Contact Company/ID', 'Contact Company', 'Contact Company/GUM Reference ID']



# ====== SESSION STATE INITIALIZATION ======
if 'contact_data' not in st.session_state:
    st.session_state.contact_data = None
if 'last_modified_contact' not in st.session_state:
    st.session_state.last_modified_contact = None
if 'gum_contact_data' not in st.session_state:
    st.session_state.gum_contact_data = None
if 'last_modified_gum' not in st.session_state:
    st.session_state.last_modified_gum = None

# ====== HELPER FUNCTIONS ======
def format_hk_time(timestamp):
    """Convert timestamp to Hong Kong time and format it"""
    if isinstance(timestamp, str):
        # If timestamp is already a string, parse it first
        timestamp = datetime.strptime(timestamp, '%Y-%m-%d %H:%M:%S')
    hk_timezone = pytz.timezone('Asia/Hong_Kong')
    # If timestamp is naive (no timezone info), assume it's HK time
    if timestamp.tzinfo is None:
        timestamp = hk_timezone.localize(timestamp)
    return timestamp.strftime('%Y-%m-%d %H:%M')

def get_file_info(file_path):
    """Get file modification time properly from Windows"""
    try:
        # Get file stats
        stats = os.stat(file_path)
        # Get the modification time as a datetime object
        mod_time = datetime.fromtimestamp(stats.st_mtime)
        return mod_time
    except Exception as e:
        st.error(f"Error getting file info: {str(e)}")
        return None

def load_contact_file():
    """Load contact file and preserve its original modification time"""
    try:
        contact_path = os.path.join(os.getcwd(), "Contact (res.partner).xlsx")
        if os.path.exists(contact_path):
            # Get the actual Windows file modification time
            last_modified = get_file_info(contact_path)
            data = pd.read_excel(contact_path)
            return data, last_modified
        return None, None
    except Exception as e:
        st.error(f"Error loading contact file: {str(e)}")
        return None, None

def load_gum_contact_file():
    """Load GUM contact file and preserve its original modification time"""
    try:
        gum_path = os.path.join(os.getcwd(), "GUM Resource Contact (gm.res.contact).xlsx")
        if os.path.exists(gum_path):
            # Get the actual Windows file modification time
            last_modified = get_file_info(gum_path)
            data = pd.read_excel(gum_path)
            return data, last_modified
        return None, None
    except Exception as e:
        st.error(f"Error loading GUM contact file: {str(e)}")
        return None, None

# For debugging purposes, add this at the start of your main UI
st.write("Debug - Current working directory:", os.getcwd())
st.write("Debug - Files in directory:", os.listdir())

def validate_todo_file(df):
    """Validate that the uploaded file has the required columns"""
    missing_columns = [col for col in TODO_REQUIRED_COLUMNS if col not in df.columns]
    if missing_columns:
        return False, f"Missing required columns: {', '.join(missing_columns)}"
    return True, "File is valid"

def process_data(todo_df, contact_df):
    """Process and merge the data"""
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

# ====== MAIN UI ======
# Title
st.title("Excel Data Processor")

# Section 1: Contact File Status and To Do Processing
st.subheader("Contact File Status")
contact_data, last_modified = load_contact_file()

# For Contact file:
if contact_data is not None:
    st.success("✅ Contact file loaded successfully")
    if last_modified:
        st.info(f"Last modified: {format_hk_time(last_modified)}")
    st.session_state.contact_data = contact_data
    st.session_state.last_modified_contact = last_modified
else:
    st.warning("⚠️ Contact file not found")

# Manual refresh button for contact file
if st.button("🔄 Refresh Contact File"):
    contact_data, last_modified = load_contact_file()
    if contact_data is not None:
        st.session_state.contact_data = contact_data
        st.session_state.last_modified_contact = last_modified
        st.rerun()

# Manual file upload section
st.subheader("Upload To Do Template File")
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

# Section divider
st.markdown("---")

# Section 2: GUM Resource Contact Lookup
st.title("GUM Resource Contact Lookup")

# Load automatic GUM contact data
gum_data, gum_modified = load_gum_contact_file()

# For GUM Resource Contact file:
if gum_data is not None:
    st.success("✅ GUM Resource Contact file loaded successfully")
    if gum_modified:
        st.info(f"Last modified: {format_hk_time(gum_modified)}")
    st.session_state.gum_contact_data = gum_data
else:
    st.warning("⚠️ GUM Resource Contact file not found")

# Refresh button for automatic GUM contact file
if st.button("🔄 Refresh GUM Contact File"):
    gum_data, gum_modified = load_gum_contact_file()
    if gum_data is not None:
        st.session_state.gum_contact_data = gum_data
        st.rerun()

# Manual override section for GUM Resource Contact file
st.subheader("Override GUM Resource Contact File (Optional)")
manual_gum_file = st.file_uploader("Upload GUM Resource Contact (gm.res.contact).xlsx", type=["xlsx", "xls"])

# Use manual file if uploaded, otherwise use automatic file
if manual_gum_file is not None:
    try:
        gum_data = pd.read_excel(manual_gum_file)
        st.success(f"Successfully loaded manual file: {manual_gum_file.name}")
    except Exception as e:
        st.error(f"Error reading manual file: {str(e)}")
        gum_data = st.session_state.gum_contact_data
else:
    gum_data = st.session_state.gum_contact_data

# Email lookup section
st.subheader("Email Lookup")
email = st.text_input("Enter Email Address for Lookup")

if st.button("Look Up Contact Details"):
    if gum_data is not None and email:
        # Find matching records
        matching_records = gum_data[gum_data['Email*'] == email]
        
        if not matching_records.empty:
            total_matches = len(matching_records)
            # Get unique records based on Contact Company/ID
            unique_records = matching_records.drop_duplicates(subset=['Contact Company/ID'])
            unique_count = len(unique_records)
            
            # Show the total matches found
            st.success(f"Found {total_matches} contact(s)!")
            
            # If there are duplicate records, show the summary first
            if total_matches > unique_count:
                st.info(f"📊 Summary: These {total_matches} results contain {unique_count} unique Contact Company/IDs.")
                st.markdown("*Showing unique records only:*")
            
            # Display only unique records
            for idx, record in unique_records.iterrows():
                contact_info = pd.DataFrame({
                    'Field': ['Contact Company/ID', 'Contact Company', 'Contact Company/GUM Reference ID'],
                    'Value': [
                        record['Contact Company/ID'],
                        record['Contact Company'],
                        record['Contact Company/GUM Reference ID']
                    ]
                })
                st.table(contact_info)
                # Add a small space between results if there are multiple
                if unique_count > 1:
                    st.markdown("")
        else:
            st.warning("No contact found with this email address")
    elif gum_data is None:
        st.error("Please ensure GUM Resource Contact file is loaded (either automatic or manual)")
    else:
        st.warning("Please enter an email address")