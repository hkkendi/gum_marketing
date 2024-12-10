import streamlit as st
import pandas as pd
import io

def process_data(todo_df, contact_df):
    # 2. Process To Do.xlsx
    # Keep only required columns
    todo_df = todo_df[['Activity Company / ID', 'Assign To (Handler 1)', 'Assign To (Handler 2)']]
    # Remove duplicates based on 'Activity Company / ID', keep first occurrence
    todo_df = todo_df.drop_duplicates(subset=['Activity Company / ID'], keep='first')
    
    # 3. Process Contact (res.partner).xlsx
    contact_df = contact_df[['Name', 'ID', 'GUM Reference ID', 'Lead Sales Rep 1', 'Lead Sales Rep 2']]
    
    # 4. Merge dataframes
    # Merge on 'Activity Company / ID' matching with 'ID' from contacts
    merged_df = todo_df.merge(contact_df, 
                            left_on='Activity Company / ID', 
                            right_on='ID', 
                            how='left')
    
    # 5. Create 'Check Handler Match' column
    def check_handlers_match(row):
        handler1_match = str(row['Assign To (Handler 1)']) == str(row['Lead Sales Rep 1'])
        handler2_match = str(row['Assign To (Handler 2)']) == str(row['Lead Sales Rep 2'])
        return 'YES' if handler1_match and handler2_match else 'NO'
    
    merged_df['Check Handler Match'] = merged_df.apply(check_handlers_match, axis=1)
    
    return merged_df

st.title("Excel Data Processor")

# File uploaders
todo_file = st.file_uploader("Upload To Do.xlsx", type="xlsx")
contact_file = st.file_uploader("Upload Contact (res.partner).xlsx", type="xlsx")

if todo_file is not None and contact_file is not None:
    try:
        # Read Excel files
        todo_df = pd.read_excel(todo_file)
        contact_df = pd.read_excel(contact_file)
        
        # Process the data
        result_df = process_data(todo_df, contact_df)
        
        # Show preview of the result
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