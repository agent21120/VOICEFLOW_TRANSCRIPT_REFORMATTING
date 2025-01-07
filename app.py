import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import streamlit as st

def format_transcript(input_file, output_name):
    # Read the uploaded file directly
    data = pd.read_csv(input_file)
    # Perform your formatting logic here...
    # Example: Save the processed data to an Excel file
    data.to_excel(output_name, index=False)


    # Prepare formatted conversation
    conversation = []
    for _, row in data.iterrows():
        start_time = row.get('start_time', 'Unknown Time')
        row_type = row.get('type', 'Unknown Type')
        if row_type not in ['debug', 'goto', 'knowledgeBase'] and pd.notna(row_type):
            if row_type == 'choice' and pd.notna(row.get('response')):
                buttons = row['response'].replace(',', ', ')
                conversation.append({
                    'START_TIME': start_time,
                    'TYPE': row_type,
                    'AGENT': f"BUTTONS DISPLAYED: {buttons}",
                    'USER': '',
                    'INTENT_MATCHED': ''
                })
            else:
                conversation.append({
                    'START_TIME': start_time,
                    'TYPE': row_type,
                    'AGENT': row.get('response', ''),
                    'USER': row.get('user_input', ''),
                    'INTENT_MATCHED': row.get('intent_matched', '')
                })

    # Create a DataFrame from the formatted data
    formatted_data = pd.DataFrame(conversation)

    # Save to Excel
    formatted_data.to_excel(output_file, index=False)

    # Add formatting
    workbook = load_workbook(output_file)
    sheet = workbook.active
    pink_fill = PatternFill(start_color="FFD1DC", end_color="FFD1DC", fill_type="solid")
    grey_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
        agent_cell = row[2]
        user_cell = row[3]
        if agent_cell.value:
            agent_cell.fill = pink_fill
        if user_cell.value:
            user_cell.fill = grey_fill

    workbook.save(output_file)

# Streamlit App Interface
st.title("Voiceflow Transcript Formatter")
st.write("Upload a Voiceflow transcript in CSV format, and get a formatted Excel file!")

uploaded_file = st.file_uploader("Upload your CSV file", type="csv")
output_name = st.text_input("Enter output file name (e.g., formatted_transcript.xlsx)")

if st.button("Process"):
    if uploaded_file and output_name:
        with open(output_name, "wb") as f:
            f.write(uploaded_file.getbuffer())
        format_transcript(uploaded_file.name, output_name)
        st.success("Transcript formatted successfully!")
        with open(output_name, "rb") as f:
            st.download_button(
                label="Download Formatted Excel File",
                data=f,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("Please upload a file and provide an output name.")
