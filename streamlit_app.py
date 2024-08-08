import streamlit as st
import os
import json
import base64

# Boeing color scheme
BOEING_BLUE = "#0033A0"
BOEING_LIGHT_BLUE = "#A0CCEB"
BOEING_GRAY = "#444444"

# Function to load and display the Boeing logo
def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

def set_boeing_theme():
    logo_path = 'Boeing_full_logo.png'
    logo_base64 = get_base64_of_bin_file(logo_path)
    
    st.markdown(
        f"""
        <style>
            .reportview-container .main .block-container{{
                max-width: 1000px;
                padding-top: 2rem;
                padding-bottom: 2rem;
                padding-left: 5rem;
                padding-right: 5rem;
            }}
            .reportview-container .main {{
                color: {BOEING_GRAY};
                background-color: white;
            }}
            .sidebar .sidebar-content {{
                background-color: {BOEING_LIGHT_BLUE};
            }}
            h1 {{
                color: {BOEING_BLUE};
            }}
            .stButton>button {{
                color: white;
                background-color: {BOEING_BLUE};
                border-radius: 5px;
            }}
            .stTextInput>div>div>input {{
                border-color: {BOEING_BLUE};
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )
    
    st.markdown(
        f"""
        <img src="data:image/png;base64,{logo_base64}" alt="Boeing Logo" style="width:200px;margin-bottom:20px;">
        """,
        unsafe_allow_html=True,
    )

# Set Boeing theme
set_boeing_theme()

st.title("RCR Sheet Automator")

# Create input fields
input1 = st.text_input("Enter the directory path where the files are stored (e.g., C:\\Users\\YourName\\Documents\\Files): ")
input2 = st.text_input("Enter the name of the permanent Excel file (e.g., permanent_file.xlsm): ")
input3 = st.text_input("Enter the pattern for the monthly files (e.g., groceries_*.xlsx): ")
input4 = st.text_input("Enter the name of the sheet in the permanent file where the data should be added: ")
input5 = st.text_input("Enter the starting row number (e.g., 4): ")

# Add a button to submit the inputs
if st.button("Submit"):
    if input1 and input2 and input3 and input4 and input5:
        try:
            # Pass inputs to the main script
            params = {
                "directory_path": input1,
                "permanent_file_name": input2,
                "monthly_file_pattern": input3,
                "sheet_name": input4,
                "start_row": input5
            }
            with open('params.json', 'w') as f:
                json.dump(params, f)

            # Notify main script to proceed
            with open('trigger.txt', 'w') as f:
                f.write('run')

            st.success("Inputs submitted successfully. The process will run in the background.")

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
    else:
        st.warning("Please fill in all fields before submitting.")
