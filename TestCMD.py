import sys
import os
import psutil
import streamlit as st
import time

def is_streamlit_running():
    """Check if a Streamlit process is already running."""
    for process in psutil.process_iter(attrs=['name']):
        if 'streamlit' in process.info['name']:
            return True
    return False

def start_streamlit():
    script_path = os.path.abspath(__file__)
    command = f'streamlit run "{script_path}"'
    os.system(f'start cmd /k {command}')
    print("Streamlit server started in a new command window.")

def run_streamlit_app():
    print("Running Streamlit app...")
    st.title("Hello, Streamlit!")
    st.write("This is a simple Streamlit app to demonstrate its functionality.")
    st.write("If you see this message, the app is running correctly.")
    print("Streamlit app should be running now.")

if __name__ == "__main__":
    print("Script started...")
    print(f"sys.argv: {sys.argv}")

    # Check if the Streamlit server is already running
    if not is_streamlit_running():
        start_streamlit()
        print("Waiting for Streamlit server to start...")
        time.sleep(10)  # Wait for the server to start

    run_streamlit_app()  # This will only run when Streamlit executes the script

    # Keep the script alive to ensure the Streamlit server continues running
    keep_open = True
    while keep_open:
        time.sleep(1)
