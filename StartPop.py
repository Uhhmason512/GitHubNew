import os
import subprocess
import sys

def main():
    # Get the directory of this script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Change the working directory to the script's directory
    os.chdir(script_dir)
    
    # Path to the Streamlit executable
    streamlit_executable = os.path.join(script_dir, 'FinalCode.py')

    # Command to open a new cmd window and run the Streamlit executable
    cmd_command = f'start cmd /k "{streamlit_executable}"'

    # Run the cmd command
    subprocess.run(cmd_command, shell=True)

if __name__ == "__main__":
    main()
