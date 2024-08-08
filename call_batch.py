import os
import sys
import psutil
import streamlit as st
import pandas as pd
import openpyxl
import glob
import base64
from geopy.distance import geodesic
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut
import time
import shutil
import subprocess

# Set the working directory to the directory of the script
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Function to copy config.toml to user home directory
def copy_config_to_user_home():
    source_config_path = os.path.join(os.path.dirname(__file__), 'config.toml')
    home_dir = os.path.expanduser('~')
    destination_dir = os.path.join(home_dir, '.streamlit')
    destination_config_path = os.path.join(destination_dir, 'config.toml')

    if not os.path.exists(destination_dir):
        os.makedirs(destination_dir)

    if not os.path.exists(destination_config_path):
        shutil.copy(source_config_path, destination_config_path)
        print(f"Copied config.toml to {destination_config_path}")

    with open(destination_config_path, 'r') as f:
        print(f"Config file content:\n{f.read()}")

# Copy the configuration file to user home
copy_config_to_user_home()

# Run the batch file
subprocess.call(['run_streamlit.bat'])
