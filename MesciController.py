import pandas as pd
import csv
import pyautogui
import pyperclip
from PIL import Image
import pytesseract
import keyboard
import time

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Prompt the user to input the path to the Excel file
file_path = input("Please enter the path to your Excel file: ")

# Read the Excel file into a DataFrame
df = pd.read_excel(file_path)

def record_key(event):
    global recorded_keys
    key = event.name
    if key == 'esc':  # Use 'esc' to stop recording
        print("Recording stopped")
        keyboard.unhook_all()
        return
    recorded_keys.append(key)

def replay_keys():
    print("Replaying the recorded keys...")
    for key in recorded_keys:
        if key == 'esc':
            continue
        pyautogui.press(key)
        time.sleep(0.1)  # Adjust the delay as needed
    print("Playback completed")

while True:
    recorded_keys = []  # Clear the list for new recording
    print("Waiting for 'ctrl + r' to start recording...")

    # Wait for the specific key combination 'ctrl + r' to start recording
    keyboard.wait('ctrl+r')
    print("Recording started... Press 'esc' to stop recording.")
    
    # Hook the key press event after the specific key combination is pressed
    keyboard.hook(record_key)

    # Wait for 'esc' key to stop recording
    keyboard.wait('esc')
    
    # Unhook the event after 'esc' is pressed to avoid conflict
    keyboard.unhook(record_key)

    # Replay the recorded key inputs
    replay_keys()

    # Small delay before starting the next round
    time.sleep(1)