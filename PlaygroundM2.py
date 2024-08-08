import pandas as pd
import pyautogui
import keyboard
import time
import pytesseract
from PIL import Image

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Prompt the user to input the path to the csv file
file_path = input("Please enter the path to your CSV file: ")

# Read the csv file into a DataFrame
df = pd.read_csv(file_path)

# Display the screen size
screenWidth, screenHeight = pyautogui.size()
print(f'Screen size: {screenWidth}x{screenHeight}')

# Initialize an empty list to store the recorded keys
recorded_keys = []

names = df['Unique Value'].tolist()

# Initialize the global counter
i = 0

# Function to get the current mouse position when Ctrl+9 is pressed
def get_mouse_position():
    time.sleep(2)
    try:
        while not keyboard.is_pressed('ctrl+9'):
            x, y = pyautogui.position()
            print(f'X: {x}, Y: {y}', end='\r')
            time.sleep(0.1)
        print(f"\nMouse position captured: X={x}, Y={y}")
        return x, y
    except Exception as e:
        print(f"Error: {e}")
        return None, None

def record_key(event):
    global recorded_keys
    key = event.name
    if key == 'esc':  # Use 'esc' to stop recording
        print("Recording stopped")
        keyboard.unhook_all()
        return
    recorded_keys.append(key)
    time.sleep(0.1)  # debounce time for sticky keys

def replay_keys():
    global i
    print("Replaying the recorded keys...")
    for key in recorded_keys:
        if key == 'esc':
            continue
        if key == 'alt':
            i += 1
            if i < len(names):
                name = names[i]
                pyautogui.typewrite(name)
            else:
                print("No more names to type.")
                break
        pyautogui.press(key)
        time.sleep(0.4)  # Adjust the delay as needed
    print("Playback completed")

print("Press Ctrl+9 to capture the current mouse position.")
top_left_x, top_left_y = get_mouse_position()

if top_left_x is None or top_left_y is None:
    exit("Failed to capture the top-left corner. Exiting...")

print("Move your mouse to the bottom-right corner of the region and press Ctrl+9")
bottom_right_x, bottom_right_y = get_mouse_position()

if bottom_right_x is None or bottom_right_y is None:
    exit("Failed to capture the bottom-right corner. Exiting...")

# Define the region (left, top, width, height)
width = bottom_right_x - top_left_x
height = bottom_right_y - top_left_y
region = (top_left_x, top_left_y, width, height)
print(f"Region defined: {region}")

# Wait for the user to press 'ctrl + r' to start recording
print("Press 'ctrl + r' to start recording keys.")
keyboard.wait('ctrl+r')
keyboard.hook(record_key)

# Wait for the user to press 'esc' to stop recording
keyboard.wait('esc')

# Unhook the record_key function
keyboard.unhook(record_key)

# Replay the recorded keys
replay_keys()