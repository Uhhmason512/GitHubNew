import os
import sys
import time
import urllib.request
import zipfile
import pandas as pd
import pyautogui
import keyboard
import pytesseract
from PIL import Image
import pyperclip
from pynput import mouse, keyboard as kb
from threading import Thread
from screeninfo import get_monitors
import tkinter as tk
from tkinter import simpledialog, filedialog, messagebox
import cv2
import numpy as np
import subprocess
import urllib.request
import zipfile


def ensure_tesseract_installed():
    tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    if not os.path.exists(tesseract_path):
        print("Tesseract OCR not found. Downloading...")
        url = "https://github.com/tesseract-ocr/tesseract/releases/download/5.0.0/tesseract-ocr-w64-setup-v5.0.0.20190623.exe"
        local_zip_path = "tesseract-ocr-w64-setup-v5.0.0.20190623.zip"
        
        # Download the file from the URL
        urllib.request.urlretrieve(url, local_zip_path)
        
        # Extract the downloaded zip file
        with zipfile.ZipFile(local_zip_path, 'r') as zip_ref:
            zip_ref.extractall(r'C:\Program Files\Tesseract-OCR')
        
        os.remove(local_zip_path)  # Clean up the zip file
        print("Tesseract OCR installed successfully.")
    else:
        print("Tesseract OCR is already installed.")

# Ensure Tesseract OCR is installed before proceeding
ensure_tesseract_installed()

# Set the path to Tesseract OCR executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


# List to store all recorded sessions
all_recorded_events = []

# Function to show input dialog using Tkinter
def show_input_dialog():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    include_csv = messagebox.askyesno("Include Excel?", "Do you want to include an Excel file?")
    sheet_name = None
    column_name = None
    if include_csv:
        file_path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            sheet_name = simpledialog.askstring("Sheet Name", "Enter the sheet name to extract data from:")
            column_name = simpledialog.askstring("Column Name", "Enter the column name to extract names from:")
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            names = df[column_name].tolist()
        else:
            messagebox.showwarning("No File", "No file was selected. Exiting.")
            exit()
    else:
        names = []

    playback_count = simpledialog.askinteger("Playback Count", "How many times do you want playback?")

    return include_csv, names, playback_count

include_csv, names, playback_count = show_input_dialog()

# Initialize the global counter
i = 0

# Debounce times in seconds
debounce_time = 0.2
f12_debounce_time = 2.0  # Longer debounce time for F12 operations
f11_debounce_time = 1.0  # Debounce time for F11 operations
double_click_threshold = 0.3  # Threshold for detecting double clicks

# Timestamps for debouncing
last_time = time.time()
last_f12_time = time.time()
last_f11_time = time.time()
last_click_time = 0  # Initialize last click time

# Variable to track the current event
current_event = None

# Current session events
current_session_events = []

# Mouse listener variable
mouse_listener = None

# Variable to track the toggle state
mouse_output_enabled = True

# Variable to track the kill switch state
kill_switch_activated = False

# Function to print new recorded events
def print_new_events():
    last_len = 0
    while recording_active and not kill_switch_activated:
        current_len = len(current_session_events)
        if current_len > last_len:
            for event in current_session_events[last_len:current_len]:
                print(f"New event recorded: {event}")
            last_len = current_len
        time.sleep(0.1)

# Function to get the current mouse position when F11 is pressed
def get_mouse_position():
    global current_event, last_f11_time
    print("Move your mouse to the desired position and press F11.")

    while not kill_switch_activated:
        current_time = time.time()
        if current_event == 'f11' and current_time - last_f11_time >= f11_debounce_time:
            last_f11_time = current_time
            x, y = pyautogui.position()
            print(f"\nMouse position captured: X={x}, Y={y}")
            current_event = None  # Reset the current event
            return x, y

# Function to define a region by capturing two mouse positions
def define_region():
    print("Move your mouse to the top-left corner and press F11.")
    top_left_x, top_left_y = get_mouse_position()
    print(f"Top-left corner: ({top_left_x}, {top_left_y})")

    print("Move your mouse to the bottom-right corner and press F11.")
    bottom_right_x, bottom_right_y = get_mouse_position()
    print(f"Bottom-right corner: ({bottom_right_x}, {bottom_right_y})")

    width = bottom_right_x - top_left_x
    height = bottom_right_y - top_left_y
    region = (top_left_x, top_left_y, width, height)
    print(f"Region defined: {region}")
    return region

# Function to handle the F12 key event
def handle_f12():
    global last_f12_time
    current_time = time.time()
    if current_time - last_f12_time < f12_debounce_time:
        return  # Skip this F12 key press if it's within the debounce time
    last_f12_time = current_time
    region = define_region()
    if region:
        current_session_events.append(('f12', region))
        print(f"Recorded f12 event with region: {region}")  # Debugging print

def handle_f12_thread():
    thread = Thread(target=handle_f12)
    thread.start()

# Function to record keyboard events
def record_event(event):
    global current_session_events, last_time, recording_active, current_event, last_f11_time
    current_time = time.time()

    if event.name == 'f12':
        handle_f12_thread()
    elif event.name == 'f11':
        if current_time - last_f11_time >= f11_debounce_time:
            current_event = 'f11'  # Update the current event
            last_f11_time = current_time
            event.suppress = True  # Suppress the F11 key event
    elif event.name == 'esc':
        print("Recording stopped")
        recording_active = False
        return
    else:
        if current_time - last_time < debounce_time:
            return  # Skip this key press if it's within the debounce time
        last_time = current_time
        key = event.name
        current_session_events.append(('key', key, current_time))

        # Suppress the key event so it doesn't reach the system
        event.suppress = True

# Function to hook keyboard and mouse events
def hook_events():
    global mouse_listener
    print("Keyboard events hooked and mouse listener started.")
    keyboard.hook(record_event, suppress=True)
    if mouse_listener:
        mouse_listener.stop()
    mouse_listener = mouse.Listener(on_click=on_click)
    mouse_listener.start()
    return mouse_listener

# Function to start recording
def start_recording():
    global recording_active, current_session_events
    current_session_events = []
    recording_active = True
    print("Recording... Press 'esc' to stop recording.")

    # Adding delay to avoid registering the starting keys
    time.sleep(0.5)

    listener = hook_events()

    # Start a thread to print new events
    event_printer_thread = Thread(target=print_new_events)
    event_printer_thread.start()

    while recording_active and not kill_switch_activated:
        time.sleep(0.1)

    keyboard.unhook_all()
    listener.stop()

    all_recorded_events.append(current_session_events)
    print("Recording stopped.")
    print(f"Final recorded events for this session: {current_session_events}")

# Function to record mouse clicks
def on_click(x, y, button, pressed):
    global last_click_time, mouse_output_enabled
    if not recording_active or kill_switch_activated:
        return False
    if button == mouse.Button.left and pressed:
        current_time = time.time()
        if current_time - last_click_time <= double_click_threshold:
            timestamp = current_time
            current_session_events.append(('click', (x, y), timestamp))
            print(f"Recorded click at ({x}, {y}) at {timestamp}")
        last_click_time = current_time

# Function to replay recorded events
def replay_events():
    global i
    delay_factor = 0.1  # Adjust this factor to control the speed of playback (0.5 means half the original delay)
    print("Replaying recorded events...")

    # Calculate the number of names to use based on playback_count
    num_names_to_use = min(playback_count, len(names))

    for playback_index in range(num_names_to_use if names else 1):
        if names:
            print(f"Playback iteration: {playback_index + 1}/{playback_count}")
            name = names[playback_index]
            i = names.index(name)
            print(f"Replaying events for name: {name}")
        else:
            print("No names provided. Replaying events without using names.")

        for session in all_recorded_events:
            session_start_time = time.time()
            for event in session:
                if kill_switch_activated:
                    print("Kill switch activated, stopping playback.")
                    return
                event_type, *event_data = event
                print(f"Replaying event: {event}")  # Debugging print
                if event_type == 'key':
                    key, event_time = event_data
                    delay = (event_time - session_start_time) * delay_factor
                    time.sleep(max(0, delay))  # Ensure non-negative delay
                    if names:
                        if key == 'esc':
                            continue
                        if key == 'alt':
                            if i < len(names):
                                pyautogui.typewrite(name)
                                i += 1
                            else:
                                print("No more names to type.")
                                break
                        else:
                            pyautogui.press(key)
                    else:
                        pyautogui.press(key)
                elif event_type == 'click':
                    position, event_time = event_data
                    delay = (event_time - session_start_time) * delay_factor
                    time.sleep(max(0, delay))  # Ensure non-negative delay
                    pyautogui.click(position)
                elif event_type == 'f12':
                    region = event_data[0]
                    print(f"Using region: {region}")  # Debugging print
                    text = extract_text_from_region(region)
                    print(f"Extracted text: {text}")
                    copy_to_clipboard(text)
                    append_to_excel(text)
                session_start_time = event_time  # Update session start time for the next event

        if names:
            print(f"Playback completed for name: {name}")
        else:
            print("Playback completed for iteration without names.")

    print("All playbacks completed.")

# Function to preprocess image for better OCR accuracy
def preprocess_image(image):
    # Convert the image to grayscale
    gray = cv2.cvtColor(np.array(image), cv2.COLOR_BGR2GRAY)
    
    # Apply thresholding to get a binary image
    _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    # Resize image to improve OCR accuracy
    scale_percent = 150  # Percent of original size
    width = int(binary.shape[1] * scale_percent / 100)
    height = int(binary.shape[0] * scale_percent / 100)
    dim = (width, height)
    resized = cv2.resize(binary, dim, interpolation = cv2.INTER_AREA)
    
    return Image.fromarray(resized)

# Function to use Tesseract OCR to extract text from a region
def extract_text_from_region(region):
    print(f"Capturing screenshot for region: {region}")
    screenshot = pyautogui.screenshot(region=region)
    preprocessed_image = preprocess_image(screenshot)
    preprocessed_image.save("preprocessed_screenshot.png")  # Save for debugging purposes
    
    # Configure Tesseract to use only digits and letters
    custom_config = r'--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'
    text = pytesseract.image_to_string(preprocessed_image, config=custom_config)
    print(f"Extracted text: {text}")
    return text

# Function to copy text to clipboard
def copy_to_clipboard(text):
    pyperclip.copy(text)

# Function to append text to an Excel file
def append_to_excel(text):
    file_path = 'extracted_text.xlsx'
    new_row = pd.DataFrame({"Extracted Text": [text]})

    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        df = pd.concat([df, new_row], ignore_index=True)
    else:
        df = new_row

    df.to_excel(file_path, index=False)
    print(f"Text appended to {file_path}")

# Function to display screen size and resolution of all displays
def display_screen_info():
    monitors = get_monitors()
    for monitor in monitors:
        print(f"Monitor: {monitor.name}")
        print(f"Width: {monitor.width}, Height: {monitor.height}")
        print(f"Width (mm): {monitor.width_mm}, Height (mm): {monitor.height_mm}")
        print(f"Pixel Density: {monitor.width / monitor.width_mm:.2f} DPI\n")

# Function to activate the kill switch
def activate_kill_switch():
    global kill_switch_activated, recording_active
    kill_switch_activated = True
    recording_active = False
    print("Kill switch activated. Exiting...")

# Class to redirect stdout to the text widget
class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.insert("end", str)
        self.widget.see("end")

    def flush(self):
        pass

# Function to create and run the persistent GUI
def create_persistent_gui():
    root = tk.Tk()
    root.title("Control Panel")
    root.geometry("600x400")  # Set window size

    text_box = tk.Text(root, wrap="word", height=20, width=80)
    text_box.pack(pady=10, padx=10)

    sys.stdout = TextRedirector(text_box)

    kill_switch_button = tk.Button(root, text="Activate Kill Switch", command=activate_kill_switch)
    kill_switch_button.pack(pady=10)

    # Make sure the window stays on top
    root.attributes("-topmost", True)

    root.mainloop()

# Main function to run the script
def main():
    display_screen_info()
    gui_thread = Thread(target=create_persistent_gui)
    gui_thread.start()
    print("Press 'ctrl + r' to start recording.")
    keyboard.wait('ctrl+r')
    if not kill_switch_activated:
        start_recording()
        if not kill_switch_activated:
            replay_events()

if __name__ == "__main__":
    main()