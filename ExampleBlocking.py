import keyboard
import time
from pynput import mouse as pynput_mouse
import pyautogui

# List to store recorded events
recorded_events = []

# Flag to control recording state
recording_active = False

# Function to record keyboard events without performing them
def record_event(event):
    global recording_active
    if event.name == 'esc':
        recording_active = False
        return False  # Suppress the esc key and stop recording
    if recording_active:
        recorded_events.append(('key', event.name, time.time()))
        print(f"Recorded key event: {event.name} at {time.time()}")  # Debug logging
        return False  # Suppress all key events

# Function to record mouse clicks
def on_click(x, y, button, pressed):
    if recording_active:
        timestamp = time.time()
        action = 'pressed' if pressed else 'released'
        recorded_events.append((f'mouse_{action}', (x, y, button.name), timestamp))
        print(f"Recorded mouse {action} at ({x}, {y}) with {button.name} at {timestamp}")  # Debug logging

# Function to reset the keyboard state
def reset_keyboard_state():
    keys_to_release = ['shift', 'ctrl', 'alt', 'cmd', 'win', 'left shift', 'right shift', 'left ctrl', 'right ctrl', 'left alt', 'right alt']
    for key in keys_to_release:
        keyboard.release(key)
    print("Keyboard state reset.")  # Debug logging

# Function to start recording
def start_recording():
    global recording_active
    recording_active = True
    print("Recording started... Press 'esc' to stop recording.")  # Debug logging

    # Start mouse listener
    mouse_listener = pynput_mouse.Listener(on_click=on_click)
    mouse_listener.start()
    print("Mouse listener started.")  # Debug logging

    # Hook keyboard events and suppress them
    keyboard.hook(record_event, suppress=True)
    print("Keyboard hook set.")  # Debug logging

    # Wait for recording to stop
    while recording_active:
        time.sleep(0.1)

    # Unhook all keyboard events to stop suppression
    keyboard.unhook_all()
    mouse_listener.stop()
    print("Recording stopped.")  # Debug logging
    reset_keyboard_state()  # Reset keyboard state after unhooking

# Function to replay recorded events
def replay_events():
    reset_keyboard_state()
    print("Replaying recorded events...")  # Debug logging
    start_time = time.time()
    for event in recorded_events:
        event_type, event_data, event_time = event
        delay = event_time - start_time
        time.sleep(max(0, delay))  # Ensure non-negative delay

        if event_type == 'key':
            key = event_data
            if key == 'space':
                pyautogui.press(' ')
            elif key in ['enter', 'backspace', 'tab', 'left', 'right', 'up', 'down', 'shift', 'ctrl', 'alt']:
                pyautogui.press(key)
            else:
                pyautogui.typewrite(key)
        elif event_type == 'mouse_pressed':
            x, y, button = event_data
            pyautogui.mouseDown(x, y, button)
        elif event_type == 'mouse_released':
            x, y, button = event_data
            pyautogui.mouseUp(x, y, button)

    print("Playback completed.")  # Debug logging

# Function to control recording from the main function
def main():
    print("Press 'ctrl + r' to start recording.")
    keyboard.wait('ctrl+r')
    start_recording()

    print("Press 'esc' to stop recording.")
    while recording_active:
        time.sleep(0.1)  # Short delay to avoid busy-waiting

    replay_events()

if __name__ == "__main__":
    main()
