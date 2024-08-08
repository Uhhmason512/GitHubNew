from PIL import ImageGrab
import pytesseract
import pyperclip
from pynput import mouse

# Ensure pytesseract can find the tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Update this path as needed

start_pos = None
end_pos = None

def on_click(x, y, button, pressed):
    global start_pos, end_pos
    if pressed:
        if start_pos is None:
            start_pos = (x, y)
        else:
            end_pos = (x, y)
            return False

def capture_screen_area(start, end):
    x1, y1 = start
    x2, y2 = end
    bbox = (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
    return ImageGrab.grab(bbox)

def main():
    global start_pos, end_pos

    # Collect mouse click positions
    with mouse.Listener(on_click=on_click) as listener:
        listener.join()

    # Ensure both positions are set
    if start_pos and end_pos:
        # Capture the selected area of the screen
        img = capture_screen_area(start_pos, end_pos)

        # Extract text from the image
        text = pytesseract.image_to_string(img)

        # Copy text to clipboard
        pyperclip.copy(text)
        print("Text copied to clipboard:", text)
    else:
        print("Selection was not completed.")

if __name__ == "__main__":
    main()
