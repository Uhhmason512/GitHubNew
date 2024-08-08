import ctypes
import os
import sys
import time
import logging
from threading import Timer
from ctypes import wintypes

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

class blockInput:
    def OnMouseEvent(self, nCode, wParam, lParam):
        if nCode < 0:
            return ctypes.windll.user32.CallNextHookEx(self.hooked, nCode, wParam, lParam)

        # Mouse left button down/up (WM_LBUTTONDOWN/WM_LBUTTONUP)
        if wParam == 0x0201 or wParam == 0x0202:
            logging.info("Blocking left mouse click event")
            return 1  # Block the event
        # Mouse right button down/up (WM_RBUTTONDOWN/WM_RBUTTONUP)
        elif wParam == 0x0204 or wParam == 0x0205:
            logging.info("Blocking right mouse click event")
            return 1  # Block the event

        return ctypes.windll.user32.CallNextHookEx(self.hooked, nCode, wParam, lParam)

    def unblock(self):
        logging.info(" -- Unblock!")
        if self.t.is_alive():
            self.t.cancel()
        if self.hooked:
            ctypes.windll.user32.UnhookWindowsHookEx(self.hooked)
            self.hooked = None

    def block(self, timeout=10):
        self.t = Timer(timeout, self.unblock)
        self.t.start()

        logging.info(" -- Block!")
        CMPFUNC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_int, ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM)
        self.pointer = CMPFUNC(self.OnMouseEvent)

        # Load the current process handle explicitly
        module_handle = ctypes.windll.kernel32.GetModuleHandleW(None)
        if not module_handle:
            logging.error(f"Failed to get module handle. Error code: {ctypes.GetLastError()}")
            return

        self.hooked = ctypes.windll.user32.SetWindowsHookExA(14, self.pointer, module_handle, 0)

        if not self.hooked:
            logging.error(f"Failed to set hook. Error code: {ctypes.GetLastError()}")
            return

        # Run the message loop
        msg = wintypes.MSG()
        while ctypes.windll.user32.GetMessageA(ctypes.byref(msg), None, 0, 0) != 0:
            ctypes.windll.user32.TranslateMessage(ctypes.byref(msg))
            ctypes.windll.user32.DispatchMessageA(ctypes.byref(msg))

    def __init__(self):
        self.hooked = None

if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)
    if not is_admin():
        # Re-run the script with admin rights
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, ' '.join(sys.argv), None, 1)
    else:
        block = blockInput()
        block.block()

        t0 = time.time()
        while time.time() - t0 < 10:
            time.sleep(1)
            print(time.time() - t0)

        block.unblock()
        logging.info("Done.")
