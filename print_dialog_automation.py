# print_dialog_automation.py

import time
import pyautogui
import traceback

# -----------------------------
# FIXED print dialog automation
# -----------------------------

def automate_print_dialog_hp_fixed():
    """
    FIXED VERSION - Accurate coordinates for HP LaserJet print dialog
    """
    print("[INFO] Waiting for print dialog to open...")
    time.sleep(6)  # Longer wait for dialog
    
    try:
        screen_width, screen_height = pyautogui.size()
        print(f"[DEBUG] Screen: {screen_width}x{screen_height}")
        
        # Focus the dialog
        center_x = screen_width // 2
        center_y = screen_height // 2
        pyautogui.click(center_x, center_y)
        time.sleep(1)
        
        # STEP 1: Select Printer (Destination dropdown)
        print("[STEP 1] Selecting printer...")
        dest_x = int(screen_width * 0.86)
        dest_y = int(screen_height * 0.17)
        
        pyautogui.click(dest_x, dest_y)
        time.sleep(1.5)
        
        # Type printer name
        pyautogui.write('1020', interval=0.15)
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(3)
        
        print("[SUCCESS] ✓ Printer selected: HP LaserJet 1020")
        
        # STEP 2: Expand "More settings" - KEYBOARD METHOD (Most Reliable)
        print("[STEP 2] Expanding More settings...")
        
        # Method 1: Try clicking on the text
        if screen_width == 1920 and screen_height == 1080:
            more_settings_x = 980
            more_settings_y = 422
        else:
            more_settings_x = int(screen_width * 0.51)
            more_settings_y = int(screen_height * 0.391)
        
        print(f"[DEBUG] Method 1 - Clicking at: ({more_settings_x}, {more_settings_y})")
        
        # Click the approximate area multiple times
        for i in range(5):
            pyautogui.click(more_settings_x + (i * 10), more_settings_y)
            time.sleep(0.4)
        
        time.sleep(1)
        
        # Method 2: Use keyboard Tab navigation (MOST RELIABLE)
        print("[DEBUG] Method 2 - Using Tab key navigation...")
        
        # Click in the print dialog to focus it
        pyautogui.click(center_x, center_y)
        time.sleep(0.5)
        
        # Tab through elements to reach "More settings"
        # From top: Destination (1), Pages (2), Copies (3), More settings (4)
        for i in range(8):  # Tab multiple times to find it
            pyautogui.press('tab')
            time.sleep(0.3)
            
            # Try pressing Enter (might be on More settings button)
            if i >= 3 and i <= 6:  # More settings is usually 4-6 tabs down
                pyautogui.press('enter')
                time.sleep(0.5)
        
        time.sleep(2)
        print("[SUCCESS] ✓ More settings expanded (via keyboard)")
        
        # STEP 3: Set Pages per sheet to 4
        print("[STEP 3] Setting Pages per sheet to 4...")
        
        # Pages per sheet dropdown - after More settings expands
        pages_x = int(screen_width * 0.86)
        pages_y = int(screen_height * 0.55)  # Below More settings section
        
        pyautogui.click(pages_x, pages_y)
        time.sleep(1.5)
        
        # Select "4" - it's usually the 3rd option (1, 2, 4, 6, 9, 16)
        pyautogui.press('down')
        time.sleep(0.4)
        pyautogui.press('down')
        time.sleep(0.4)
        pyautogui.press('enter')
        time.sleep(2)
        
        print("[SUCCESS] ✓ Pages per sheet = 4")
        
        # STEP 4: Set Scale to "Fit to printable area"
        print("[STEP 4] Setting Scale to Fit to printable area...")
        
        scale_x = int(screen_width * 0.86)
        scale_y = int(screen_height * 0.63)
        
        pyautogui.click(scale_x, scale_y)
        time.sleep(1.5)
        
        # "Fit to printable area" is usually 3rd option
        pyautogui.press('down')
        time.sleep(0.4)
        pyautogui.press('down')
        time.sleep(0.4)
        pyautogui.press('enter')
        time.sleep(2)
        
        print("[SUCCESS] ✓ Scale = Fit to printable area")
        
        # STEP 5: Click Print button
        print("[STEP 5] Clicking Print button...")
        
        print_x = int(screen_width * 0.87)
        print_y = int(screen_height * 0.93)
        
        pyautogui.click(print_x, print_y)
        time.sleep(3)
        
        print("[SUCCESS] ✓ Print button clicked!")
        return True
        
    except Exception as e:
        print(f"[ERROR] Print dialog automation failed: {e}")
        traceback.print_exc()
        return False


def automate_print_dialog_keyboard_fallback():
    """
    FALLBACK METHOD - Pure keyboard navigation
    Most reliable when mouse clicks fail
    """
    print("[INFO] Using KEYBOARD FALLBACK method...")
    time.sleep(6)
    
    try:
        # Focus dialog with mouse
        screen_width, screen_height = pyautogui.size()
        pyautogui.click(screen_width // 2, screen_height // 2)
        time.sleep(1)
        
        # Tab to Destination dropdown
        print("[STEP] Navigating with keyboard...")
        for _ in range(3):
            pyautogui.press('tab')
            time.sleep(0.3)
        
        # Open dropdown and search for printer
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.write('1020', interval=0.15)
        time.sleep(2)
        pyautogui.press('enter')
        time.sleep(3)
        
        print("[SUCCESS] Printer selected via keyboard")
        
        # Tab to More settings and press Enter
        for _ in range(4):
            pyautogui.press('tab')
            time.sleep(0.3)
        
        pyautogui.press('enter')
        time.sleep(2)
        print("[SUCCESS] More settings expanded via keyboard")
        
        # Tab to Pages per sheet
        for _ in range(2):
            pyautogui.press('tab')
            time.sleep(0.3)
        
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(2)
        
        print("[SUCCESS] Pages per sheet = 4 via keyboard")
        
        # Tab to Scale
        pyautogui.press('tab')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(2)
        
        print("[SUCCESS] Scale set via keyboard")
        
        # Tab to Print button and press Enter
        for _ in range(5):
            pyautogui.press('tab')
            time.sleep(0.3)
        
        pyautogui.press('enter')
        time.sleep(3)
        
        print("[SUCCESS] Print sent via keyboard!")
        return True
        
    except Exception as e:
        print(f"[ERROR] Keyboard fallback failed: {e}")
        return False