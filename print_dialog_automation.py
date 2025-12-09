# print_dialog_automation.py

import time
import pyautogui
import traceback

# -----------------------------
# Core print dialog automation (Keyboard Fallback / Method 2)
# FIX: Printer selection steps removed. Now using 5 tabs to reach More settings.
# -----------------------------

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
        
        # 1. Click screen center to focus print dialog
        pyautogui.click(screen_width // 2, screen_height // 2)
        print("[STEP 1] Focused print dialog via center click.")
        time.sleep(1)
        
        # New Tab sequence to skip Destination, Pages, Copies, Layout:
        # Destination (1), Pages (2), Copies (3), Layout (4), More Settings (5)
        
        # 2. Tab to More settings and press Enter (5 tabs)
        print("[STEP 2] Navigating directly to More settings (5 tabs)...")
        for _ in range(4):
            pyautogui.press('tab')
            time.sleep(0.3)
        
        pyautogui.press('enter')
        time.sleep(2)
        print("[SUCCESS] ✓ More settings expanded")
        
        # After expanding, new elements appear:
        
        # 3. Tab to Paper Size (1 Tab after expand), select A4
        print("[STEP 3] Setting Paper Size to A4...")
        pyautogui.press('tab') # Tab 1: Paper size
        time.sleep(0.3)
        
        pyautogui.press('enter')
        time.sleep(1)
        
        # Select A4 (3 times down, yaani 4th option par A4 hota hai)
        for _ in range(3):
            pyautogui.press('down')
            time.sleep(0.3)
            
        pyautogui.press('enter')
        time.sleep(2)
        print("[SUCCESS] ✓ Paper Size = A4")
        
        # 4. Tab to Pages per sheet (2nd dropdown now), set to 4
        print("[STEP 4] Setting Pages per sheet to 4...")
        pyautogui.press('tab') # Tab 2: Pages per sheet
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(1)
        
        # Select 4 pages per sheet (2 times down)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(2)
        
        print("[SUCCESS] ✓ Pages per sheet = 4")
        
        # 5. Tab to Scale (3rd dropdown now), set to Fit to printable area
        print("[STEP 5] Setting Scale to Fit to printable area...")
        pyautogui.press('tab') # Tab 3: Scale
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(1)
        
        # Select Fit to printable area (1 time down)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(2)
        
        print("[SUCCESS] ✓ Scale = Fit to printable area")
        
        # 6. Tab to Print button and press Enter
        # From Scale, it usually takes 4 tabs to reach Print button
        print("[STEP 6] Navigating to Print button...")
        for _ in range(4):
            pyautogui.press('tab')
            time.sleep(0.3)
        
        pyautogui.press('enter')
        time.sleep(3)
        
        print("[SUCCESS] ✓ Print sent via keyboard!")
        return True
        
    except Exception as e:
        print(f"[ERROR] Keyboard fallback failed: {e}")
        traceback.print_exc()
        return False