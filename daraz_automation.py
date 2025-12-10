# daraz_automation.py

import time
import schedule
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.keys import Keys
from print_dialog_automation import automate_print_dialog_keyboard_fallback 

# -----------------------------
# Credentials and configuration
# -----------------------------
EMAIL = "03332758016" # UPDATED EMAIL
PASSWORD = "T@ha2005" # UPDATED PASSWORD
PRINTER_NAME = "HP LaserJet 1020" # Your printer name

# PyAutoGUI settings
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

# -----------------------------
# Core automation flow
# -----------------------------

def run_flow():
    """Runs the Daraz Seller Center order print flow once."""
    print("[INFO] Starting Daraz Seller Center automation...")
    driver = None

    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        
        # Set HP LaserJet as default printer
        options.add_experimental_option("prefs", {
            "printing.print_preview_sticky_settings.appState": f'{{"recentDestinations":[{{"id":"{PRINTER_NAME}","origin":"local","account":""}}],"selectedDestinationId":"{PRINTER_NAME}","version":2}}',
            "savefile.default_directory": "C:\\Users\\Wajiz.pk\\Downloads"
        })
        
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 20)

        # 1) Open initial page
        driver.get("https://sellercenter.daraz.pk/apps/seller/login")
        print("[STEP 1] Opened initial page")
        time.sleep(3)

        # 2-6) Login and Navigate to Orders List 
        try:
            # Login link clicks and window switch (if applicable)
            try:
                initial_login_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Log in")))
                initial_login_link.click()
                print("[STEP 2] Clicked initial 'Log in' link")
                time.sleep(3)
                if len(driver.window_handles) > 1:
                    driver.switch_to.window(driver.window_handles[-1])
                    print("[STEP 2] Switched to new login tab")
                time.sleep(3)
            except:
                 try:
                    initial_login_link = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "to-login-link")))
                    initial_login_link.click()
                    if len(driver.window_handles) > 1:
                        driver.switch_to.window(driver.window_handles[-1])
                    time.sleep(3)
                 except:
                    print("[WARN] Initial login link not found, proceeding...")

            # Enter email
            email_field = wait.until(EC.presence_of_element_located((By.ID, "account")))
            email_field.send_keys(EMAIL)
            print("[STEP 3] Entered email")
            time.sleep(2)

            # Enter password
            password_field = wait.until(EC.presence_of_element_located((By.ID, "password")))
            password_field.send_keys(PASSWORD)
            print("[STEP 4] Entered password")
            time.sleep(2)

            # Click login button
            login_button = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@type='submit' and contains(@class, 'login-button')]")))
            driver.execute_script("arguments[0].click();", login_button)
            print("[STEP 5] Clicked login submit button")
            time.sleep(5)

            # Wait for home page
            wait.until(EC.url_contains("/apps/home"))
            print("[STEP 6] Logged in and on home page")
            time.sleep(3)
        except Exception as e:
            print(f"[ERROR] Login failed: {e}")
            raise

        # 7) Navigate to orders list
        driver.get("https://sellercenter.daraz.pk/apps/order/list")
        print("[STEP 7] Opened orders list")
        time.sleep(4)
        
        # Get window handles before print dialog *might* open
        window_handles_before = driver.window_handles


        # 8) Check if "To Pack" tab has empty data and proceed
        print("[DEBUG] Checking if 'To Pack' has orders...")
        
        # Check for 'Empty data' message first (more explicit check for no orders)
        has_orders = True
        try:
            # Look for the specific 'Empty data' div. We use a short timeout (5s) for this check.
            empty_data_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CLASS_NAME, "list-empty-content"))
            )
            # We also check the text content to be absolutely sure it's the "Empty data" message
            if empty_data_element.text.strip() == "Empty data":
                has_orders = False
                print("[DEBUG] Found 'Empty data' text. Setting has_orders = False.")
        except TimeoutException:
            # If the empty data element is not found in 5s, we assume orders are present.
            has_orders = True
            print("[DEBUG] 'Empty data' element not found (Timeout). Assuming orders are present.")

        try:
            if has_orders:
                # --- PRIMARY FLOW START: To Pack Has Orders (COMBINED FLOW) ---
                
                print("[INFO] 'To Pack' has orders. Proceeding with Pack & Print and full print sequence...")
                
                # 8.A.1 Select all orders in 'To Pack'
                select_all_checkbox = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input.next-checkbox-input[type='checkbox']"))
                )
                driver.execute_script("arguments[0].click();", select_all_checkbox)
                print("[STEP 8.A.1] Selected all pending orders in 'To Pack'")
                time.sleep(3)

                # 8.A.2 Click Pack & Print
                pack_print_button = wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//span[contains(text(), 'Pack') and contains(text(), 'Print')]/parent::button")
                    )
                )
                driver.execute_script("arguments[0].click();", pack_print_button)
                print("[STEP 8.A.2] Clicked Pack & Print")
                time.sleep(3)

                # 8.A.3 Handle Confirmation Pop-up (Click X button to close it)
                try:
                    # Targeting the close icon element
                    close_button = WebDriverWait(driver, 8).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.next-dialog-close[aria-label='Close']"))
                    )
                    driver.execute_script("arguments[0].click();", close_button)
                    print("[STEP 8.A.3] Clicked 'Close (X)' on confirmation dialog")
                except TimeoutException:
                    print("[WARN] Could not find 'Close (X)' button. Assuming dialog did not appear.")
                time.sleep(5)
                
                # --- The 'Pack & Print' is done, orders are now in 'To Arrange Shipment' ---

                # 8.A.4 Switch to 'To Arrange Shipment' tab
                to_arrange_tab = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'To Arrange Shipment')]"))
                )
                driver.execute_script("arguments[0].click();", to_arrange_tab)
                print("[STEP 8.A.4] Switched to 'To Arrange Shipment' tab")
                time.sleep(4)

                # 8.A.5 Select all orders in 'To Arrange Shipment'
                select_all_checkbox_ta = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input.next-checkbox-input[type='checkbox']"))
                )
                driver.execute_script("arguments[0].click();", select_all_checkbox_ta)
                print("[STEP 8.A.5] Selected all orders in 'To Arrange Shipment'")
                time.sleep(3)

                # 8.A.6 Click "Ready To Ship"
                ready_to_ship_button = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Ready To Ship')]/parent::button"))
                )
                driver.execute_script("arguments[0].click();", ready_to_ship_button)
                print("[STEP 8.A.6] Clicked 'Ready To Ship'")
                # *** INCREASED WAIT TIME FOR STABILITY AFTER STATUS CHANGE ***
                print("[DEBUG] Waiting 8s for status change and page refresh after 'Ready To Ship'...")
                time.sleep(8) 
                
                # --- Orders are now in 'To Handover' ---
                
                # 8.A.7 Switch to 'To Handover' tab
                print("[DEBUG] Attempting to find and click 'To Handover' tab...")
                to_handover_tab = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'To Handover')]"))
                )
                driver.execute_script("arguments[0].click();", to_handover_tab)
                print("[STEP 8.A.7] Switched to 'To Handover' tab")
                # *** INCREASED WAIT TIME FOR STABILITY AFTER TAB SWITCH ***
                print("[DEBUG] Waiting 6s for 'To Handover' content to load...")
                time.sleep(6)

                # 8.A.8 Select all orders in 'To Handover'
                print("[DEBUG] Checking for orders in 'To Handover' to select...")
                select_all_checkbox_ho = wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input.next-checkbox-input[type='checkbox']"))
                )
                driver.execute_script("arguments[0].click();", select_all_checkbox_ho)
                print("[STEP 8.A.8] Selected all orders in 'To Handover'")
                time.sleep(3)

                # 8.A.9 Click "Print AWB" button (This triggers the print window)
                print("[DEBUG] Looking for final 'Print AWB' button...")
                print_awb_button = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Print AWB')]/parent::button"))
                )
                driver.execute_script("arguments[0].click();", print_awb_button)
                print("[STEP 8.A.9] Clicked 'Print AWB' button (Initiates Print Window)")
                time.sleep(5) # Increased wait time

                # --- PRIMARY FLOW END (Jumps to Step 9) ---
            
            else:
                # --- SECONDARY FLOW START: To Pack is Empty (SKIP Arrange Shipment) ---
                
                print("[INFO] 'To Pack' is empty. Switching directly to 'To Handover' to catch ready orders...")
                
                # 8.B.0 Switch to 'To Handover' tab
                try:
                    to_handover_tab = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'To Handover')]"))
                    )
                    driver.execute_script("arguments[0].click();", to_handover_tab)
                    print("[STEP 8.B.0] Clicked 'To Handover' tab")
                    time.sleep(4)
                except TimeoutException:
                    print("[ERROR] Could not find 'To Handover' tab. Automation cannot proceed without print candidates.")
                    return
                
                # 8.B.1 Select all orders in 'To Handover'
                try:
                    select_all_checkbox_ho = wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "input.next-checkbox-input[type='checkbox']"))
                    )
                    driver.execute_script("arguments[0].click();", select_all_checkbox_ho)
                    print("[STEP 8.B.1] Selected all orders in 'To Handover'")
                except TimeoutException:
                    print("[WARN] No orders found in 'To Handover'. Ending job as nothing needs printing.")
                    return
                time.sleep(3)

                # 8.B.2 Click "Print AWB" button
                print_awb_button = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Print AWB')]/parent::button"))
                )
                driver.execute_script("arguments[0].click();", print_awb_button)
                print("[STEP 8.B.2] Clicked 'Print AWB' button (Initiates Print Window)")
                time.sleep(5)

                # --- SECONDARY FLOW END (Jumps to Step 9) ---

        except Exception as e:
            print(f"[ERROR] Error during order processing: {e}")
            import traceback
            traceback.print_exc()
            return

        # --- Print Flow Begins (Steps 9-11 - Same for both flows) ---

        # 9) Switch to print window
        # Wait until a new window/tab appears
        try:
            WebDriverWait(driver, 15).until(EC.number_of_windows_to_be(len(window_handles_before) + 1))
        except TimeoutException:
            print("[ERROR] Timeout waiting for new print window to open. Did the print action succeed?")
            return

        # Switch to the new window/tab
        driver.switch_to.window(driver.window_handles[-1])
        print("[STEP 9] Switched to print window")
        
        wait.until(EC.url_contains("/apps/order/print"))
        print("[STEP 9] Print page loaded")
        time.sleep(3)

        # 10) Click the "Print" button on Daraz page
        try:
            daraz_print_button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-spm='d_button_print']"))
            )
            driver.execute_script("arguments[0].click();", daraz_print_button)
            print("[STEP 10] Clicked Daraz 'Print' button - Chrome print dialog should open now")
            time.sleep(3)
        except TimeoutException:
            print("[ERROR] Could not find Daraz 'Print' button")
            return

        # 11) Automate Chrome print dialog with ONLY KEYBOARD FALLBACK
        print("=" * 60)
        print("[INFO] Starting Chrome print dialog automation...")
        print("[INFO] **Using ONLY KEYBOARD FALLBACK method**")
        print("=" * 60)
        
        success = automate_print_dialog_keyboard_fallback()
        
        if success:
            print("[SUCCESS] ✓✓✓ AUTOMATION COMPLETED SUCCESSFULLY! ✓✓✓")
        else:
            print("[ERROR] ❌ Automation failed to complete automatically")
        
        # BROWSER WILL STAY OPEN 
        print("\n" + "=" * 60)
        print("🔵 BROWSER WILL STAY OPEN FOR MANUAL INSPECTION")
        print("=" * 60)
        
        # Keep browser open indefinitely
        while True:
            time.sleep(10)

    except KeyboardInterrupt:
        print("\n[INFO] Script interrupted by user. Closing browser...")
    except Exception as e:
        print(f"[ERROR] Automation failed: Message: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if driver:
            try:
                driver.quit()
                print("[INFO] ✓ Browser closed.")
            except:
                pass


if __name__ == "__main__":
    print("=" * 60)
    print("DARAZ AUTOMATION WITH HP LASERJET 1020")
    print("=" * 60)
    run_flow()