import time
import schedule
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.keys import Keys

# -----------------------------
# Credentials and configuration
# -----------------------------
EMAIL = "ahmedsemporium@gmail.com"
PASSWORD = "Ahm@d2003"
PRINTER_NAME = "HP LaserJet 1020"  # Your printer name

# PyAutoGUI settings
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5

# -----------------------------
# Fixed print dialog automation with HP LaserJet
# -----------------------------

def automate_print_dialog_hp():
    """
    Automates Chrome print dialog for HP LaserJet using MOUSE CLICKS:
    1. Clicks on Destination dropdown and selects HP LaserJet
    2. Clicks on More settings button to expand
    3. Clicks on Pages per sheet dropdown and selects 4
    4. Clicks on Scale dropdown and selects Fit to printable area
    5. Clicks Print button
    """
    print("[INFO] Waiting for print dialog to open...")
    time.sleep(5)  # Wait for dialog to fully load
    
    try:
        # Get screen size for reference
        screen_width, screen_height = pyautogui.size()
        print(f"[DEBUG] Screen resolution: {screen_width}x{screen_height}")
        
        # Step 1: Click on Destination dropdown (right side of screen, near top)
        print("[STEP 1] Clicking on Destination dropdown...")
        # Destination is usually at right side, about 1/4 from top
        dest_x = int(screen_width * 0.85)  # 85% from left
        dest_y = int(screen_height * 0.20)  # 20% from top
        
        pyautogui.click(dest_x, dest_y)
        time.sleep(1)
        
        # Type to search for HP printer
        print(f"[DEBUG] Searching for printer: {PRINTER_NAME}")
        pyautogui.write('hp laser 1020')
        time.sleep(1.5)
        
        # Press Enter to select
        pyautogui.press('enter')
        time.sleep(2)
        
        print("[SUCCESS] HP LaserJet printer selected")
        
        # Step 2: Click on "More settings" button
        print("[STEP 2] Clicking on More settings button...")
        # More settings button is usually in middle-right area, below destination
        more_settings_x = int(screen_width * 0.85)
        more_settings_y = int(screen_height * 0.48)  # Adjust based on your screen
        
        pyautogui.click(more_settings_x, more_settings_y)
        time.sleep(2)
        
        print("[SUCCESS] More settings expanded")
        
        # Step 3: Click on "Pages per sheet" dropdown
        print("[STEP 3] Clicking on Pages per sheet dropdown...")
        pages_per_sheet_x = int(screen_width * 0.85)
        pages_per_sheet_y = int(screen_height * 0.58)  # Below More settings
        
        pyautogui.click(pages_per_sheet_x, pages_per_sheet_y)
        time.sleep(0.8)
        
        # Select 4 using keyboard (more reliable)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(1)
        
        print("[SUCCESS] Pages per sheet set to 4")
        
        # Step 4: Click on "Scale" dropdown
        print("[STEP 4] Clicking on Scale dropdown...")
        scale_x = int(screen_width * 0.85)
        scale_y = int(screen_height * 0.65)  # Below Pages per sheet
        
        pyautogui.click(scale_x, scale_y)
        time.sleep(0.8)
        
        # Select "Fit to printable area"
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(1)
        
        print("[SUCCESS] Scale set to Fit to printable area")
        
        # Step 5: Click on Print button (bottom right, blue button)
        print("[STEP 5] Clicking Print button...")
        print_button_x = int(screen_width * 0.88)  # Far right
        print_button_y = int(screen_height * 0.92)  # Near bottom
        
        pyautogui.click(print_button_x, print_button_y)
        time.sleep(2)
        
        print("[SUCCESS] Print dialog automated successfully!")
        return True
        
    except Exception as e:
        print(f"[ERROR] Failed to automate print dialog: {e}")
        return False


def automate_print_dialog_hp_method2():
    """
    Method 2: Using more accurate mouse coordinates
    Clicks directly on visible elements
    """
    print("[INFO] Using Method 2 - Accurate mouse clicks...")
    time.sleep(5)
    
    try:
        screen_width, screen_height = pyautogui.size()
        
        # Click anywhere on dialog to ensure it's focused
        pyautogui.click(screen_width // 2, screen_height // 2)
        time.sleep(0.5)
        
        # Step 1: Destination dropdown
        print("[STEP] Clicking Destination dropdown...")
        # Right side panel, first dropdown
        pyautogui.click(int(screen_width * 0.88), int(screen_height * 0.18))
        time.sleep(1)
        
        # Type and select printer
        pyautogui.write('1020')
        time.sleep(1.5)
        pyautogui.press('enter')
        time.sleep(2)
        
        print("[SUCCESS] Printer selected")
        
        # Step 2: More settings button (it's a button with text)
        print("[STEP] Clicking More settings...")
        # Usually around middle-right
        pyautogui.click(int(screen_width * 0.82), int(screen_height * 0.48))
        time.sleep(2)
        
        # Step 3: Pages per sheet
        print("[STEP] Setting Pages per sheet...")
        pyautogui.click(int(screen_width * 0.88), int(screen_height * 0.58))
        time.sleep(0.8)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(1)
        
        # Step 4: Scale
        print("[STEP] Setting Scale...")
        pyautogui.click(int(screen_width * 0.88), int(screen_height * 0.66))
        time.sleep(0.8)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('down')
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(1)
        
        # Step 5: Print button (blue button at bottom right)
        print("[STEP] Clicking Print button...")
        pyautogui.click(int(screen_width * 0.87), int(screen_height * 0.93))
        time.sleep(2)
        
        print("[SUCCESS] Method 2 completed!")
        return True
        
    except Exception as e:
        print(f"[ERROR] Method 2 failed: {e}")
        return False


def automate_print_dialog_hp_method3():
    """
    Method 3: Hybrid approach - Mouse for buttons, keyboard for selections
    Most reliable method
    """
    print("[INFO] Using Method 3 (Hybrid Mouse + Keyboard)...")
    time.sleep(6)
    
    try:
        screen_width, screen_height = pyautogui.size()
        print(f"[DEBUG] Screen: {screen_width}x{screen_height}")
        
        # Click on dialog to focus it
        pyautogui.click(screen_width // 2, screen_height // 2)
        time.sleep(0.8)
        
        # Step 1: Destination
        print("[ACTION] Clicking Destination...")
        # Try multiple possible positions for destination dropdown
        positions = [
            (int(screen_width * 0.85), int(screen_height * 0.17)),
            (int(screen_width * 0.87), int(screen_height * 0.19)),
            (int(screen_width * 0.90), int(screen_height * 0.18))
        ]
        
        for pos_x, pos_y in positions:
            pyautogui.click(pos_x, pos_y)
            time.sleep(0.5)
        
        # Type to search
        time.sleep(0.5)
        pyautogui.write('1020')
        time.sleep(1.5)
        pyautogui.press('enter')
        time.sleep(2)
        
        print("[ACTION] Clicking More settings button...")
        # More settings is a text button, try multiple positions
        more_settings_positions = [
            (int(screen_width * 0.80), int(screen_height * 0.48)),
            (int(screen_width * 0.83), int(screen_height * 0.47)),
            (int(screen_width * 0.78), int(screen_height * 0.49))
        ]
        
        for pos_x, pos_y in more_settings_positions:
            pyautogui.click(pos_x, pos_y)
            time.sleep(0.8)
        
        time.sleep(1.5)
        
        print("[ACTION] Setting Pages per sheet...")
        # Click Pages per sheet dropdown
        pyautogui.click(int(screen_width * 0.88), int(screen_height * 0.57))
        time.sleep(1)
        
        # Select 4
        pyautogui.press('down')
        time.sleep(0.4)
        pyautogui.press('down')
        time.sleep(0.4)
        pyautogui.press('enter')
        time.sleep(1.5)
        
        print("[ACTION] Setting Scale...")
        # Click Scale dropdown
        pyautogui.click(int(screen_width * 0.88), int(screen_height * 0.65))
        time.sleep(1)
        
        # Select Fit to printable area
        pyautogui.press('down')
        time.sleep(0.4)
        pyautogui.press('down')
        time.sleep(0.4)
        pyautogui.press('enter')
        time.sleep(1.5)
        
        print("[ACTION] Clicking Print button...")
        # Click Print button - bottom right corner
        print_button_positions = [
            (int(screen_width * 0.87), int(screen_height * 0.93)),
            (int(screen_width * 0.85), int(screen_height * 0.92)),
            (int(screen_width * 0.89), int(screen_height * 0.94))
        ]
        
        # Try clicking multiple times in different positions
        for pos_x, pos_y in print_button_positions:
            pyautogui.click(pos_x, pos_y)
            time.sleep(0.5)
        
        time.sleep(2)
        
        print("[SUCCESS] Method 3 completed!")
        return True
        
    except Exception as e:
        print(f"[ERROR] Method 3 failed: {e}")
        return False


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
        print("[STEP] Opened initial page")
        time.sleep(3)

        # 2) Click the initial "Log in" link
        try:
            initial_login_link = wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Log in"))
            )
            initial_login_link.click()
            print("[STEP] Clicked initial 'Log in' link")
            time.sleep(3)
            
            if len(driver.window_handles) > 1:
                driver.switch_to.window(driver.window_handles[-1])
                print("[STEP] Switched to new login tab")
            time.sleep(3)
            
        except TimeoutException:
            try:
                initial_login_link = wait.until(
                    EC.element_to_be_clickable((By.CLASS_NAME, "to-login-link"))
                )
                initial_login_link.click()
                print("[STEP] Clicked initial 'Log in' link (via class)")
                time.sleep(3)
                
                if len(driver.window_handles) > 1:
                    driver.switch_to.window(driver.window_handles[-1])
                    print("[STEP] Switched to new login tab")
                time.sleep(3)
                
            except TimeoutException:
                print("[WARN] Initial login link not found, proceeding...")

        print(f"[DEBUG] Current URL: {driver.current_url}")
        
        # 3) Enter email
        try:
            print("[DEBUG] Waiting for account field...")
            email_field = wait.until(
                EC.presence_of_element_located((By.ID, "account"))
            )
            print("[DEBUG] Found email field")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", email_field)
            time.sleep(2)
            driver.execute_script("arguments[0].click();", email_field)
            time.sleep(1)
            email_field.clear()
            email_field.send_keys(EMAIL)
            print("[STEP] Entered email")
            time.sleep(2)
        except TimeoutException:
            print(f"[ERROR] Could not find email field. Current URL: {driver.current_url}")
            driver.save_screenshot("error_screenshot.png")
            raise

        # 4) Enter password
        try:
            print("[DEBUG] Waiting for password field...")
            password_field = wait.until(
                EC.presence_of_element_located((By.ID, "password"))
            )
            print("[DEBUG] Found password field")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", password_field)
            time.sleep(2)
            driver.execute_script("arguments[0].click();", password_field)
            time.sleep(1)
            password_field.clear()
            password_field.send_keys(PASSWORD)
            print("[STEP] Entered password")
            time.sleep(2)
        except TimeoutException:
            print(f"[ERROR] Could not find password field.")
            driver.save_screenshot("error_screenshot_password.png")
            raise

        # 5) Click login button
        try:
            print("[DEBUG] Looking for Login button...")
            login_button = wait.until(
                EC.presence_of_element_located((By.XPATH, "//button[@type='submit' and contains(@class, 'login-button')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", login_button)
            time.sleep(2)
            driver.execute_script("arguments[0].click();", login_button)
            print("[STEP] Clicked login submit button")
            time.sleep(5)
        except TimeoutException:
            print("[ERROR] Could not find login submit button")
            raise

        # 6) Wait for home page
        wait.until(EC.url_contains("/apps/home"))
        print("[STEP] Logged in and on home page")
        time.sleep(3)

        # 7) Navigate to orders list
        driver.get("https://sellercenter.daraz.pk/apps/order/list")
        print("[STEP] Opened orders list")
        time.sleep(4)

        # 8) Check if "To Pack" tab has empty data
        print("[DEBUG] Checking if 'To Pack' has orders...")
        try:
            empty_data = driver.find_elements(By.CLASS_NAME, "list-empty-content")
            if len(empty_data) > 0 and empty_data[0].is_displayed():
                print("[INFO] 'To Pack' is empty. Switching to 'To Handover' tab...")
                
                try:
                    to_handover_tab = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'To Handover')]"))
                    )
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", to_handover_tab)
                    time.sleep(2)
                    driver.execute_script("arguments[0].click();", to_handover_tab)
                    print("[STEP] Clicked 'To Handover' tab")
                    time.sleep(4)
                except TimeoutException:
                    print("[ERROR] Could not find 'To Handover' tab")
                    return
                
                print("[DEBUG] Looking for select-all checkbox in 'To Handover'...")
                try:
                    select_all_checkbox = wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "input.next-checkbox-input[type='checkbox']"))
                    )
                    print("[DEBUG] Found checkbox")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_all_checkbox)
                    time.sleep(2)
                    driver.execute_script("arguments[0].click();", select_all_checkbox)
                    print("[STEP] Selected all orders in 'To Handover'")
                except TimeoutException:
                    print("[ERROR] Could not find checkbox in 'To Handover'")
                    return
                time.sleep(3)

                try:
                    print("[DEBUG] Looking for 'Print AWB' button...")
                    print_awb_button = wait.until(
                        EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Print AWB')]/parent::button"))
                    )
                    print("[DEBUG] Found 'Print AWB' button")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", print_awb_button)
                    time.sleep(2)
                    driver.execute_script("arguments[0].click();", print_awb_button)
                    print("[STEP] Clicked 'Print AWB' button")
                except TimeoutException:
                    print("[ERROR] Could not find 'Print AWB' button")
                    return
                time.sleep(3)

            else:
                print("[INFO] 'To Pack' has orders. Proceeding with Pack & Print...")
                
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.next-checkbox-input")))
                print("[STEP] Orders list loaded")

                try:
                    print("[DEBUG] Looking for select-all checkbox...")
                    select_all_checkbox = wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "input.next-checkbox-input[type='checkbox']"))
                    )
                    print("[DEBUG] Found checkbox")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_all_checkbox)
                    time.sleep(2)
                    driver.execute_script("arguments[0].click();", select_all_checkbox)
                    print("[STEP] Selected all pending orders")
                except TimeoutException:
                    print("[ERROR] Could not find the select-all checkbox.")
                    driver.save_screenshot("error_checkbox.png")
                    return
                time.sleep(3)

                try:
                    print("[DEBUG] Looking for Pack & Print button...")
                    pack_print_button = wait.until(
                        EC.presence_of_element_located(
                            (By.XPATH, "//span[contains(text(), 'Pack') and contains(text(), 'Print')]/parent::button")
                        )
                    )
                    print("[DEBUG] Found Pack & Print button")
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", pack_print_button)
                    time.sleep(2)
                    driver.execute_script("arguments[0].click();", pack_print_button)
                    print("[STEP] Clicked Pack & Print")
                except TimeoutException:
                    print("[ERROR] Could not find Pack & Print button.")
                    driver.save_screenshot("error_pack_print.png")
                    return
                time.sleep(3)

                try:
                    ship_and_print = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Ship and Print')]"))
                    )
                    ship_and_print.click()
                    print("[STEP] Clicked Ship and Print")
                except TimeoutException:
                    try:
                        print_only = WebDriverWait(driver, 8).until(
                            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Print Only')]"))
                        )
                        print_only.click()
                        print("[STEP] Clicked Print Only")
                    except TimeoutException:
                        print("[ERROR] Neither 'Ship and Print' nor 'Print Only' button was found.")
                        return
                time.sleep(5)

        except Exception as e:
            print(f"[ERROR] Error during order checking: {e}")
            return

        # 9) Switch to print window
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])
            print("[STEP] Switched to print window")
        else:
            print("[WARN] Print window did not open automatically. Stopping.")
            return

        wait.until(EC.url_contains("/apps/order/print"))
        print("[STEP] Print page loaded")
        time.sleep(3)

        # 10) Click the "Print" button on Daraz page
        try:
            print("[DEBUG] Looking for Daraz 'Print' button...")
            daraz_print_button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-spm='d_button_print']"))
            )
            print("[DEBUG] Found Daraz 'Print' button")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", daraz_print_button)
            time.sleep(2)
            driver.execute_script("arguments[0].click();", daraz_print_button)
            print("[STEP] Clicked Daraz 'Print' button - Chrome print dialog should open now")
            time.sleep(3)
        except TimeoutException:
            print("[ERROR] Could not find Daraz 'Print' button")
            driver.save_screenshot("error_daraz_print.png")
            return

        # 11) Automate Chrome print dialog with HP LaserJet
        print("=" * 60)
        print("[INFO] Starting Chrome print dialog automation for HP LaserJet...")
        print(f"[INFO] Target printer: {PRINTER_NAME}")
        print("=" * 60)
        
        success = automate_print_dialog_hp()
        
        if not success:
            print("[WARN] Method 1 failed, trying Method 2...")
            success = automate_print_dialog_hp_method2()
        
        if not success:
            print("[WARN] Method 2 failed, trying Method 3...")
            success = automate_print_dialog_hp_method3()
        
        if success:
            print("=" * 60)
            print("[SUCCESS] ✓ Automation completed successfully!")
            print("[SUCCESS] ✓ Print job sent to HP LaserJet")
            print("=" * 60)
            time.sleep(5)
        else:
            print("=" * 60)
            print("[ERROR] All automated methods failed.")
            print("[INFO] Keeping browser open for manual intervention...")
            print("[INFO] Please manually complete these steps:")
            print("  1. Select 'HP LaserJet M1005 MFP' from Destination")
            print("  2. Click 'More settings' to expand")
            print("  3. Set 'Pages per sheet' to '4'")
            print("  4. Set 'Scale' to 'Fit to printable area'")
            print("  5. Click the 'Print' button (NOT Cancel)")
            print("=" * 60)
            time.sleep(120)  # Wait 2 minutes for manual intervention

    except KeyboardInterrupt:
        print("\n[INFO] Script interrupted by user.")
    except Exception as e:
        print(f"[ERROR] Automation failed: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if driver:
            print("[INFO] Closing browser...")
            try:
                driver.quit()
            except:
                pass
            print("[INFO] Browser closed.")


def schedule_job():
    """Schedules the flow to run daily at 11:59 PM."""
    print("[INFO] Scheduling daily job for 23:59...")
    schedule.every().day.at("23:59").do(run_flow)
    while True:
        schedule.run_pending()
        time.sleep(30)


if __name__ == "__main__":
    print("=" * 60)
    print("DARAZ AUTOMATION WITH HP LASERJET PRINTER")
    print("=" * 60)
    print(f"Target Printer: {PRINTER_NAME}")
    print("Settings: 4 pages per sheet, Fit to printable area")
    print("=" * 60)
    print()
    
    # Uncomment for an immediate one-time run:
    run_flow()
    
    # Uncomment to schedule daily at 11:59 PM:
    # schedule_job()