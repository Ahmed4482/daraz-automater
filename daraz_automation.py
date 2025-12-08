import time
import schedule
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException

# -----------------------------
# Credentials and configuration
# -----------------------------
EMAIL = "ahmedsemporium@gmail.com"
PASSWORD = "Ahm@d2003"

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
        
        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 20)

        # 1) Open initial page
        driver.get("https://sellercenter.daraz.pk/apps/seller/login")
        print("[STEP] Opened initial page")
        time.sleep(3)

        # 2) Click the initial "Log in" link (opens in new tab)
        try:
            initial_login_link = wait.until(
                EC.element_to_be_clickable((By.LINK_TEXT, "Log in"))
            )
            initial_login_link.click()
            print("[STEP] Clicked initial 'Log in' link")
            time.sleep(3)
            
            # Switch to the new tab that opened
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
                
                # Click "To Handover" tab
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
                
                # Now we're in To Handover tab, proceed with checkbox
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

                # Click "Print AWB" button
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
                
                # Wait for orders list to load
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.next-checkbox-input")))
                print("[STEP] Orders list loaded")

                # Click select-all checkbox
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

                # Click Pack & Print button
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

                # Modal: click Ship and Print or Print Only
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

        # Wait for print page
        wait.until(EC.url_contains("/apps/order/print"))
        print("[STEP] Print page loaded")
        time.sleep(5)

        # 10) Handle print dialog - Manual intervention needed
        print("\n" + "="*60)
        print("[IMPORTANT] MANUAL STEPS REQUIRED:")
        print("1. Click 'More settings' if available")
        print("2. Set 'Pages per sheet' to 4")
        print("3. Click 'Print' or 'Save as PDF'")
        print("="*60 + "\n")
        
        # Keep browser open for manual intervention
        print("[INFO] Keeping browser open for 60 seconds for manual steps...")
        print("[INFO] You can close the script anytime by pressing Ctrl+C")
        
        for i in range(60, 0, -10):
            print(f"[INFO] {i} seconds remaining...")
            time.sleep(10)
        
        print("[INFO] Flow completed.")

    except KeyboardInterrupt:
        print("\n[INFO] Script interrupted by user.")
    except Exception as e:
        print(f"[ERROR] Automation failed: {e}")
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
    # Uncomment for an immediate one-time run:
    run_flow()
    # schedule_job()