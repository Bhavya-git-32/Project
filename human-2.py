import pandas as pd
import subprocess
import time
import os
import pyautogui
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, TimeoutException, NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.common.action_chains import ActionChains
import pyperclip
import random
import re
import tkinter as tk
from tkinter import messagebox

# === PERFORMANCE OPTIMIZATION & RELIABILITY CLASS ===
class WhatsAppOptimizer:
    def __init__(self):
        self.cached_xpaths = {}
        self.performance_stats = {
            'contacts_processed': 0,
            'start_time': time.time(),
            'search_times': [],
            'send_times': []
        }

    def smart_wait_for_element(self, driver, xpath_options, timeout=10):
        """More robustly waits for an element to be present and then clickable."""
        element = None
        start_time = time.time()
        for xpath in xpath_options:
            try:
                # First, wait for the element to be present in the DOM
                element = WebDriverWait(driver, timeout - (time.time() - start_time)).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                # Then, wait for the same element to be clickable
                element = WebDriverWait(driver, timeout - (time.time() - start_time)).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                if element:
                    return element
            except TimeoutException:
                continue
        raise TimeoutException(f"Element not found or not clickable after {timeout} seconds.")

    def ensure_search_ui_is_ready(self, driver):
        """
        Ensures the UI is clean before every search.
        This is critical because the UI can reset after sending a message.
        """
        try:
            # Close the "Turn on notifications" pop-up if it exists
            notification_close_button_xpath = '//div[@role="button" and @aria-label="Dismiss"]'
            close_button = WebDriverWait(driver, 1).until( # Faster check
                EC.element_to_be_clickable((By.XPATH, notification_close_button_xpath))
            )
            close_button.click()
        except TimeoutException:
            pass # Pop-up wasn't there, which is fine.

        try:
            # Click the search bar to dismiss the "Download for Windows" or other panes
            search_box_xpaths = [
                '//div[@title="Search input textbox"]',
                '//div[@role="textbox"][@title="Search input textbox"]'
            ]
            search_box = self.smart_wait_for_element(driver, search_box_xpaths, timeout=2)
            search_box.click()
            ActionChains(driver).send_keys(Keys.ESCAPE).perform() # Press escape to clear any focus
        except Exception:
            pass # If it fails, continue anyway.

    def search_contact(self, driver, contact):
        """Performs the search action for a contact."""
        search_box_xpaths = [
            '//div[@title="Search input textbox"]',
            '//div[@role="textbox"][@title="Search input textbox"]',
            '//div[contains(@class, "selectable-text")]//div[@contenteditable="true"]',
            '//div[@data-tab="3"][@contenteditable="true"]'
        ]
        search_box = self.smart_wait_for_element(driver, search_box_xpaths, timeout=8)

        # Clear the search box reliably
        search_box.click()
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).send_keys(Keys.BACKSPACE).perform()
        
        search_box.send_keys(contact)
        
        try:
            WebDriverWait(driver, 2).until(
                EC.presence_of_element_located((By.XPATH, '//div[@data-testid="chat-list-search-results"] | //div[contains(@class, "_ak_l")]'))
            )
        except TimeoutException:
            try:
                no_results_xpath = '//div[@data-testid="search-no-results-title"] | //span[contains(text(), "No results found for")]'
                WebDriverWait(driver, 1).until(EC.presence_of_element_located((By.XPATH, no_results_xpath)))
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                return 'not_found'
            except TimeoutException:
                pass

        search_box.send_keys(Keys.ENTER)
        return 'searched'


    def instant_message_send(self, driver, message, timeout=15):
        """
        Sends the whole message at once using the clipboard for speed and reliability.
        The timeout is now configurable.
        """
        send_start = time.time()
        
        message_box_xpaths = [
            '//footer//div[@contenteditable="true"]'
        ]
        
        message_box = self.smart_wait_for_element(driver, message_box_xpaths, timeout=timeout)
        
        pyperclip.copy(message)
        
        message_box.click()
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).send_keys(Keys.ENTER).perform()

        send_time = time.time() - send_start
        self.performance_stats['send_times'].append(send_time)
        return send_time

    def close_current_chat(self, driver):
        """
        Closes the currently open chat to reset the UI state.
        """
        try:
            back_button_xpath = '//div[@role="button" and @title="Back"] | //button[@aria-label="Back"]'
            back_button = self.smart_wait_for_element(driver, [back_button_xpath], timeout=3)
            back_button.click()
        except TimeoutException:
            try:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            except:
                pass

# --- HELPER FUNCTIONS ---
def format_number_for_api(number):
    """Formats a phone number to the correct 91XXXXXXXXXX format for the API URL."""
    # Remove any non-digit characters
    number = re.sub(r'\D', '', str(number))
    
    # If number starts with 91 and is 12 digits, strip the leading '+' if it exists
    if len(number) == 12 and number.startswith('91'):
        return number
    # If number is 10 digits, assume it's Indian and add 91
    elif len(number) == 10:
        return '91' + number
    # If it already starts with 91, assume it's correctly formatted
    elif number.startswith('91'):
        return number
    # Fallback for other cases
    else:
        return number

def ask_for_retry():
    """Displays a GUI pop-up to ask the user if they want to retry."""
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    response = messagebox.askyesno(
        "Retry Failed Contacts",
        "Do you want to retry the failed contacts using the direct link method?"
    )
    root.destroy()
    return response


# --- Configuration ---
edge_path = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
user_data_dir = r"D:\Auto send\edge_user_data_whatsapp"
remote_debugging_port = "9222"
msedgedriver_path = r"D:\Auto send\msedgedriver.exe"
contacts_excel_path = r"D:\Auto send\contacts.xlsx" 
message_txt_path = r"D:\Auto send\message.txt"
log_file_path = r"D:\Auto send\send_log_phase1_search.xlsx"
retry_log_file_path = r"D:\Auto send\send_log_phase2_retry.xlsx"


# Initialize optimizer
optimizer = WhatsAppOptimizer()

if not os.path.exists(user_data_dir):
    os.makedirs(user_data_dir)
    print(f"Created user data directory: {user_data_dir}")

# === OPTIMIZED Step 1: Fast Edge Launch ===
print(f"🚀 Starting Edge with performance optimizations...")
try:
    creationflags = 0
    if os.name == 'nt':
        creationflags = subprocess.DETACHED_PROCESS

    subprocess.Popen([
        edge_path,
        f'--remote-debugging-port={remote_debugging_port}',
        f'--user-data-dir={user_data_dir}',
        '--no-first-run',
        '--no-default-browser-check',
        '--disable-extensions',
        '--disable-plugins',
        '--disable-images',
        '--disable-javascript-harmony-shipping',
        '--disable-background-timer-throttling'
    ], creationflags=creationflags)

    for i in range(10):
        time.sleep(0.5)
        try:
            import socket
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            result = sock.connect_ex(('localhost', int(remote_debugging_port)))
            sock.close()
            if result == 0:
                print(f"✅ Edge ready in {(i+1)*0.5:.1f}s")
                break
        except:
            pass
    else:
        print("⏳ Edge taking longer than expected, continuing...")

except Exception as e:
    print(f"❌ Error launching Edge: {e}")
    exit()

# === OPTIMIZED Step 2: Fast Selenium Connection ===
print("🔌 Connecting Selenium with performance settings...")
try:
    options = webdriver.EdgeOptions()
    options.add_experimental_option("debuggerAddress", f"localhost:{remote_debugging_port}")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-sync")
    options.add_argument("--disable-translate")
    options.add_argument("--disable-ipc-flooding-protection")
    options.page_load_strategy = 'eager'

    service = Service(executable_path=msedgedriver_path)
    driver = webdriver.Edge(service=service, options=options)

    driver.set_page_load_timeout(80)
    driver.implicitly_wait(1)
    print("✅ Selenium connected with optimizations.")
except Exception as e:
    print(f"❌ Error connecting Selenium: {e}")
    exit()

# === Step 3: Load Data ===
try:
    df = pd.read_excel(contacts_excel_path, header=None, dtype={0: str})
    with open(message_txt_path, 'r', encoding='utf-8') as f:
        message = f.read().strip()
    print(f"✅ Loaded {len(df)} contacts from a single column file.")
except Exception as e:
    print(f"❌ Error loading data: {e}")
    if 'driver' in locals():
        driver.quit()
    exit()

# === OPTIMIZED Step 4: Smart WhatsApp Web Loading ===
print("🌐 Opening WhatsApp Web...")
driver.get("https://web.whatsapp.com")

print("⏳ Checking login status...")
try:
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.XPATH, '//canvas[@aria-label="Scan me!"] | //div[@aria-label="Chat list"]'))
    )
    try:
        driver.find_element(By.XPATH, '//div[@aria-label="Chat list"]')
        print("✅ Already logged in! Ready immediately.")
    except NoSuchElementException:
        print("⏳ Please scan the QR code with your phone to log in.")
        print("✅ QR Code is ready for scanning. You have 2 minutes to log in.")
        WebDriverWait(driver, 120).until(
            EC.presence_of_element_located((By.XPATH, '//div[@aria-label="Chat list"]'))
        )
        print("✅ Login successful.")
    time.sleep(1) # Minimal settling time
except TimeoutException:
    print("❌ WhatsApp Web did not load correctly. Please try running the script again.")
    driver.quit()
    exit()

# === PHASE 1: Search-Based Sending ===
print(f"\n⚡ Starting Phase 1: Sending via Search Method...")
log = []
start_processing = time.time()

# **NEW**: Shuffle the contact list before starting
contacts_list = df[0].dropna().tolist()
random.shuffle(contacts_list)
print(f"🔀 Shuffled {len(contacts_list)} contacts for random processing.")

for index, contact_raw in enumerate(contacts_list):
    contact = str(contact_raw).strip()
    if contact.endswith('.0'):
        contact = contact[:-2]
    
    contact_start_time = time.time()

    if not contact or contact == 'nan':
        print(f"[{index+1:3d}/{len(contacts_list)}] ⚠️ Skipping empty row.")
        continue

    print(f"[{index+1:3d}/{len(contacts_list)}] 📱 {contact:<20}", end="")

    try:
        optimizer.ensure_search_ui_is_ready(driver)
        search_result = optimizer.search_contact(driver, contact)
        
        if search_result == 'not_found':
            raise TimeoutException("Contact not found in search")

        send_time = optimizer.instant_message_send(driver, message, timeout=7)
        elapsed = time.time() - contact_start_time
        print(f" ✅ Sent ({elapsed:.2f}s)")
        log.append([contact, 'Success', f"{elapsed:.2f}s"])
        
        optimizer.close_current_chat(driver)

    except TimeoutException:
        elapsed = time.time() - contact_start_time
        print(f" ❌ Failed ({elapsed:.2f}s)")
        log.append([contact, 'Failed', f"{elapsed:.2f}s"])
        try:
            invalid_number_popup_xpath = '//div[contains(text(), "is not on WhatsApp")]'
            ok_button_xpath = '//div[@role="button" and text()="OK"]'
            WebDriverWait(driver, 2).until(EC.presence_of_element_located((By.XPATH, invalid_number_popup_xpath)))
            driver.find_element(By.XPATH, ok_button_xpath).click()
        except:
            try:
                ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            except:
                pass

    except Exception as e:
        elapsed = time.time() - contact_start_time
        print(f" ❌ Error ({elapsed:.2f}s) - {str(e)[:40]}...")
        log.append([contact, 'Failed - Script Error', f"{elapsed:.2f}s"])
        try:
            ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        except:
            pass

# === Save Phase 1 Log ===
log_df = pd.DataFrame(log, columns=["Contact", "Status", "Time"])
log_df.to_excel(log_file_path, index=False)
print(f"\n📄 Phase 1 log saved to {os.path.basename(log_file_path)}")

# === PHASE 2: Retry Failed Contacts via API URL ===
print("\n" + "="*50)
# Use the new GUI pop-up instead of console input
if ask_for_retry():
    print("User chose to retry. Starting Phase 2...")
    try:
        failed_df = log_df[log_df['Status'].str.contains('Failed')]
        if failed_df.empty:
            print("✅ No failed contacts to retry. All done!")
        else:
            failed_contacts = failed_df['Contact'].tolist()
            # **NEW**: Shuffle the failed contacts list as well
            random.shuffle(failed_contacts)
            print(f"🔀 Shuffled {len(failed_contacts)} failed contacts for random retry.")
            print(f"⚡ Starting Phase 2: Retrying {len(failed_contacts)} failed contacts via API URL...")
            retry_log = []

            for contact in failed_contacts:
                contact_start_time = time.time()
                print(f"📱 Retrying {contact:<20}", end="")
                
                formatted_number = format_number_for_api(contact)
                # **DEFINITIVE FIX**: Use the direct API URL to avoid pop-ups
                api_url = f"https://web.whatsapp.com/send?phone={formatted_number}&text&app_absent=0"
                
                try:
                    driver.get(api_url)
                    
                    # **SMARTER WAIT**: First, quickly check for the "invalid number" error.
                    try:
                        invalid_url_popup_xpath = '//div[contains(text(), "Phone number shared via url is invalid")]'
                        WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, invalid_url_popup_xpath)))
                        # If found, we know the number is bad. Raise an exception to fail fast.
                        raise Exception("Invalid phone number detected by URL")
                    except TimeoutException:
                        # This is the expected path for a valid number. Proceed to send.
                        pass

                    # Send the message with a patient timeout
                    send_time = optimizer.instant_message_send(driver, message, timeout=15)
                    elapsed = time.time() - contact_start_time
                    print(f" ✅ Sent ({elapsed:.2f}s)")
                    retry_log.append([contact, 'Success on Retry', f"{elapsed:.2f}s"])
                    
                    # Go back to the main screen to be ready for the next one
                    driver.get("https://web.whatsapp.com")
                    time.sleep(1)


                except Exception as e:
                    elapsed = time.time() - contact_start_time
                    print(f" ❌ Failed on Retry ({elapsed:.2f}s)")
                    retry_log.append([contact, 'Failed on Retry', f"{elapsed:.2f}s"])
                    
                    # Try to click the "OK" button on the invalid number pop-up if it exists
                    try:
                        ok_button_xpath = '//div[@role="button" and text()="OK"]'
                        driver.find_element(By.XPATH, ok_button_xpath).click()
                    except:
                        pass # No pop-up found, continue
            
            # Save Retry Log
            retry_log_df = pd.DataFrame(retry_log, columns=["Contact", "Status", "Time"])
            retry_log_df.to_excel(retry_log_file_path, index=False)
            print(f"\n📄 Phase 2 log saved to {os.path.basename(retry_log_file_path)}")

    except Exception as e:
        print(f"❌ An error occurred during the retry phase: {e}")

else:
    print("User chose not to retry. Skipping Phase 2.")


# === Final Exit ===
print("\n🔚 Cleaning up...")
try:
    driver.quit()
except:
    pass
print("✅ Automation complete! 🚀")
