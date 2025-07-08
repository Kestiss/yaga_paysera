import time
import math
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ---- Configuration ----
HEADLESS = False  # Set to False to see the browser
EMAIL = "email@email.com"
PASSWORD = "password"
GMAIL_LOGIN_URL = "https://accounts.google.com/v3/signin/identifier?continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&hl=en&service=mail&flowName=GlifWebSignIn&flowEntry=AddSession"
PAYSERA_ADMIN_URL = "https://tickets.paysera.com/lt-LT/admin/login"
DASHBOARD_URL = "https://tickets.paysera.com/lt/my/self-service/events"
GSHEET_ID = "1vRUvZ5sftDG6aZkK0Nq9aztgLM6QVzHwPtg-QN1Ad0g"
GSHEET_TAB = "MARKETING25"
CREDENTIALS_FILE = "credentials.json"
TICKET_COUNT_HEADER = "MAIN PARKING"
WRISTBAND_HEADER = "Number of WRIST BANDS"
CAMPER_PASS_HEADER = "CAMPER PASS"
FULLPASS_LABEL = "FULL PASS ★"
MAINPARKING_LABEL = "MAIN PARKING ★"
CAMPERPASS_LABEL  = "CAMPER PASS + ELECTRICITY ★"
INV_SENT_HEADER = "Invitations SENT"
PARKING_SENT_HEADER = "Main parking SENT"
CAMPER_SENT_HEADER = "Camper SENT"

# ---- Selenium Utility Functions ----
def wait_for_url(url_fragment, timeout=30):
    WebDriverWait(driver, timeout).until(EC.url_contains(url_fragment))

def wait_for_xpath(xpath, visible=False, timeout=45):
    if visible:
        return WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((By.XPATH, xpath)))
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))

def wait_ready_state(timeout=30):
    WebDriverWait(driver, timeout).until(lambda d: d.execute_script('return document.readyState') == 'complete')

def click_and_tab_handling(element):
    old_tabs = driver.window_handles.copy()
    element.click()
    # Wait for possible new tab up to 7 sec
    for _ in range(14):
        if len(driver.window_handles) > len(old_tabs):
            new_tabs = driver.window_handles
            driver.switch_to.window(list(set(new_tabs) - set(old_tabs))[0])
            return True, old_tabs[0]
        time.sleep(0.5)
    # No new tab opened
    return False, old_tabs[0]

def close_optional_tab(switched, main_tab):
    if switched:
        try:
            driver.close()
        except: pass
        try:
            driver.switch_to.window(main_tab)
        except: pass

def fill_person_fields(name_cell):
    parts = name_cell.strip().split()
    first_name = parts[0] if parts else ""
    surname = " ".join(parts[1:]) if len(parts) > 1 else ""
    try:
        el = driver.find_element(By.XPATH, "//label[contains(text(),'Vardas')]/../input")
        el.clear()
        el.send_keys(first_name)
    except Exception: pass
    try:
        el = driver.find_element(By.XPATH, "//label[contains(text(),'Pavardė')]/../input")
        el.clear()
        el.send_keys(surname)
    except Exception: pass

def fill_email(email_cell):
    try:
        el = driver.find_element(By.CSS_SELECTOR, "input[type='email']")
        el.clear()
        el.send_keys(email_cell)
    except Exception: pass

# ---- Chrome Setup ----
chrome_options = Options()
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-infobars")
chrome_options.add_argument("--disable-dev-shm-usage")
if HEADLESS:
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

try:
    # ------ Google Login ------
    driver.get(GMAIL_LOGIN_URL)
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.ID, "identifierId"))).send_keys(EMAIL)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Next']"))).click()

    # Wait for password field
    WebDriverWait(driver, 30).until(EC.visibility_of_element_located(
        (By.CSS_SELECTOR, "input[type='password']"))).send_keys(PASSWORD)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Next']"))).click()

    wait_for_url("mail.google.com")
    print("Gmail login successful.")

    # ------ Paysera Login ------
    driver.get(PAYSERA_ADMIN_URL)
    time.sleep(5)
    WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "[class*='btn-google']"))).click()
    wait_for_xpath("//div[contains(text(),'Yaga Gathering')]")
    driver.execute_script('let el=document.querySelector("chrome-signin-app");if(el)el.remove();')
    print("Logged in to Paysera.")

    # ------ Google Sheet Data ------
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    worksheet = gspread.authorize(creds).open_by_key(GSHEET_ID).worksheet(GSHEET_TAB)
    rows = worksheet.get_all_values()
    headers = rows[1]
    col_index_parkcount = headers.index(TICKET_COUNT_HEADER)
    col_index_wristbands = headers.index(WRISTBAND_HEADER)
    col_index_camper_pass = headers.index(CAMPER_PASS_HEADER)
    col_index_active = headers.index("-")
    col_index_email = headers.index("E-MAIL")
    col_index_name = headers.index("NAME SURNAME")
    col_index_inv_sent = headers.index(INV_SENT_HEADER)
    col_index_parking_sent = headers.index(PARKING_SENT_HEADER)
    col_index_camper_sent = headers.index(CAMPER_SENT_HEADER)

    # ------ Main Automation Loop ------
    any_processed = False
    for i, row in enumerate(rows[2:], start=3):  # Data starts at 3rd row in sheet

        # 1. Skip if entire row is empty
        if all(cell.strip() == "" for cell in row):
            continue

        # 2. Stop loop if NAME SURNAME is THE END
        if row[col_index_name].strip() == "THE END":
            print(f"Found THE END at row {i}. Stopping script.")
            break

        # 3. Skip already processed rows (any SENT cell not empty)
        if (
            row[col_index_inv_sent].strip() or
            row[col_index_parking_sent].strip() or
            row[col_index_camper_sent].strip()
        ):
            continue

        if len(row) <= max(
            col_index_parkcount, col_index_wristbands, col_index_camper_pass, col_index_active,
            col_index_email, col_index_name, col_index_inv_sent,
            col_index_parking_sent, col_index_camper_sent
        ):
            continue

        is_active = row[col_index_active].strip() != "-"
        tickets_cell = row[col_index_parkcount].strip()
        wristbands_cell = row[col_index_wristbands].strip()
        camper_pass_cell = row[col_index_camper_pass].strip() if len(row) > col_index_camper_pass else ""
        has_camper_pass = bool(camper_pass_cell)

        try:
            main_parking_count = float(tickets_cell.replace(',', '.'))
        except:
            main_parking_count = 0

        try:
            wristbands_count = float(wristbands_cell.replace(',', '.'))
        except:
            wristbands_count = 0

        has_any_ticket = (main_parking_count > 0) or (wristbands_count > 0) or has_camper_pass

        email_cell = row[col_index_email].strip()
        name_cell = row[col_index_name].strip()

        if not (is_active and has_any_ticket and email_cell and name_cell):
            continue

        print(f"\nRow {i}: Processing {name_cell} / {email_cell} / Parking: {main_parking_count} / Wristbands: {wristbands_count} / Camper: {has_camper_pass}")

        # --- Support for multiple submissions if wristbands or main parking >10 ---
        wristbands_total = int(wristbands_count)
        parking_total = int(main_parking_count)
        camper_total = 1 if has_camper_pass else 0

        num_wristband_loops = math.ceil(wristbands_total / 10) if wristbands_total > 0 else 0
        num_parking_loops   = math.ceil(parking_total   / 10) if parking_total   > 0 else 0
        max_loops           = max(num_wristband_loops, num_parking_loops, camper_total)

        remaining_wristbands = wristbands_total
        remaining_parking = parking_total
        remaining_camper = camper_total

        for loop_idx in range(max_loops):
            wb_this_loop   = min(remaining_wristbands, 10)
            park_this_loop = min(remaining_parking, 10)
            camper_this_loop = 1 if (loop_idx == 0 and has_camper_pass) else 0

            # skip submissions that would do nothing (shouldn't happen but just in case)
            if wb_this_loop == 0 and park_this_loop == 0 and camper_this_loop == 0:
                break

            print(f"   Submission {loop_idx+1}: WRISTBANDS={wb_this_loop}, PARKING={park_this_loop}, CAMPER={camper_this_loop}")

            driver.get(DASHBOARD_URL)
            wait_for_xpath("//div[contains(text(),'Yaga Gathering')]")
            driver.execute_script('let el=document.querySelector("chrome-signin-app");if(el)el.remove();')

            # --- Click "Kurti užsakymą", handle tab ---
            switched = False
            main_tab = None
            try:
                order_btn = WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, '[class*="MuiCardActions-root"] a[title="Kurti užsakymą"][href*="yaga"]')
                    )
                )
                switched, main_tab = click_and_tab_handling(order_btn)
                wait_ready_state()
                time.sleep(3)

                # --- Wait for Event Page ---
                wait_for_xpath("//bdo[contains(text(),'Yaga Gathering 2025')]", visible=True)

                # --- MAIN PARKING ★ selection (if needed) ---
                if park_this_loop > 0:
                    row_elem = driver.find_element(
                        By.XPATH,
                        f"//div[contains(@class,'ticket-type-row')]//span[contains(text(),'{MAINPARKING_LABEL}')]//ancestor::div[contains(@class,'ticket-type-row')]"
                    )
                    select = Select(row_elem.find_element(By.CSS_SELECTOR, "select.ticket-count-box"))
                    select.select_by_visible_text(str(park_this_loop))

                # --- FULL PASS ★ (wristband) selection (if needed) ---
                if wb_this_loop > 0:
                    fullpass_row = driver.find_element(
                        By.XPATH,
                        f"//div[contains(@class,'ticket-type-row')]//span[contains(text(),'{FULLPASS_LABEL}')]//ancestor::div[contains(@class,'ticket-type-row')]"
                    )
                    fullpass_select = Select(fullpass_row.find_element(By.CSS_SELECTOR, "select.ticket-count-box"))
                    fullpass_select.select_by_visible_text(str(wb_this_loop))

                # --- CAMPER PASS + ELECTRICITY selection (if needed) ---
                if camper_this_loop > 0:
                    camper_row = driver.find_element(
                        By.XPATH,
                        f"//div[contains(@class,'ticket-type-row')]//span[contains(text(),'{CAMPERPASS_LABEL}')]//ancestor::div[contains(@class,'ticket-type-row')]"
                    )
                    camper_select = Select(camper_row.find_element(By.CSS_SELECTOR, "select.ticket-count-box"))
                    camper_select.select_by_visible_text("1")

                # --- Scroll to and click the "Tęsti" button ---
                continue_btn = WebDriverWait(driver, 30).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Tęsti')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", continue_btn)
                time.sleep(0.3)
                continue_btn.click()

                wait_for_xpath("//*[contains(text(),'Pateikite informaciją')]")

                try:
                    driver.find_element(By.CSS_SELECTOR, "span.clear-icon").click()
                except Exception:
                    pass

                # --- FILL EMAIL, NAME, CHECKBOXES ---
                fill_email(email_cell)
                fill_person_fields(name_cell)

                try:
                    checkbox_ids = [
                        "user_details_fill_form_records_order_fields_113534_0",  # 1st
                        "user_details_fill_form_records_order_fields_113535_0",  # 2nd
                        "user_details_fill_form_serviceAgreement",               # 3rd
                        "user_details_fill_form_upcomingEventsSubscription",     # 4th
                    ]

                    for cid in checkbox_ids:
                        try:
                            el = driver.find_element(By.ID, cid)
                            # Only click if not selected
                            if not el.is_selected():
                                driver.execute_script(
                                    "arguments[0].checked = true;"
                                    "arguments[0].dispatchEvent(new Event('input', {bubbles: true}));"
                                    "arguments[0].dispatchEvent(new Event('change', {bubbles: true}));"
                                    , el
                                )
                        except Exception as e:
                            print(f"Checkbox with id '{cid}' error: {e}")
                except Exception as e:
                    print(f"Checkbox marking error: {e}")

                any_processed = True

                input(f"Row {i} submission {loop_idx+1} processed. Press Enter to continue...")

            except Exception as e:
                print(f"Row {i}, submission {loop_idx+1}: Error during processing - {e}")
                try:
                    driver.save_screenshot(f"debug_row_{i}_submission_{loop_idx+1}.png")
                    with open(f"debug_row_{i}_submission_{loop_idx+1}.html", "w", encoding="utf-8") as f:
                        f.write(driver.page_source)
                except Exception:
                    pass

            finally:
                close_optional_tab(switched, main_tab)
                # Always reset to main dashboard
                driver.get(DASHBOARD_URL)
                try:
                    wait_for_xpath("//div[contains(text(),'Yaga Gathering')]")
                    driver.execute_script('let el=document.querySelector("chrome-signin-app");if(el)el.remove();')
                except:
                    pass

            # Decrease remaining tickets for next submission
            remaining_wristbands -= wb_this_loop
            remaining_parking -= park_this_loop
            remaining_camper -= camper_this_loop

        # After all submissions for this row are done, update sheet as before!
        worksheet.update_cell(i, col_index_inv_sent+1, wristbands_cell)
        worksheet.update_cell(i, col_index_parking_sent+1, tickets_cell)
        if has_camper_pass:
            worksheet.update_cell(i, col_index_camper_sent+1, "1")

    print("Finished." if any_processed else "No rows were processed.")
    input("Press Enter to close browser...")

finally:
    driver.quit()