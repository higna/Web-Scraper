from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import os
import time

# Path to Chrome driver executable
chrome_driver_path = './chromedriver.exe'

# Setup Chrome options
chrome_options = Options()
# chrome_options.add_argument("--headless")  # Uncomment for headless mode

# Initialize WebDriver
service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service, options=chrome_options)

# Define WebDriver wait functions
def element_wait(by, value, timeout=120):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )
# Elemnet in view
def element_view(element):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    time.sleep(1)
# Page load & Reload
def page_load(timeout=120):
    try:
        WebDriverWait(driver, timeout).until(
            lambda driver: driver.execute_script('return document.readyState') == 'complete'
        )
        print("Page loaded successfully")
    except TimeoutException:
        print("Page load timed out")
def reload_page():
    try:
        driver.refresh()
        print("Page reloaded")
        page_load()
    except TimeoutException:
        print("Page reload timed out")
    except Exception as e:
        print(f"An error occurred while reloading the page: {e}")

# Login
def login():
    try:
        driver.get('https://url/#/login')
        print("Navigated to login page")
        
        username = element_wait(By.XPATH, "//input[@type='tel']")
        password = element_wait(By.XPATH, "//input[@type='password']")
        
        username.send_keys('yourusername')
        password.send_keys('yourpassword')
        password.send_keys(Keys.RETURN)
        
        print("Login successful")
        page_load()  # Wait for the page to load after login
    except TimeoutException:
        print("Login timed out")
    except NoSuchElementException as e:
        print(f"Element not found: {e}")
    except Exception as e:
        print(f"An error occurred during login: {e}")


# Page load
def page_load(timeout=40):
    try:
        WebDriverWait(driver, timeout).until(
            lambda driver: driver.execute_script('return document.readyState') == 'complete'
        )
        print("Page loaded successfully")
    except TimeoutException:
        print("Page load timed out")
# Reload page
def reload_page():
    try:
        driver.refresh()
        print("Page reloaded")
        page_load()
    except TimeoutException:
        print("Page reload timed out")
    except Exception as e:
        print(f"An error occurred while reloading the page: {e}")

# Open webform
def open_webform():
    try:
        driver.get('https://enketo.seedtracker.org/x/ql9Kcse2')
        print("Navigated to webform")
        
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//input[@name='/data/producercode']"))
        )
        print("Webform loaded successfully")
        
        reload_page()
        page_load()
    except TimeoutException:
        print("Webform load timed out")
    except NoSuchElementException as e:
        print(f"Element not found: {e}")
    except Exception as e:
        print(f"An error occurred while opening the webform: {e}")

# Pull data
def pulldata(field_id):
    try:
        producercode = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="/data/producercode"]'))
        )
        print(f"Found field-id: {field_id}")
        producercode.send_keys(field_id)
        producercode.send_keys(Keys.TAB)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//label//span[@class='question-label active']//strong[text()='Seed Supplier:']"))
        )
        print("Data Pulled Successfully")
        prodcode = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="/data/prodcode"]'))
        )
        prodcode.send_keys(field_id)
    except NoSuchElementException as e:
        print(f"Element not found while pulling data: {e}")
    except TimeoutException as e:
        print(f"Timeout while waiting for element: {e}")
    except Exception as e:
        print(f"An error occurred while pulling data: {e}")

# Fill form
def fillform(inspector, ass_inspector, date_inspect, date_planting, status, accreditation_status, site_suitability, 
              seed_source_verification, seed_source_suitability, seed_quality_verification, decision, comment):
    
    def select_radio_button(name, value):
        try:
            css_selector = f'input[name="{name}"][value="{value}"]'
            print(f"Attempting to select radio button for: {value}")
            WebDriverWait(driver, 120).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, css_selector))
            )
            radio_button = driver.find_element(By.CSS_SELECTOR, css_selector)
            if not radio_button.is_selected():
                radio_button.click()
                print(f"Selected: {value}")
            else:
                print(f"Radio button for {value} is already selected")
        except TimeoutException:
            print(f"Timeout while locating radio button for: {value}")
        except Exception as e:
            print(f"An error occurred while selecting radio button for {value}: {e}")

    try:
        # Inspector name
        inspector_name = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="/data/inspect/nameinsp"]'))
        )
        inspector_name.send_keys(inspector)
        print("Inspector name entered")
        time.sleep(2)

        # Assistant Inspector name
        other_inspectname = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="/data/inspect/other_inspectname"]'))
        )
        other_inspectname.send_keys(ass_inspector)
        print("Assistant Inspector name entered")
        time.sleep(2)

        # Date of Inspection
        date_input1 = WebDriverWait(driver, 90).until(
            EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Date for 1st inspection (pre-planting)")]/following::input[@placeholder="yyyy-mm-dd"]'))
        )
        date_input1.click()
        time.sleep(2)
        date_input1.clear()
        time.sleep(2)
        date_input1.send_keys(date_inspect)
        time.sleep(2)
        date_input1.send_keys(Keys.RETURN)
        time.sleep(2)
        print(f"Date of Inspection field populated with: {date_inspect}")
        time.sleep(2)

        # Actual Date of Planting
        date_input2 = WebDriverWait(driver, 90).until(
            EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Actual date of planting")]/following::input[@placeholder="yyyy-mm-dd"]'))
        )
        date_input2.click()
        time.sleep(2)
        date_input2.clear()
        time.sleep(2)
        date_input2.send_keys(date_planting)
        time.sleep(2)
        date_input2.send_keys(Keys.RETURN)
        time.sleep(2)
        print(f"Actual Date of Planting field populated with: {date_planting}")
        time.sleep(2)

        # Select radio buttons
        select_radio_button("/data/grp_sitesuit/accreditation", accreditation_status)
        select_radio_button("/data/grp_sitesuit/Site_Verification", status)
        select_radio_button("/data/grp_sitesuit/Site_suitability", site_suitability)
        select_radio_button("/data/grp_sitesuit/Seed_source_verification", seed_source_verification)
        select_radio_button("/data/grp_sitesuit/Seed_source_suitability", seed_source_suitability)
        select_radio_button("/data/grp_sitesuit/Seed_quality_verification", seed_quality_verification)
        select_radio_button("/data/grp_sitesuit/Decision", decision)

        # Comments by Inspector
        comments_field = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="/data/grp_sitesuit/Comments_by_Inspector"]'))
        )
        comments_field.send_keys(comment)
        print("Comments entered")

    except NoSuchElementException as e:
        print(f"Element not found: {e}")
    except TimeoutException as e:
        print(f"Timeout while waiting for an element: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Upload image
def upload_photo(photo_path):
    try:
        # Convert relative path to absolute path
        absolute_path = os.path.abspath(photo_path)
        
        # Locate the actual file input element
        file_input = WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[name="/data/grp_sitesuit/photoid1"]'))
        )
        file_input.send_keys(absolute_path)
        print(f"Photo uploaded from: {absolute_path}")
        
        # Debugging: Check if the file preview or feedback is updated
        time.sleep(5)  # Adjust sleep if needed to ensure the upload is completed
        feedback = driver.find_element(By.CSS_SELECTOR, 'div.file-feedback').text
        print(f"Upload feedback: {feedback}")

    except NoSuchElementException as e:
        print(f"Element not found for photo upload: {e}")
    except TimeoutException as e:
        print(f"Timeout while locating the photo upload field: {e}")
    except Exception as e:
        print(f"An error occurred while uploading photo: {e}")

# Submit form
def submit_form(field_id):
    try:
        # Wait for the submit button to be clickable
        submit_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.ID, 'submit-form'))
        )
        submit_button.click()
        print(f"Form with field-id {field_id} submitted successfully")

    except NoSuchElementException as e:
        print(f"Submit button not found: {e}")
    except TimeoutException as e:
        print(f"Timeout while waiting for submit button: {e}")
    except Exception as e:
        print(f"An error occurred while submitting the form: {e}")

# Main execution
def main():
    login()
    time.sleep(2)
    open_webform()
    time.sleep(2)
    # Load data from XLSX file
    xlsx_file_path = os.path.join(os.getcwd(), 'PTST.xlsx')
    df = pd.read_excel(xlsx_file_path)

    for index, row in df.iterrows():
        print(f"Processing row {index + 1}")

        # Pull data
        pulldata(row['field_id'])
        
        # Fill form
        fillform(row['inspector'], row['ass_inspector'], row['date_inspect'], row['date_planting'],
                 row['status'], row['accreditation_status'], row['site_suitability'],
                 row['seed_source_verification'], row['seed_source_suitability'],
                 row['seed_quality_verification'], row['decision'], row['comment'])
        
        # Upload photo (use a fixed path or ensure photo exists at the specified path)
        upload_photo('./ptst.JPG')  # Adjust path if necessary
        time.sleep(2)
        # Submit form
        #submit_form(row['field_id'])

        # Sleep between submissions to avoid overwhelming the server
        time.sleep(5)  # Adjust as necessary

    driver.quit()

if __name__ == "__main__":
    main()