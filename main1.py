"""
Main DDTF Execution Engine for Automation Testing
"""

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from data import Data
from locators import WebLocators
from excel_functions import AkshyaExcelReader

# Setup WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.maximize_window()
driver.implicitly_wait(10)
driver.get(Data.URL)

# Read Excel Data
excel_reader = AkshyaExcelReader(Data.EXCEL_FILE, Data.SHEET_NUMBER)
rows = excel_reader.row_count()

for row in range(2, rows + 1):
    try:
        username = excel_reader.read_data(row, column_number=6).strip()
        password = excel_reader.read_data(row, column_number=7).strip()

        # Locate elements and perform login
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, WebLocators.USERNAME_INPUT_BOX))
        ).send_keys(username)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, WebLocators.PASSWORD_INPUT_BOX))
        ).send_keys(password)

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, WebLocators.SUBMIT_BUTTON))
        ).click()

        # Wait for the page to load and check if the login is successful
        WebDriverWait(driver, 10).until(
            EC.url_to_be(Data.DASHBOARD_URL)
        )

        if driver.current_url == Data.DASHBOARD_URL:
            print(f"SUCCESS: Login successful with USERNAME={username}")
            excel_reader.write_data(row, column_number=8, data="TEST PASSED")

            # Logout
            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, WebLocators.HAMBURGER_BUTTON))
            ).click()

            WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, WebLocators.LOGOUT_BUTTON))
            ).click()

        else:
            print(f"ERROR: Login Failed with USERNAME={username}")
            excel_reader.write_data(row, column_number=8, data="TEST FAIL")
            driver.refresh()

    except Exception as e:
        print(f"ERROR: Exception occurred for USERNAME={username}: {e}")
        excel_reader.write_data(row, column_number=8, data="ERROR")
        driver.refresh()

# Close the WebDriver
driver.quit()