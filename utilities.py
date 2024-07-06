import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
import xlsxwriter
import logging

def scroll_click(driver, element, attempts=3):
    for attempt in range(attempts):
        try:
            driver.execute_script("arguments[0].scrollIntoView();", element)
            driver.execute_script("arguments[0].click();", element)
            return True
        except Exception as e:
            logging.warning(f"Scroll click attempt {attempt + 1} failed: {e}")
            time.sleep(2)
    return False

def open_in_new_window(driver, url):
    """Opens a URL in a new window."""
    driver.execute_script(f"window.open('{url}', '_blank');")

def initialize_driver(wait=5, headless=True):
    """Initializes the Selenium WebDriver with specified options."""
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920x1080")
    driver = webdriver.Chrome(options=chrome_options)
    driver.implicitly_wait(wait)
    return driver

def replace_non_char_with_hyphen(text):
    """Replaces non-alphanumeric characters in a string with hyphens."""
    return ''.join([c if c.isalnum() else '-' for c in text])

def dump_with_xlsxwriter(data, file_path):
    """Dumps the data into an Excel file using xlsxwriter."""
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet()

    if not data:
        return

    headers = list(data[0].keys())
    for col, header in enumerate(headers):
        worksheet.write(0, col, header)

    for row, item in enumerate(data, start=1):
        for col, header in enumerate(headers):
            worksheet.write(row, col, item.get(header, ""))

    workbook.close()

