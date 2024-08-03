from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import os
import time

# Initialize the Chrome driver
service = ChromeService(executable_path=ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# URL to scrape
url = "https://www.mordorintelligence.com/market-analysis/aerospace-defense"

# Open the URL
driver.get(url)

# Click the "See More Aerospace & Defense Reports" button twice
for _ in range(2):
    try:
        see_more_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CLASS_NAME, 'view-more-report-btn'))
        )
        see_more_button.click()
        time.sleep(5)  # Wait for new content to load
    except Exception as e:
        print(f"Error clicking see more button: {e}")

# Collect report data
reports = []
report_elements = driver.find_elements(By.CSS_SELECTOR, 'div.single-report-card-wrap')
for report in report_elements:
    title_element = report.find_element(By.CSS_SELECTOR, 'a.report-list-single-card-title')
    title = title_element.text
    link = title_element.get_attribute('href')
    reports.append({
        'Title': title.strip(),
        'URL': link
    })

# Close the browser
driver.quit()

# Convert data to DataFrame
df = pd.DataFrame(reports)

# File name
file_name = 'mordorintelligence_reports.xlsx'

# Check if the Excel file already exists
if os.path.exists(file_name):
    # Read existing file
    existing_df = pd.read_excel(file_name, engine='openpyxl')
    # Append new data
    df = pd.concat([existing_df, df], ignore_index=True)

# Save data to Excel
df.to_excel(file_name, index=False, engine='openpyxl')
