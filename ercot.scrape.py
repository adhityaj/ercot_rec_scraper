from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import time

# Specify the correct path to your Chrome binary
chrome_binary_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"  # Adjust this path if necessary

# Initialize Chrome options
options = Options()
options.add_argument("--headless")  # Run Chrome in headless mode (without GUI)
options.binary_location = chrome_binary_path  # Set the Chrome binary location

# Use WebDriverManager to manage ChromeDriver automatically
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)


def scrape_table():
    # Wait for the table to be present and visible in the DOM
    try:
        table = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "table.table"))
        )
    except:
        print("Timeout waiting for table")
        print("Page source:", driver.page_source)
        return None, None

    # Give a little more time for the table to be fully populated
    time.sleep(2)

    # Get the table HTML after it's loaded
    table_html = table.get_attribute('outerHTML')
    soup = BeautifulSoup(table_html, 'html.parser')

    # Extract headers
    headers = [header.text.strip() for header in soup.select('thead th')]

    # Extract rows
    rows = []
    for row in soup.select('tbody tr'):
        row_data = [cell.text.strip() for cell in row.select('td')]
        rows.append(row_data)

    return headers, rows


def get_max_page_number():
    try:
        pagination = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "ul.pagination"))
        )
        page_items = pagination.find_elements(By.CSS_SELECTOR, "li.page-item")
        last_page_item = page_items[-2]  # The second to last item should be the last page number
        return int(last_page_item.text)
    except:
        print("Error finding max page number, defaulting to 1")
        return 1


def scrape_category(url):
    driver.get(url)
    all_rows = []
    headers = None
    max_page = get_max_page_number()

    for page in range(1, max_page + 1):
        print(f"Scraping page {page} of {max_page}")
        page_headers, page_rows = scrape_table()

        if page_headers is None or page_rows is None:
            print(f"No data found on page {page}")
            break

        if headers is None:
            headers = page_headers
        all_rows.extend(page_rows)

        if page < max_page:
            try:
                # Click the next page number
                next_page_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH,
                                                f"//ul[contains(@class, 'pagination')]//li[contains(@class, 'page-item')]/a[text()='{page + 1}']"))
                )
                next_page_button.click()
                time.sleep(2)  # Wait for the page to load
            except:
                print(f"Error navigating to page {page + 1}")
                break

    return headers, all_rows


# Define the base URLs for each category
categories = {
    "REC Generator": "https://sa.ercot.com/rec/account-type"
}

# Create an Excel writer
with pd.ExcelWriter('ercot_data_output.xlsx', engine='openpyxl') as writer:
    for category, url in categories.items():
        print(f"Scraping category: {category}")
        headers, data = scrape_category(url)
        if headers and data:
            df = pd.DataFrame(data, columns=headers)
            df.to_excel(writer, sheet_name=category, index=False)
            print(f"Data for {category} written to Excel. Total rows: {len(data)}")
        else:
            print(f"No data found for category: {category}")

    # Ensure at least one sheet is written
    if writer.sheets:
        print("Excel file created successfully")
    else:
        print("No data was written to the Excel file")
        # Create a dummy sheet to avoid the "At least one sheet must be visible" error
        pd.DataFrame().to_excel(writer, sheet_name="Empty", index=False)

# Quit the browser when done
driver.quit()

print("Scraping completed")