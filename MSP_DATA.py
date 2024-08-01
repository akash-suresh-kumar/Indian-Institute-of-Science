import os
import time
import csv
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import io

# Set up logging
logging.basicConfig(filename='msp_data_script.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# Set Tesseract path
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def setup_chrome_driver(download_dir):
    options = Options()
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
    options.add_experimental_option('prefs', prefs)
    return webdriver.Chrome(options=options)

def navigate_to_year(driver, year):
    try:
        year_select = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, 'msp_filter_year'))
        )
        select = Select(year_select)
        year_options = [option.text for option in select.options]
        
        if str(year) not in year_options:
            archive_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, '//a[@class="btn btn-info my-4" and contains(text(), "Archive")]'))
            )
            archive_button.click()
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, 'msp_filter_year'))
            )
            year_select = driver.find_element(By.ID, 'msp_filter_year')
            select = Select(year_select)
            year_options = [option.text for option in select.options]
            
            if str(year) not in year_options:
                logging.warning(f"Year {year} not found in both main and archive dropdowns. Skipping.")
                return False
        
        select.select_by_visible_text(str(year))
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//table[@class="table"]'))
        )
        time.sleep(5)
        return True
    except (TimeoutException, NoSuchElementException) as e:
        logging.error(f"Error navigating to year {year}: {e}")
        return False

def download_pdf_files(driver):
    pdf_links = driver.find_elements(By.XPATH, '//table[@class="table"]//a[contains(text(), "Download File (English)")]')
    if not pdf_links:
        logging.warning("No 'Download File (English)' links found.")
        return
    
    for link in pdf_links:
        try:
            link_url = link.get_attribute("href")
            if link_url:
                logging.info(f'Attempting to download: {link_url}')
                driver.execute_script("window.open('');")
                driver.switch_to.window(driver.window_handles[-1])
                driver.get(link_url)
                time.sleep(15)
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
                logging.info(f'Initiated download for: {link_url}')
            else:
                logging.warning(f'Link URL is empty or missing')
        except Exception as e:
            logging.error(f'Error downloading link: {e}')

def download_msp_data(start_year, end_year, download_dir):
    driver = setup_chrome_driver(download_dir)
    try:
        driver.get("https://desagri.gov.in/statistics-type/latest-minimum-support-price-msp-statement/")
        for year in range(start_year, end_year + 1):
            logging.info(f"Processing year: {year}")
            if navigate_to_year(driver, year):
                download_pdf_files(driver)
            driver.get("https://desagri.gov.in/statistics-type/latest-minimum-support-price-msp-statement/")
    except Exception as e:
        logging.error(f"An error occurred: {e}")
    finally:
        time.sleep(15)
        driver.quit()

def pdf_to_excel(pdf_path, excel_path):
    try:
        images = convert_from_path(pdf_path)
        text = ""
        for image in images:
            text += pytesseract.image_to_string(image)
        
        lines = text.split('\n')
        data = [line.split() for line in lines if line.strip()]
        
        df = pd.DataFrame(data)
        df.to_excel(excel_path, index=False, header=False, engine='openpyxl')
        logging.info(f"Converted {os.path.basename(pdf_path)} to Excel")
        return True
    except Exception as e:
        logging.error(f"Error converting {os.path.basename(pdf_path)} to Excel: {e}")
        return False

def excel_to_csv(excel_path, csv_path):
    try:
        df = pd.read_excel(excel_path, header=None, engine='openpyxl')
        df.to_csv(csv_path, index=False, header=False)
        logging.info(f"Converted {os.path.basename(excel_path)} to CSV")
        return True
    except Exception as e:
        logging.error(f"Error converting {os.path.basename(excel_path)} to CSV: {e}")
        return False

def convert_pdfs_to_excel_and_csv(directory):
    for filename in os.listdir(directory):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(directory, filename)
            excel_path = os.path.join(directory, filename[:-4] + '.xlsx')
            csv_path = os.path.join(directory, filename[:-4] + '.csv')
            if pdf_to_excel(pdf_path, excel_path):
                excel_to_csv(excel_path, csv_path)

def get_user_input(prompt, valid_responses):
    while True:
        response = input(prompt).lower().strip()
        if response in valid_responses:
            return response
        print(f"Please enter one of: {', '.join(valid_responses)}")

if __name__ == "__main__":
    download_dir = os.path.join(os.path.expanduser('~'), 'Downloads', 'msp_data')
    
    print("MSP Data Download and Conversion Script")
    print("---------------------------------------")

    download_pdfs = get_user_input("Do you want to download MSP data PDFs? (y/n): ", ['y', 'n'])
    if download_pdfs == 'y':
        start_year = int(input("Enter the start year for MSP data: "))
        end_year = int(input("Enter the end year for MSP data: "))
        print(f"Downloading MSP data for years {start_year} to {end_year}...")
        download_msp_data(start_year, end_year, download_dir)
        print("Download process completed.")

    convert_files = get_user_input("Do you want to convert the PDFs to Excel and CSV? (y/n): ", ['y', 'n'])
    if convert_files == 'y':
        print("Starting conversion process...")
        convert_pdfs_to_excel_and_csv(download_dir)
        print("Conversion process completed.")

    print("Script execution finished.")
    print("Check the log file 'msp_data_script.log' for detailed information.")