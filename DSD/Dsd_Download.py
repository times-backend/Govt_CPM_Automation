import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


def setup_driver(download_dir):
    os.makedirs(download_dir, exist_ok=True)
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--headless=new")  # Run in headless mode (use "new" for Chrome 109+)
    options.add_argument("--disable-gpu")
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=options)

def login(driver, username, password):
    driver.get("https://expresso.colombiaonline.com/expresso/home.htm")
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, "username"))).send_keys(username)
    driver.find_element(By.ID, "password").send_keys(password)
    driver.find_element(By.ID, "loginBtn").click()

def search_expresso_id(driver, expresso_id):
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.ID, "select2-headerRoId-container"))).click()
    search_expresso = WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.CLASS_NAME, "select2-search__field")))
    search_expresso.send_keys(expresso_id)
    time.sleep(3)
    dropdown_options = WebDriverWait(driver, 100).until(EC.presence_of_all_elements_located((By.CLASS_NAME, "select2-results__option")))
    if dropdown_options:
        dropdown_options[-1].click()  # Select the last option
    else:
        print("No options found in dropdown.")

def switch_to_new_tab(driver):
    WebDriverWait(driver, 5).until(lambda d: len(driver.window_handles) > 1)
    driver.switch_to.window(driver.window_handles[-1])

def find_and_download_file(driver, download_dir):
    time.sleep(2)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CLASS_NAME, "row.tab-details"))
    )
    rows = driver.find_elements(By.CSS_SELECTOR, "div.row.tab-details tr")
    before_download = set(os.listdir(download_dir))    
    for row in rows:
        tds = row.find_elements(By.TAG_NAME, "td")
        for td in tds:
            if "DSD" in td.text.upper():
                try:
                    download_link = WebDriverWait(row, 5).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, "a.text-success[title='Download']"))
                    )
                    download_link.click()
                    print("Download button clicked for:", td.text)
                    return before_download  # Return set of files before download
                except Exception as e:
                    print("Download button not found in row:", e)
        else:
            continue
        break

def wait_for_download(download_dir, before_download, timeout=30):
    start_time = time.time()
    downloaded_file = None  
    while time.time() - start_time < timeout:
        after_download = set(os.listdir(download_dir))
        new_files = after_download - before_download
        for file in new_files:
            if file.endswith(".xlsx"):
                downloaded_file = os.path.join(download_dir, file)
                return downloaded_file
        time.sleep(1)
    return None

def fetch_campaign_details(driver):
    time.sleep(2)
    order_name_selectors = [
        "#yoyoId > div.m-content.clearfix > div > div.m-portlet__body > div:nth-child(6) > div > div:nth-child(1) > div > div > div.caption > span.t-caption",
        "#yoyoId > div.m-content.clearfix > div > div.m-portlet__body > div:nth-child(5) > div > div:nth-child(1) > div > div > div.caption > span.t-caption"
    ]
    advertiser_name_selectors = [
        "#yoyoId > div.m-content.clearfix > div > div.m-portlet__body > div:nth-child(6) > div > div:nth-child(3) > div > table > tbody > tr > td:nth-child(2) > div:nth-child(1) > span.uppercase.font-grey-mint.m--font-boldest",
        "#yoyoId > div.m-content.clearfix > div > div.m-portlet__body > div:nth-child(5) > div > div:nth-child(3) > div > table > tbody > tr > td:nth-child(2) > div:nth-child(1) > span.uppercase.font-grey-mint.m--font-boldest"
    ]
    order_name = None
    for selector in order_name_selectors:
        try:
            order_name = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, selector))
            ).text
            break
        except TimeoutException:
            continue
    if not order_name:
        raise Exception("Order name not found ")
    advertiser_name = None
    for selector in advertiser_name_selectors:
        try:
            advertiser_name = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, selector))
            ).text
            advertiser_name = advertiser_name.split('(')[0].strip()
            break
        except TimeoutException:
            continue
    if not advertiser_name:
        raise Exception("Advertiser name not found using available selectors.")
    return order_name, advertiser_name

def Dsd_Download(expresso_id):
    download_dir = os.path.abspath("downloads")
    driver = setup_driver(download_dir)
    downloaded_file = None  # Initialize to None

    try:
        login(driver, "anurag.mishra1@timesinternet.in", "Times@9899")
        search_expresso_id(driver, expresso_id)
        switch_to_new_tab(driver)
        order_name,advertiser_name=fetch_campaign_details(driver)
        before_download = find_and_download_file(driver, download_dir)
        if before_download:
            downloaded_file = wait_for_download(download_dir, before_download)
            if downloaded_file:
                print("Downloaded Excel file:", downloaded_file)
            else:
                print("Download timed out or failed.")
        return order_name,advertiser_name,downloaded_file
    finally:
        driver.quit()

if __name__ == "__main__":
    expresso_id="265810"
    order_name,advertiser_name,dsd_file_location=Dsd_Download(expresso_id)
    print(f'Name of the order {order_name}')
    print(f'Name Of the Advertiser{advertiser_name}')
    print(f'DSD File location is {dsd_file_location}')