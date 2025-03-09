import time
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import os

def setup_driver():
    # Add a bypass in case the webdriver path is incorrect in drivers.json file
    chromedriver_path = ChromeDriverManager().install()
    if not chromedriver_path.endswith("chromedriver.exe"):
        chromedriver_dir = os.path.dirname(chromedriver_path)
        chromedriver_path = os.path.join(chromedriver_dir, "chromedriver.exe")
    
    service = Service(chromedriver_path)
    options = webdriver.ChromeOptions()
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    # The below feature stops chrome from flagging the browser as being controlled by and automated software, thereby allowing to bypass all bot detection
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument("--headless=new")
    prefs = {
        # Adjust path as needed
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True,
        "profile.managed_default_content_settings.images": 2,
    }
    options.add_experimental_option('prefs', prefs)

    # Begin the driver
    driver = webdriver.Chrome(service=service, options=options)
    actions = ActionChains(driver)
    driver.maximize_window()
    return driver, actions

def scroll(driver, i, j):
    try:
        compbut = driver.find_element(By.XPATH,f'//tbody//tr[{i+1}]//td[@headers="companyN"]//a')
        complocate = driver.find_element(By.XPATH,'//table[@id="shareHolderPattern"]//thead')
        driver.execute_script("arguments[0].scrollIntoView(true);", complocate)
        time.sleep(0.5)
        compbut.click()

    except WebDriverException:
        view = driver.find_element(By.XPATH,f'//tbody//tr[{j-4}]//td[@headers="companyN"]//a')
        driver.execute_script("arguments[0].scrollIntoView(true);", view)
        time.sleep(0.5)
        scroll(driver, i, j+5)