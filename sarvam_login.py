import selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.remote.webelement import WebElement

import time
from pathlib import Path
import os
from dotenv import load_dotenv

current_dir = Path(__file__).parent
os.chdir(current_dir)
load_dotenv("env_file.env")

chrome_options = Options()
chrome_options.add_experimental_option("detach", True) # This will keep the browser open after the script finishes

chrome_service = Service(executable_path=ChromeDriverManager().install())

driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
web_driver_wait = WebDriverWait(driver, timeout=30)
driver.set_window_position(x=450, y=10)

main_page_url = "https://dashboard.sarvam.ai/vision"
error_counter = 0

def main():

    driver.get("https://dashboard.sarvam.ai/login")
    web_driver_wait.until(EC.url_contains('https://dashboard.sarvam.ai/login'))

    web_driver_wait.until(EC.presence_of_element_located((By.NAME, "identifier")))
    driver.find_element(By.NAME, "identifier").send_keys(os.getenv("google_account_email"))

    driver.find_element(By.XPATH, "//input[@name='identifier']//following::button").click()

    web_driver_wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='password']"))) # wait until the password input field is present
    driver.find_element(By.XPATH, "//input[@name='password']").send_keys(os.getenv("google_account_password"))

    login_button = driver.find_element(By.XPATH, "//button[@type='submit']")
    login_button.click()

    time.sleep(3) # wait for the login process to complete and the vision page to load
    
    driver.get(main_page_url) # navigate to the main page
    web_driver_wait.until(EC.url_to_be(main_page_url)) # wait until the URL changes to the vision spage
    

def upload_image_and_get_captcha(path):
    global error_counter

    try:
        web_driver_wait.until(EC.url_contains(main_page_url))

        file_upload:WebElement = web_driver_wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='file']"))) # wait until the URL changes to the vision page
        file_upload.send_keys(path)

        web_driver_wait.until(EC.presence_of_element_located((By.XPATH, "//button[@type='button']"))) # wait until the "Submit" button is present
        ele = driver.switch_to.active_element
        ele.find_elements(By.XPATH, "//button[@type='button']")[-1].click()

        driver.switch_to.default_content()

        captcha_answer_ele:WebElement = web_driver_wait.until(EC.presence_of_element_located((By.XPATH, "//p[@class='text-sm leading-relaxed text-tatva-text-primary break-words']"))) # wait until the captcha value is displayed

        captcha_value = captcha_answer_ele.text

        driver.get(main_page_url) # navigate back to the main page
        driver.refresh()
        
        return captcha_value
    
    except Exception:
        error_counter += 1
        
        if error_counter == 2:
            error_counter = 0
            driver.refresh()

        print("An error occurred while uploading the image and getting the captcha value.")
        return "login-error"

