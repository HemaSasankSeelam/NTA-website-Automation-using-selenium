import selenium
from selenium.common import exceptions
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.remote.webelement import WebElement


import time
from pathlib import Path
import pandas as pd


class JEE_2026_APPLICATION:

    def __init__(self, excel_path, download_dir):
        
        self.url = "https://examinationservices.nic.in/JeeMainx2026/Root/CandidateLogin.aspx?enc=Ei4cajBkK1gZSfgr53ImFVj34FesvYg1WX45sPjGXBqfcvMYv/FHq/Da9QEnq781"

        self.excel_path = excel_path
        
        self.chrome_service = Service(executable_path= ChromeDriverManager().install())
        while True:

            self.df = pd.read_excel(self.excel_path)
            self.df["PAYMENT_STATUS"] = self.df["PAYMENT_STATUS"].astype("str")
            self.filtered_df = self.df.loc[self.df["STATUS"] == "NO"]

            self.driver = webdriver.Chrome(service=self.chrome_service)
            self.driver.get(self.url)
            self.driver.maximize_window()
            
            self.webdriver_wait = WebDriverWait(driver=self.driver, timeout=15)

            if len(self.filtered_df) == 0:
                break
            
            application_no = self.filtered_df["APPLICATION_NO"].to_list()[0]
            self.login(application_no)
            time.sleep(0.5)

        print("ALL completed")

    def login(self,application_no):

        self.application_no = str(application_no)
        password = str(self.filtered_df.loc[self.filtered_df["APPLICATION_NO"] == application_no,["PASSWORD"]].values[0][0])

        print(password)

        self.driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtRegno").send_keys(self.application_no)

        self.driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtPassword").send_keys(password)

        self.driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtsecpin").send_keys("")

        try:
            WebDriverWait(self.driver, 10).until((EC.url_changes(self.driver.current_url)))

            try:
                # Wait briefly to see if URL contains the duplicate login parameter
                WebDriverWait(self.driver, 5).until(
                    (EC.url_matches("appFormId="))
                )
                has_duplicate_login = True
            except:
                has_duplicate_login = False

            if has_duplicate_login:
                try:
                    ele = self.driver.find_element(By.CLASS_NAME, "btn-danger")
                    if ele:
                        ele.click()
                        self.driver.quit()
                except:
                    print("invalid credentials")
                    self.df.loc[self.df["APPLICATION_NO"] == int(self.application_no),["STATUS"]] = "YES"
                    self.df.to_excel(self.excel_path, index=False)
                    return
                
            else:
                self.home_page()

            return

        except exceptions.TimeoutException:
            print("login failed")
            self.driver.quit()
            return
    

    def home_page(self):
        is_payment_competed = self.driver.find_element(By.ID, "ctl00_LoginContent_rptApplicationStatus_ctl02_lblstatus").text.strip().lower()


        if is_payment_competed != "completed":
            self.df.loc[self.df["APPLICATION_NO"] == int(self.application_no),["PAYMENT_STATUS"]] = "NOT PAID"
        
        else:
            self.driver.find_element(By.ID, "ctl00_LoginContent_linkDownConfirm").click()
            self.driver.implicitly_wait(2)

            self.driver.find_element(By.ID, "downloadpdfbtn").click()

           
            WebDriverWait(self.driver,100).until(EC.alert_is_present())
            
            self.df.loc[self.df["APPLICATION_NO"] == int(self.application_no),["PAYMENT_STATUS"]] = "PAID"
            self.driver.switch_to.alert.accept()
        

        self.df.loc[self.df["APPLICATION_NO"] == int(self.application_no),["STATUS"]] = "YES"
        self.df.to_excel(self.excel_path, index=False)
        self.driver.quit()
        return


if __name__ == "__main__":

    excel_path = r"F:\2026_jee_application.xlsx"
    download_dir = r"F:\\2026_applicaitons_jee"

    JEE_2026_APPLICATION(excel_path=excel_path,
                        download_dir=download_dir)
