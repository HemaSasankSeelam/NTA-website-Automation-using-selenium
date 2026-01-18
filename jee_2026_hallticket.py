import selenium

from selenium.common import exceptions
import selenium.webdriver.remote.webelement
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.print_page_options import PrintOptions

import pandas as pd
import os, shutil
import itertools
from pathlib import Path
import traceback



class JEE_2026_Halltickets:

    def __init__(self, url, excel_path, downloads_folder):

        self.url = url
        chrome_service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=chrome_service)
        self.webdriver_wait = WebDriverWait(driver=self.driver, timeout=20) # wait the driver for max 20 seconds

        self.excel_path = Path(excel_path)
        self.downloads_folder = Path(downloads_folder)

        if not self.downloads_folder.exists():
            self.downloads_folder.mkdir(parents=True)
        
        self.results_excel_path = self.downloads_folder / "jee_2026_hall_tickets_data.xlsx"

        self.df = pd.read_excel(self.excel_path)
        self.filtered_df = self.df.loc[self.df["STATUS"] == "NO"]

        if len(self.filtered_df) == 0:
            print("  ALL HALL TICKETS ARE DOWNLOADED  ".center(50, "*"))
            return
        
        else:
            application_nos = self.filtered_df["APPLICATION_NO"].tolist()

            for application_no in application_nos:

                self.current_application_no = application_no
                password = self.filtered_df.loc[self.filtered_df["APPLICATION_NO"] == application_no, ["PASSWORD"]].values[0][0]

                result_from_page = self.login(password)

                if result_from_page == True:
                    self.df.loc[self.df["APPLICATION_NO"] == self.current_application_no,["STATUS"]] = "YES"
                    self.df.to_excel(self.excel_path, index=False)
                
                elif result_from_page == "server error":
                    print(" SERVER ERROR ".center(50, "-"))
                    exit()
                
            
            print(" ALL APPLICATION COMPLETED ".center(50, "*"))
            

    def login(self, password):

        self.driver.get(self.url)
        self.driver.find_element(By.ID, "ApplicationNo").send_keys(self.current_application_no) # sends application no
        self.driver.find_element(By.ID, "txtPassword").send_keys(password) # sends password

        self.driver.find_element(By.ID, "Captcha1").send_keys() # click on the captcha

        try:
            # helps to wait until the dropdown is clickable in the home page
            self.webdriver_wait.until((EC.visibility_of_element_located((By.ID, "ddlpostapplied"))))

        except exceptions.TimeoutException:
            try:
                self.driver.find_element(By.XPATH,"//h3[contains(text(),'HTTP Error 500.0 - Internal Server Error')]")
                return "server error"
            except Exception as e:
                pass
        
            try:
                last_ele = self.driver.find_elements(By.CLASS_NAME, "text-danger").pop()
                print(f"The Application no {self.current_application_no} is having {last_ele.text}")
                return "error"
            except Exception as e:
                return
        else:
            print("login successful")
            return self.home_page()



    def home_page(self):

        self.driver.implicitly_wait(5) # waits for 5 seconds manually

        try:
            self.driver.find_element(By.XPATH,"//h3[contains(text(),'HTTP Error 500.0 - Internal Server Error')]")
            return "server error"
        except Exception as e:
            pass

        
        select_ele = Select(self.driver.find_element(By.ID, "ddlpostapplied"))
        select_ele.select_by_index(1)
        self.driver.implicitly_wait(2) # waits for 2 seconds manually

        download_button_ele = self.driver.find_element(By.ID, "i-downloadbtn")

        download_button_ele.click()
        self.driver.implicitly_wait(2) # waits for 2 seconds manually

        th_list = []
        td_list = []

        th = self.driver.find_elements(By.XPATH, "//table[contains(@class,'tablefont')]//tr/th[not(.//img)]")
        th_list = [i.text for i in th if i.text.strip()]

        td = self.driver.find_elements(By.XPATH, "//table[contains(@class,'tablefont')]//tr/td[not(.//img) and count(../td | ../th) > 1]")
        for i in td:
            try:
                if i.text.strip():
                    td_list.append(i.text)
            except: 
                pass
        data_dict = dict(itertools.zip_longest(th_list, td_list, fillvalue="None"))

        self.webdriver_wait.until(EC.text_to_be_present_in_element_attribute((By.ID, "i-progress-inner"), "aria-valuenow", "1"))

        return self.write_excel(data_dict)

    def write_excel(self, data_dict):

        try:
            if not self.results_excel_path.exists():
                pd.DataFrame(data=[data_dict]).to_excel(self.results_excel_path, index=False)
            
            else:

                df1 = pd.read_excel(self.results_excel_path)
                new_df = pd.DataFrame(data=[data_dict])

                pd.concat([df1,new_df], axis="index", ignore_index=True).to_excel(self.results_excel_path, index=False)
            
    
        except Exception as e:
            print("Error in writing Excel file")
            print(traceback.format_exc())
            return False

        username = os.getlogin()
        c_drive_downloads = Path(r"C:\users\{}\Downloads".format(username))
        list_of_files = []

        for i in c_drive_downloads.glob(pattern=r"AdmitCard-{} (*.pdf".format(self.current_application_no)):
            list_of_files.append(i)

        for i in c_drive_downloads.glob(pattern=r"AdmitCard-{}.pdf".format(self.current_application_no)):
            list_of_files.append(i)

        list_of_files.sort()
        
        last_file:Path = list_of_files.pop()

        for file in list_of_files[:-1:]:
            os.remove(file)
        
        file_name = last_file.name
        final_pdf_path = self.downloads_folder / file_name
        shutil.move(last_file, final_pdf_path)

        return True
        



url = "https://examinationservices.nic.in/AdmitCardService/Admitcard/Login?enc=FG2WYxLrtc4EZvpaCNt7wpNzMY3W0bCWpSj8NNxqvms="
downloads_folder = r"F://jee_mais_2026_haltickets"

# [APPLICATION_NO, PASSWORD, STATUS] these are the columns
excel_path = r"C:\Users\seela\Downloads\2026_JEE MAIN CITY.xlsx"

JEE_2026_Halltickets(url, excel_path, downloads_folder)



