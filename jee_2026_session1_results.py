import selenium

from selenium.common import exceptions
import selenium.webdriver.remote.webelement
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import pandas as pd
from pathlib import Path

import traceback

import sarvam_login

class JeeSession1:

    def __init__(self,url,excel_path) -> None:

        self.url = url
        self.excel_path = excel_path
        self.parent_path = Path(excel_path).parent
        self.file_name = Path(excel_path).stem
        
        chrome_service = Service(executable_path=ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=chrome_service)

        self.actions_chains = ActionChains(self.driver)
        self.webdriver_wait = WebDriverWait(driver=self.driver, timeout=15)

    def read_from_excel(self) -> None:

        self.dataframe = pd.read_excel(self.excel_path)

        self.results_saving_path = self.parent_path / Path(str(self.file_name) + "_results_by_selenium.xlsx")

        if Path(self.results_saving_path).exists():
            print("The File Name is already Available...")
       
        # print(self.dataframe.columns) # prints the available columns
        self.dataframe = self.dataframe.drop_duplicates(subset=['APPLICATION NO'])
        self.dataframe.to_excel(self.excel_path, index=False)

        self.dataframe = pd.read_excel(self.excel_path)
        self.filtered_dataframe = self.dataframe.loc[self.dataframe["STATUS"] == "NO"]

        if len(self.filtered_dataframe) == 0:
            print("ALL MEMBERS STATUS IS YES")
            print("CHECK THE STATUS IN EXCEL SHEET")
            return
        
        applications_nos = self.filtered_dataframe['APPLICATION NO']
        
        for application_no in applications_nos.tolist():
            
            password = self.filtered_dataframe.loc[self.filtered_dataframe["APPLICATION NO"] == application_no]["PASSWORD"].values[0]
            
            is_verified = self.main_page(application_no=application_no, password=password)

            if is_verified:
                self.dataframe.loc[self.dataframe["APPLICATION NO"] == application_no,["STATUS"]] = "YES"
                self.dataframe.to_excel(self.excel_path, index=False)
        
        print("COMPLETED".center(20,"*"))
        return
    
    def main_page(self, application_no, password) -> bool:
        
        self.driver.get(url = self.url)

        self.driver.find_element(by=By.ID, value="txtAppNo").send_keys(application_no)

        self.driver.find_element(by=By.ID, value="txtPassword").send_keys(password)

        captcha_img_path = (self.parent_path/"captcha_image.png").as_posix()
        img_field = self.driver.find_element(by=By.XPATH, value="//a[@id='captcha-refresh-btn']//preceding-sibling::img")
        img_field.screenshot(captcha_img_path)

        captcha_value = sarvam_login.upload_image_and_get_captcha(path=captcha_img_path)
        self.driver.find_element(by=By.ID, value="Captcha1").send_keys(captcha_value)

        self.actions_chains.send_keys(Keys.RETURN).perform() # clicks on the enter button

        try:
            self.webdriver_wait.until((EC.visibility_of_element_located((By.XPATH,"(//strong[normalize-space()='Physics'])[1]"))))

        except exceptions.TimeoutException:
            print("Login Failed")
            return False
        
        else:
            self.inside_page()
            return True
    
    def inside_page(self) -> bool:
        try:
            tables = self.driver.find_elements(by=By.XPATH,value="//table[@class='table-bordered']")

            person_details_table = tables[0]
            person_scores_table = tables[1]

            person_details = {}
            person_scores = {}

            # person details table
            previous_tag = ""
            index = 0
            for element in person_details_table.find_elements(by=By.TAG_NAME, value="td"):
                
                element:selenium.webdriver.remote.webelement.WebElement = element

                text_value = element.text
                if not text_value:
                    continue

                if index%2 == 0:
                    person_details[text_value] = ""
                    previous_tag = text_value
                else:
                    person_details[previous_tag] = text_value
                
                index += 1
            
            # person scores table
            previous_tag = ""
            index = 0
            for element in person_scores_table.find_elements(by=By.TAG_NAME, value="td"):

                element:selenium.webdriver.remote.webelement.WebElement = element

                text_value = element.text
                if not text_value or element.get_attribute("rowspan") != None:
                    continue

                if index%2 == 0:
                    person_scores[text_value] = ""
                    previous_tag = text_value
                else:
                    person_scores[previous_tag] = text_value
                
                index += 1

            person_final_dict = person_details | person_scores
            return self.write_to_excel_file(data_dict = person_final_dict)
        
        except Exception as e:
            print(f"Exception in inside page {e}")
            print(traceback.format_exc())
            exit()
        
    def write_to_excel_file(self,data_dict:dict) -> bool:
        try:
            if not Path(self.results_saving_path).exists():
                
                new_dataframe = pd.DataFrame(data=[data_dict])
                new_dataframe.to_excel(self.results_saving_path, index=False)
                return True

            new_dataframe = pd.read_excel(self.results_saving_path)
            new_row = pd.DataFrame(data=[data_dict])

            new_dataframe = pd.concat([new_dataframe, new_row], ignore_index=True)
            new_dataframe.to_excel(self.results_saving_path, index=False)
            return True
        
        except Exception as e:
            print(f"Exception in write_to_excel_file function {e}")
            print(traceback.format_exc())
            exit()
        

url = "https://examinationservices.nic.in/ResultoService26/JE26S1P1/Login"
excel_path = r"C:\Users\seela\Downloads\APPLICATION NUMBER.xlsx"

sarvam_login.main()

jee_object = JeeSession1(url=url, excel_path=excel_path)
jee_object.read_from_excel()
