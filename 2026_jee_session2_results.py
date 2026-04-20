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

import pandas as pd
from pathlib import Path

import traceback 
class JeeSession2:

    def __init__(self,url,excel_path) -> None:

        self.url = url
        self.excel_path = excel_path
        self.parent_path = Path(excel_path).parent
        self.file_name = Path(excel_path).stem
        
        chrome_service = Service(executable_path=ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=chrome_service)

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
        
        print("COMPLETED".center(5,"*"))
        return
    
    def main_page(self, application_no, password) -> bool:
        
        self.driver.get(url = self.url)

        self.driver.find_element(by=By.ID, value="txtAppNo").send_keys(application_no)

        self.driver.find_element(by=By.ID, value='txtPassword').send_keys(password)
        self.driver.find_element(by=By.ID, value='Captcha1').send_keys()

        try:
            self.webdriver_wait.until((EC.visibility_of_element_located((By.XPATH,"(//strong[normalize-space()='Physics'])[1]"))))

        except exceptions.TimeoutException:
            print("Login Failed")
            return False
        
        else:
            self.inside_page()
            return True
        
    def inside_page(self) -> None:

        try:
            person_details_table = self.driver.find_elements(by=By.XPATH,value="//table[@class='tablecustom']")[0]
            person_details = {}

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
            
            #person scores table
            person_scores_table = self.driver.find_elements(by=By.XPATH, value="//table[@class='tablecustom']")[1]
            person_scores = {}
            list_of_elements = []

            for element in person_scores_table.find_elements(by=By.TAG_NAME, value="td"):

                element:selenium.webdriver.remote.webelement.WebElement = element

                if not element.get_attribute(name='rowspan') and not element.get_attribute(name='colspan'):
                    list_of_elements.append(element.text)

            from itertools import batched
            BATCH_SIZE = 4
            
            data = {}
            for each_tuple in list(batched(list_of_elements, n=BATCH_SIZE)):
                
                if len(each_tuple) < BATCH_SIZE:
                    continue
                
                key,*value = each_tuple
                data[key] = value

            scores_dataframe = pd.DataFrame(data)
            scores_dataframe.set_index(keys=list(data.keys())[0], inplace=True)
            
            scores_dataframe = scores_dataframe.stack().to_frame().T
            scores_dataframe.columns = [f"{col}_{idx}" for idx, col in scores_dataframe.columns]
            person_scores = scores_dataframe.to_dict(orient='index')[0]

            # person rank
            
            d1 = {"CRL":"","GEN-EWS":"","OBC-NCL":"","SC":"","ST":"","PWD_CRL":"","PWD_GEN-EWS":"","PWD_OBC-NCL":"","PWD_SC":"","PWD_ST":""}
        
            last = self.driver.find_element(by=By.XPATH, value="//body[1]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[5]/td[1]/table[1]/tbody[1]/tr[4]")
            index = 0
            keys = list(d1.keys())
            for i in last.find_elements(by=By.TAG_NAME, value="td"):
                d1[keys[index]] = i.text
                index +=1
        
            is_eligible = self.driver.find_element(by=By.XPATH, value="//body[1]/div[1]/div[2]/div[1]/div[1]/table[1]/tbody[1]/tr[8]/td[1]/table[1]/tbody[1]/tr[1]/td[1]").text

            person_final_dict = person_details | person_scores | {"IS ELIGIBLE":is_eligible.split(":")[1].strip()} | d1

            self.write_to_excel_file(data_dict = person_final_dict)
        
        except Exception as e:
            print(f"Exception in Inside Page function {e}" )
            print(traceback.format_exc())
            exit()

    def write_to_excel_file(self,data_dict:dict) -> None:
        
        try:
            if not Path(self.results_saving_path).exists():
                
                new_dataframe = pd.DataFrame(data=[data_dict])
                new_dataframe.to_excel(self.results_saving_path, index=False)
                return True

            new_dataframe = pd.read_excel(self.results_saving_path)
            new_row = pd.DataFrame(data=[data_dict])

            new_dataframe = pd.concat([new_dataframe, new_row], axis='index', ignore_index=True)
            new_dataframe.to_excel(self.results_saving_path, index=False)
            return True
        
        except Exception as e:
            print(f"Exception in write_in_excel_file function {e}")
            print(traceback.format_exc())
            exit()

if  __name__ == "__main__":
    url = "https://examinationservices.nic.in/ResultoService26/P1S2JM26/Login"
    excel_path = r"C:\Users\seela\Downloads\APPLICATION NUMBER.xlsx"

    jee_object = JeeSession2(url=url, excel_path=excel_path)
    jee_object.read_from_excel()