import selenium

from selenium.common import exceptions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

import pandas as pd
from pathlib import Path
import traceback 
import time

class JEE_RESPONSE_SHEETS:

    def __init__(self, url, excel_path):
        self.url = url # website url
        self.excel_path = excel_path # credentials excel path
        self.parent_path = Path(excel_path).parent # parent of the file
        self.file_name = Path(excel_path).stem # file name without extension
        self.results_saving_path = self.parent_path / Path(str(self.file_name) + "_results_by_selenium.xlsx") # excel path for the response sheet links

        chrome_service = Service(executable_path=ChromeDriverManager().install()) # installing the chrome driver
        self.driver = webdriver.Chrome(service=chrome_service) # driver 
        self.driver.set_window_position(x=200, y=50) # sets the window position
        self.webdriver_wait = WebDriverWait(driver=self.driver, timeout=10) # webdriver wait

    def read_from_excel(self) -> None:
        self.dataframe = pd.read_excel(self.excel_path) # reads the credential excel
        self.dataframe = self.dataframe.drop_duplicates(subset=['APPLICATION_NO']) # removes the duplicates from the application column
        self.dataframe.to_excel(self.excel_path, index=False) # and save the excel path
        self.dataframe = pd.read_excel(self.excel_path)

        self.filtered_dataframe = self.dataframe.loc[self.dataframe["STATUS"] == "NO"] # filters the dataframe whose status is NO
        if len(self.filtered_dataframe) == 0:
            print("ALL MEMBERS STATUS IS YES".center(50, "*"))
            return
        
        applications_nos = self.filtered_dataframe['APPLICATION_NO']

        for application_no in applications_nos.tolist():

            try:
                password = self.filtered_dataframe.loc[self.filtered_dataframe["APPLICATION_NO"] == application_no]["PASSWORD"].values[0]

                is_verified = self.main_page(application_no=application_no, password=password)

                if is_verified:
                    # if the application is verified then update the status to YES
                    self.dataframe.loc[self.dataframe["APPLICATION_NO"] == application_no, ["STATUS"]] = "YES"
                    self.dataframe.to_excel(self.excel_path, index=False) # update the excel file 

            except Exception as e:
                # if any error handling the error
                print(f"Error for application {application_no}: {e}")

        print("COMPLETED".center(50, "*"))

    def main_page(self, application_no, password) -> bool:

        self.driver.get(self.url) # gets the website url

        try:
            application_no_field = self.driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtRegno")
            application_no_field.send_keys(application_no) # sends application no

            password_field = self.driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtPassword")
            password_field.send_keys(password) # sends the password

            captcha_field = self.driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtsecpin")
            captcha_field.send_keys() # clicks on the captcha column

            
            list_of_xpaths = ["//a[@id='ctl00_LoginContent_linkDownConfirm']", "//input[@id='ctl00_ContentPlaceHolder1_btnYes']"]
            # if any one is visible then skips the next step
            # the first one is for the inside page for downloading the response sheet link
            # the second one is for the duplicate application login handling
            self.webdriver_wait.until(EC.visibility_of_any_elements_located((By.XPATH, "|".join(list_of_xpaths))))

            
            try:
                duplicate_login= self.driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnYes")
                if duplicate_login.is_displayed():
                    # if the duplicate login detected then we click on the logout and again enter password and enter's captcha
                    duplicate_login.click()

                    self.driver.implicitly_wait(time_to_wait=3) # manually wait for 3 sec

                    password_field.send_keys(password) # sends password
                    captcha_field.find_element(By.ID, "ctl00_ContentPlaceHolder1_txtsecpin").send_keys() # click on the captcha column

                    # wait to fill the captcha manually 
                    self.webdriver_wait.until(EC.visibility_of_element_located((By.ID, "ctl00_LoginContent_linkDownConfirm")))
            
            except Exception as e:
                pass
                
            
        except exceptions.TimeoutException:
            print("Login Failed")
            return False

        else:
            self.inside_page(application_no, password)
            return True

    def inside_page(self, application_no, password) -> None:
        try:
            # getting the response sheet link
            url_link = self.driver.find_element(By.ID, "ctl00_LoginContent_rptViewQuestionPaper_ctl01_lnkviewKey").get_attribute("href")

            data_dict = {"APPLICATION_NO": application_no, "PASSWORD": password, "URL": url_link, "STATUS": "NO"}
            self.write_to_excel_file(data_dict)

        except Exception as e:
            print(f"Error inside page: {e}")
            # if any error occurs then we will save the data with URL as AB
            data_dict = {"APPLICATION_NO": application_no, "PASSWORD": password, "URL": "AB", "STATUS": "NO"}
            self.write_to_excel_file(data_dict)
        
        finally:

            self.driver.implicitly_wait(time_to_wait=3) # manually wait for 3 sec

            self.driver.find_element(By.LINK_TEXT, "Logout").click() # clicks on the logout button in the main page
            self.driver.find_element(By.ID, "btnLogout").click() # clicks on the logout button in the previous page

    def write_to_excel_file(self, data_dict: dict) -> None:
        try:
            if not self.results_saving_path.exists():
                # this is for the first time creation of the file
                pd.DataFrame([data_dict]).to_excel(self.results_saving_path, index=False)
                return
            
            #columns = ["APPLICATION_NO", "PASSWORD", "URL", "STATUS"]
            df_existing = pd.read_excel(self.results_saving_path)
            df_new = pd.DataFrame([data_dict])
            # combined the two dataframes existing then new dataframe
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            df_combined.to_excel(self.results_saving_path, index=False)

        except Exception as e:
            print(f"Exception in write_to_excel_file: {e}")
            traceback.print_exc()
            exit()

if __name__ == "__main__":
    url = "https://examinationservices.nic.in/JeeMainx2026/Root/CandidateLogin.aspx?enc=Ei4cajBkK1gZSfgr53ImFVj34FesvYg1WX45sPjGXBqfcvMYv/FHq/Da9QEnq781"
    excel_path = r"F:\2026_jee_response_sheets\credentials.xlsx"

    # columns = [APPLICATION_NO, PASSWORD, STATUS]
    jee_object = JEE_RESPONSE_SHEETS(url=url, excel_path=excel_path)
    jee_object.read_from_excel()
