import selenium
import traceback
from selenium import webdriver
import pandas as pd
from pathlib import Path

from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait


class JEE_RESPONSE_VERIFICATION:

    def __init__(self, path):

        self.path = path # response sheet path
        chrome_service = Service(executable_path=ChromeDriverManager().install()) # chrome driver manager

        self.driver = webdriver.Chrome(service=chrome_service) # webdriver instance
        self.driver.set_window_position(x=200, y=50) # sets the window position
        self.wait = WebDriverWait(driver=self.driver, timeout=12) # webdriver wait

        self.df = pd.read_excel(self.path) # reads the excel
        self.filtered_df = self.df.loc[self.df["STATUS"] == "NO"] # filters the dataframe whose status is NO
        self.filtered_df = self.filtered_df.loc[self.filtered_df["URL"] != "AB"] # removes the invalid urls

        if len(self.filtered_df) == 0:
            print("ALL Completed".center(50, "*"))
            return
        
        self.application_nos: list[int] = self.filtered_df["APPLICATION_NO"].tolist()

        for application_no in self.application_nos:

            # extract the URL from the dataframe
            url = self.filtered_df.loc[self.filtered_df["APPLICATION_NO"] == application_no, ["URL"]].values[0][0]

            if not str(url).startswith("https://"): # adding the https to prefix of the url if not present
                url = "https://" + url 

            self.driver.get(url=url) # get the response sheet url

            if self.inside_page(url=url):
                # update the status to YES if all correctly done
                self.df.loc[self.df["APPLICATION_NO"] == application_no,["STATUS"]] = "YES"
                self.df.to_excel(self.path, index=False)
            

        print(" COMPLETED ".center(50, "*"))

    def inside_page(self, url):

        try:
            basic_details_table = self.driver.find_element(by=By.TAG_NAME, value="table") # find the element by tagname
            person_details = {}

            last_key = None
            # if the index is even then its store as key.
            # if the index is odd then its store as value.
            for index,value in enumerate(basic_details_table.find_elements(by=By.TAG_NAME, value="td")):

                if index%2 == 0:
                    last_key = value.text.lower()
                    person_details[last_key] = ""

                else:
                    person_details[last_key] = value.text.lower()
            
            question_answer_ids = {}
            subject_and_type_of_questions = []
            answered_types = [] # helps to capture marked for review questions also

            for index1,each_element in enumerate(self.driver.find_elements(by=By.XPATH, value='//table[@class="menu-tbl"]')):
        
                options_menu = {}
                """
                options menu consists of the 
                question type, status, chosen answer and answer id.
                for mcq questions the answer id is the option id of the chosen answer.
                for numerical questions the answer id is the answer given by the candidate.
                """
                last_key = None
                for index2,value in enumerate(each_element.find_elements(by=By.TAG_NAME, value="td")):

                    if index2%2 == 0:
                        last_key = value.text.lower()
                        options_menu[last_key] = ""

                    else:
                        options_menu[last_key] = value.text.lower()
                
                question_id = list(options_menu.values())[1]
                if options_menu["question type :"] == "mcq" and options_menu["status :"] == "answered":
                    chosen_answer = int(list(options_menu.values())[-1])
                    answer_id = list(options_menu.values())[chosen_answer+1]
                    answered_types.append("mcq_answered")

                elif options_menu["question type :"] == "mcq" and options_menu["status :"] == "marked for review":
                    # for marked for review will not get the marks even if the answer is correct.


                    # if you want marked for review marks marks to count
                    # remove the answer_id = "--" and uncomment the below two lines

                    # chosen_answer = int(list(options_menu.values())[-1])
                    # answer_id = list(options_menu.values())[chosen_answer+1]

                    answer_id = "--"
                    answered_types.append("mcq_answered_marked_for_review")

                elif options_menu["question type :"] != "mcq" and options_menu["status :"] == "answered":
                    # this is for numerical questions answered
                    ele = self.driver.find_elements(by=By.XPATH, value='//table[@class="questionRowTbl"]')[index1]
                    last_tr_row = ele.find_elements(by=By.TAG_NAME, value="tr")[-1]

                    answer_id = last_tr_row.text.split(":")[-1].strip()
                    answered_types.append("numerical_answered")
                
                else:
                    # for all the other cases like not answered and marked for review not answered
                    # answer id is set to --
                    answer_id = "--"
                    answered_types.append("not_answered")

                subject_and_type = each_element.find_element(by=By.XPATH, value='.//ancestor::div[2]//span[2]').text
                subject_and_type_of_questions.append(subject_and_type)
                question_answer_ids[question_id] = answer_id

            new_dict = {}
            new_dict["SUBJECT_TYPE_OF_QUESTION"] = subject_and_type_of_questions
            new_dict["QUESTION_IDS"] = list(question_answer_ids.keys())
            new_dict["ANSWER_IDS"] = list(question_answer_ids.values())
            new_dict["ANSWERED_TYPE"] = answered_types

            new_df = pd.DataFrame(data=new_dict)
            length_of_new_dataframe = len(new_df)

            # inserting the additional columns to the dataframe in the front side                               
            new_df.insert(loc=0, column="APPLICATION_NO", value=[person_details["application no"]]*length_of_new_dataframe)
            new_df.insert(loc=1, column="TEST_DATE", value=[person_details["test date"]]*length_of_new_dataframe)
            new_df.insert(loc=2, column="TEST_TIME", value=[person_details["test time"]]*length_of_new_dataframe)
            new_df.insert(loc=3, column="URL", value=[url]*length_of_new_dataframe)
            new_df.insert(loc=4, column="QUESTION_NO", value=list(range(1,length_of_new_dataframe+1)))
            new_df["QUESTION_NO"] = new_df["QUESTION_NO"].astype(dtype="object")

            # formatting the test date to dd-mm-yyyy format
            new_df["TEST_DATE"] = pd.to_datetime(new_df["TEST_DATE"], dayfirst=True, errors="coerce")
            new_df["TEST_DATE"] = new_df["TEST_DATE"].dt.strftime("%d-%m-%Y")

            self.make_excel_sheet(data_frame=new_df)
            return True

        except Exception as e:
            print("Error in the main page")
            traceback.print_exc()
            return False

    def make_excel_sheet(self, data_frame):
        
        try:

            test_date = str(data_frame["TEST_DATE"].iloc[0])
            # determine the test phase based on the test time
            test_phase = "phase-1" if "am" in str(data_frame["TEST_TIME"].iloc[0]) else "phase-2"

            current_folder_name = Path(self.path).parent # gets the parent folder name
            new_folder_path = Path(current_folder_name) / test_date # making the parent folder name / test_date
            if not Path(new_folder_path).exists():
                Path(new_folder_path).mkdir(parents=True, exist_ok=True) # creates the folder if not exists
           
            current_folder_name = new_folder_path
            new_folder_path = current_folder_name / test_phase
            if not Path(new_folder_path).exists():
                Path(new_folder_path).mkdir(parents=True, exist_ok=True) # creates the folder if not exists
            
            current_folder_name = new_folder_path
            application_no = str(data_frame["APPLICATION_NO"].iloc[0])
            new_folder_path = current_folder_name / application_no
            if not(Path(new_folder_path)).exists():
                Path(new_folder_path).mkdir(parents=True, exist_ok=True) # creates the folder if not exists
        
            current_folder_name = new_folder_path
            new_excel_path = Path(current_folder_name) / f"{application_no}.xlsx" # new excel file path with application no
            
            # columns = [APPLICATION_NO, TEST_DATE, TEST_TIME, URL, QUESTION_NO, SUBJECT_TYPE_OF_QUESTION, QUESTION_IDS, ANSWER_IDS, ANSWERED_TYPE]
            pd.DataFrame(data=data_frame).to_excel(new_excel_path, index=False)
            return True
         
        except Exception as e:
            print("Error in the Excel sheet")
            traceback.print_exc()
            return False

        
# columns = [APPLICATION_NO, URL, STATUS]
response_sheet_path = r"F:\2026_jee_response_sheets\credentials_results_by_selenium.xlsx"
JEE_RESPONSE_VERIFICATION(path=response_sheet_path)

