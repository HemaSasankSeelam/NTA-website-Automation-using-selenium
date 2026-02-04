
import selenium

from selenium.common import exceptions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

import pandas as pd
from pathlib import Path
import traceback 
import time
import pyhtml2pdf.converter
import pdfkit

class WRONG_ANSWERS:

    def __init__(self):

        chrome_service = Service(executable_path=ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=chrome_service)
        self.webdriver_wait = WebDriverWait(driver=self.driver, timeout=7)

    def take_screen_shot(self,student_excel_path):

        self.student_excel_path = Path(student_excel_path)
        self.student_df = pd.read_excel(student_excel_path)

        application_no = self.student_excel_path.stem # get the application no from the file name
        base_folder = self.student_excel_path.parent.parent # goes back 2 back parent

        if (base_folder / f"{application_no}.pdf").exists():
            # means already pdf generated
            return

        self.images_folder = self.student_excel_path.parent / "wrong_questions"
        self.images_folder.mkdir(parents=True, exist_ok=True)

        filtered_df = self.student_df.loc[:,["APPLICATION_NO","TEST_DATE","TEST_TIME","URL","QUESTION_NO","QUESTION_IDS","ANSWERED_TYPE","MARKS"]]

        marks_filtered_df = filtered_df.loc[filtered_df["MARKS"] != 4] # filters the dataframe whose marks is not 4

        application_no = filtered_df.loc[0,"APPLICATION_NO"]
        test_date = filtered_df.loc[0,"TEST_DATE"]
        test_time = filtered_df.loc[0,"TEST_TIME"]

        question_nos_based_on_marks = marks_filtered_df["QUESTION_NO"].to_list() # gets the question nos whose marks is not 4

        question_nos_based_on_answered_type = filtered_df.loc[filtered_df["ANSWERED_TYPE"] == "mcq_answered_marked_for_review",["QUESTION_NO"]]\
                                                         .values.flatten().tolist() # gets the question nos whose answered type is marked for review
         
        sorted_question_ids = sorted(list(set(question_nos_based_on_marks + question_nos_based_on_answered_type)))
        
        question_ids = filtered_df.loc[filtered_df["QUESTION_NO"].isin(sorted_question_ids),["QUESTION_IDS"]].values.flatten().tolist() # gets the question ids whose question nos is in the sorted question nos

        url = filtered_df["URL"].iloc[0]

        self.driver.get(url=url)
        self.driver.maximize_window()
    
        time.sleep(3)
        application_heading_img = self.driver.find_element(by=By.XPATH, value='//div[@class="main-info-pnl"]')
        application_heading_img.screenshot((self.images_folder/"00.png").as_posix())

        html_code = f"""
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>{application_no}</title>
        </head>
        <body>
        <div class="head" style="display: flex; justify-content:space-evenly; background-color: rgb(113, 183, 211);">
            <h2>Test Date: {test_date}</h2>
            <h2>Application No: {application_no}</h2>
            <h2>Test Time: {test_time}</h2>
        </div>
        <br>
        <img src="./00.png" alt="Unable to display main heading" style='display:block; border-bottom:2px solid black;'\n>
        <br>
        <hr>
        """
        maths_entered = False
        physics_entered = False
        chemistry_entered = False

        for question_id in question_ids:

            question_no = int(filtered_df.loc[filtered_df["QUESTION_IDS"] == question_id,["QUESTION_NO"]].values[0][0])

            question_image = self.driver.find_elements(by=By.XPATH, value="//table[@class='questionPnlTbl']")[question_no-1]
            question_image_path = self.images_folder / f"{question_no:02d}.png"
            question_image.screenshot(question_image_path.as_posix())

            if question_no <= 25 and not maths_entered:
                maths_entered = True
                html_code += "\n<hr><br><h2 style='text-align:center;'>MATHEMATICS</h2><br>"
            elif 25 < question_no <= 50 and not physics_entered:
                physics_entered = True
                html_code += "\n<br><h2 style='text-align:center;' class='break'>PHYSICS</h2><hr><br>"
            elif 50 < question_no <= 75 and not chemistry_entered:
                chemistry_entered = True
                html_code += "\n<br><h2 style='text-align:center;' class='break'>CHEMISTRY</h2><hr><br>"

            html_code += f"\n<img src='./{question_image_path.name}' alt='Unable to display question no {question_image_path.stem}' style='display:block; border-bottom:2px solid black;'>\n"
        
        html_code += """
        <style>
            hr {
                border: 2px dashed pink;
            }
            @media print
            {
                .break
                {
                    page-break-before: always;
                }
            }
        </style>
        </body>
        </html>
        """
        html_file_path = (self.images_folder/"index.html").as_posix()
        with open(html_file_path,'w') as fo:
            fo.write(html_code)

        pdf_path = (self.images_folder.parent.parent/f"{application_no}.pdf").as_posix()
        # pyhtml2pdf.converter.convert(source=html_file_path, target=pdf_path,
        #                              print_options={"paperWidth":11})
        
        # convert the html to pdf using pdfkit module
        pdfkit.from_file(input = html_file_path,
                        output_path = pdf_path,
                        options={'enable-local-file-access': '',
                                'page-size': 'A4'})

        
if __name__ == "__main__":
    pass
    # WRONG_ANSWERS(r"F:\response_sheets\22-01-2025\phase-1\application_no\application_no.xlsx")