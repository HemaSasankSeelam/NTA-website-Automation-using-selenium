import pandas as pd
from pathlib import Path
import wrong_answers

class FINAL_KEY:

    def __init__(self, final_key_path, take_images, required_students_path=None):
        
        self.final_key_path = Path(final_key_path) # creates a path for the final key excel file
        self.take_images = take_images # if you want to take the images of wrong answers then set it to True else set it to False
        saved_folder = self.final_key_path.parent # creates a path for the saved folder where we are going to save the merged excel file and the images of wrong answers
        self.merged_excel_path = saved_folder / Path("merged.xlsx") # creates a excel path
        self.wrong_answers_module = wrong_answers.WRONG_ANSWERS() # importing the wrong answers module

        if required_students_path:
            # if you want to verify the keys for the required students then provide the path of the excel file 
            self.required_students_path:Path = Path(required_students_path)
    
        if not required_students_path:
            sheet_names = pd.ExcelFile(self.final_key_path).sheet_names # gets all the sheet names in the final key excel file

            for each_sheet_name in sheet_names:
                test_date, test_phase = each_sheet_name.split(" ") # splits the test date and test phase
                excel_files_folder = self.final_key_path.parent / f"{test_date}/{test_phase}" 

                keys_df = pd.read_excel(self.final_key_path, sheet_name=f"{test_date} {test_phase}") # reads the key df for the test date and test phase

                if not excel_files_folder.exists(): 
                    # checks if the test date and test phase exists
                    # if not exists then it means there are no response sheets for that test date and test phase so we can skip it
                    continue

                for each_folder in excel_files_folder.iterdir():
                    # iter to the each folder in the test date and test phase folder
                    if each_folder.is_file():
                        continue

                    student_id = str(each_folder.name) # gets the student id from the folder name
                    excel_path = each_folder / Path(student_id + ".xlsx") # creates the excel path for the student response sheet
                    
                    is_application_no_contains = False
                    if self.merged_excel_path.exists():
                        merged_df = pd.read_excel(self.merged_excel_path)
                        if len(merged_df) > 0 and int(each_folder.stem) in merged_df["APPLICATION_NO"].to_list():
                            # if the application no is already present in the merged excel then skip it
                            # else we can compare the excel files
                            is_application_no_contains = True

                    if not is_application_no_contains:
                        # compare the excel files
                        questions_df = pd.read_excel(excel_path)
                        self.excel_compare(keys_df=keys_df, question_df=questions_df, excel_path=excel_path)
                    
                    if self.take_images:
                        # if you want to take the images of wrong answers
                        self.wrong_answers_module.take_screen_shot(student_excel_path=excel_path)

            
            print(" ALL COMPLETED ".center(50, "*"))
            return

        student_df = pd.read_excel(self.required_students_path)
        filtered_student_df = student_df.loc[student_df["STATUS"] == "NO"] # filters the dataframe whose status is NO

        if len(filtered_student_df) == 0:
            print("ALL THE STATUS ARE YES")
            print("COMPLETED")
            return
        
        application_nos = filtered_student_df["APPLICATION_NO"].tolist()
        filtered_student_df = filtered_student_df.copy()
        for application_no in application_nos:
            # iter through each application no
            filtered_student_df["TEST_DATE"] = pd.to_datetime(filtered_student_df["TEST_DATE"], dayfirst=True, errors="coerce", format="mixed")
            filtered_student_df["TEST_DATE"] = filtered_student_df["TEST_DATE"].dt.strftime("%d-%m-%Y")

            test_date = filtered_student_df.loc[filtered_student_df["APPLICATION_NO"] == application_no,["TEST_DATE"]].values[0][0]
            test_time = filtered_student_df.loc[filtered_student_df["APPLICATION_NO"] == application_no,["TEST_TIME"]].values[0][0]
            test_phase = "phase-1" if "am" in str(test_time) else "phase-2"

            # makes the excel path to store
            excel_path = self.final_key_path.parent / Path(f"{test_date}/{test_phase}/{application_no}/{application_no}.xlsx")

            is_application_no_contains = False
            if self.merged_excel_path.exists():
                merged_df = pd.read_excel(self.merged_excel_path)
                if len(merged_df) > 0 and application_no in merged_df["APPLICATION_NO"].to_list():
                    # if the application no is already present in the merged excel then skip it
                    # else we can compare the excel files
                    is_application_no_contains = True
    
            if not is_application_no_contains:
                keys_df = pd.read_excel(self.final_key_path, sheet_name=f"{test_date} {test_phase}")
                questions_df = pd.read_excel(excel_path)
                # compare the excel files
                self.excel_compare(keys_df=keys_df, question_df=questions_df, excel_path=excel_path)

                
            if self.take_images:
                # you can take images for the wrong answers
                self.wrong_answers_module.take_screen_shot(student_excel_path=excel_path)
                student_df.loc[student_df["APPLICATION_NO"] == application_no,["STATUS"]] = "YES"
                student_df.to_excel(self.required_students_path, index=False)

        print("*** ALL COMPLETED ***")


    def excel_compare(self, keys_df:pd.DataFrame, question_df:pd.DataFrame, excel_path:Path):

        # compares the key df and question df and updates the question df with correct answer ids and marks
        questions_ids_in_key = keys_df["QUESTION_IDS"].values.tolist()
        question_df["CORRECT_ANSWER_IDS"] = -1 # by default all the correct answer ids are -1
        question_df["MARKS"] = 0 # by default all the marks are 0
        question_df["CORRECT_ANSWER_IDS"] = question_df["CORRECT_ANSWER_IDS"].astype("object")

        for each_question_id in questions_ids_in_key:
            
            # for multiple keys we are going to evaluate the org answer id
            # and then check if the kept answer id is in the org answer ids
            # if yes then we are going to give full marks else negative marks

            org_answer_id = str(keys_df.loc[keys_df["QUESTION_IDS"] == each_question_id,["ANSWER_IDS"]].values[0][0]).lower().replace("or",",")
            org_answers_ids = []

            if org_answer_id != "dropped" and type(eval(org_answer_id)) == tuple: # checks the type of the evaluated org answer
                # if there are multiple keys then it going to be tuple
                # else it is a string datatype
                org_answers_ids = eval(org_answer_id) # evaluates the tuple to list of answers
                org_answers_ids = list(map(lambda x:str(float(x)),org_answers_ids)) # converts all the answers to string


                answer = question_df.loc[question_df["QUESTION_IDS"] == each_question_id,["ANSWER_IDS"]].values[0][0]
                if answer != "--":
                    kept_answer_id = str(float(answer))
                else:
                    kept_answer_id = "--"

                # we changing the correct answer ids to " or " separated string
                question_df.loc[question_df["QUESTION_IDS"] == each_question_id,["CORRECT_ANSWER_IDS"]] = " or ".join(map(str,org_answers_ids))

            else:
                # else for single key
                kept_answer_id = str(question_df.loc[question_df["QUESTION_IDS"] == each_question_id,["ANSWER_IDS"]].values[0][0])
                # changing the correct answer ids to org answer id
                question_df.loc[question_df["QUESTION_IDS"] == each_question_id,["CORRECT_ANSWER_IDS"]] = str(org_answer_id)

            if org_answers_ids and kept_answer_id == "--":
                question_df.loc[question_df["QUESTION_IDS"] == each_question_id,["MARKS"]] = 0

            elif org_answers_ids and kept_answer_id in org_answers_ids:
                question_df.loc[question_df["QUESTION_IDS"] == each_question_id,["MARKS"]] = 4

            elif org_answers_ids and kept_answer_id not in org_answers_ids:
                question_df.loc[question_df["QUESTION_IDS"] == each_question_id,["MARKS"]] = -1


            elif org_answer_id == kept_answer_id or str(org_answer_id).lower() == "dropped":
                # dropped is also considered as correct answer
                # it means the question is dropped by the exam conducting body
                question_df.loc[question_df["QUESTION_IDS"] == each_question_id,["MARKS"]] = 4

            elif kept_answer_id == "--":
                question_df.loc[question_df["QUESTION_IDS"] == each_question_id,["MARKS"]] = 0

            elif org_answer_id != kept_answer_id:
                question_df.loc[question_df["QUESTION_IDS"] == each_question_id,["MARKS"]] = -1

        question_df.to_excel(excel_path, index=False) 

        self.group_by(excel_path=excel_path)

    def group_by(self, excel_path):
        
        question_df = pd.read_excel(excel_path)
        application_no = question_df["APPLICATION_NO"].iloc[0]
        test_date = question_df["TEST_DATE"].iloc[0]
        test_time = question_df["TEST_TIME"].iloc[0]

        subject_keys = ["MATHS_CORRECT", "MATHS_WRONG", "MATHS_NOT_ATTEMPTED/MARKED_FOR_REVIEW", "MATHS_ESTIMATED_MARKS",
                        "PHYSICS_CORRECT", "PHYSICS_WRONG", "PHYSICS_NOT_ATTEMPTED/MARKED_FOR_REVIEW", "PHYSICS_ESTIMATED_MARKS",
                        "CHEMISTRY_CORRECT", "CHEMISTRY_WRONG", "CHEMISTRY_NOT_ATTEMPTED/MARKED_FOR_REVIEW", "CHEMISTRY_ESTIMATED_MARKS"]
        subject_values = []
        total_marks = 0
        for each_subject in question_df["SUBJECT_TYPE_OF_QUESTION"].str.split(" ",expand=True)[0].unique():
            
            subject = question_df.loc[question_df["SUBJECT_TYPE_OF_QUESTION"].str.match(pat=r"^{}".format(each_subject))]

            correct = subject.loc[subject["MARKS"] == 4]["MARKS"].count()
            wrong = subject.loc[subject["MARKS"] == -1]["MARKS"].count()
            not_attempted = subject.loc[subject["MARKS"] == 0]["MARKS"].count()
            each_subject_estimated_marks = (correct * 4) - (wrong * 1)

            total_marks += each_subject_estimated_marks
            
            subject_values.extend([correct, wrong, not_attempted, each_subject_estimated_marks])

        data_dict = {"APPLICATION_NO":application_no, "TEST_DATE":test_date, "TEST_TIME":test_time} | dict(zip(subject_keys, subject_values)) | {"TOTAL_ESTIMATED_MARKS":total_marks}

        if not self.merged_excel_path.exists():
            # columns = [APPLICATION_NO, TEST_DATE, TEST_TIME, 
            #           "MATHS_CORRECT", "MATHS_WRONG", "MATHS_NOT_ATTEMPTED/MARKED_FOR_REVIEW", "MATHS_ESTIMATED_MARKS",
            #            "PHYSICS_CORRECT", "PHYSICS_WRONG", "PHYSICS_NOT_ATTEMPTED/MARKED_FOR_REVIEW", "PHYSICS_ESTIMATED_MARKS",
            #            "CHEMISTRY_CORRECT", "CHEMISTRY_WRONG", "CHEMISTRY_NOT_ATTEMPTED/MARKED_FOR_REVIEW", "CHEMISTRY_ESTIMATED_MARKS"
            #            "TOTAL_ESTIMATED_MARKS"]
            df = pd.DataFrame(data=[data_dict])
            df.to_excel(self.merged_excel_path, index=False)

        else:
            # appends the data to the merged excel file
            df = pd.read_excel(self.merged_excel_path)
            df.loc[len(df)] = list(data_dict.values())
            df.to_excel(self.merged_excel_path, index=False)

# key df columns = [TEST_DATE, TEST_TIME, QUESTION_IDS, ANSWER_IDS]
# sheet names = "date-month-year phase-1","date-month-year phase-2"
# question in this we place the question_id space answer_id

# final_key_path = r"F:\response_sheets\final key.xlsx"
# sheet name dd-mm-yyyy phase-no
# columns = [TEST_DATE, TEST_TIME, QUESTION_IDS, ANSWER_IDS]

final_key_path = r"F:\2026_jee_response_sheets\final key.xlsx"
take_images = True  # set the True if you want to take the images of wrong answers else set False

# columns = [APPLICATION_NO, TEST_DATE,	TEST_TIME, STATUS]
# required_students_path = r"F:\response_sheets\required_students.xlsx" 
required_students_path = r"F:\~today\top2.xlsx"


FINAL_KEY(final_key_path=final_key_path,
          take_images=take_images)