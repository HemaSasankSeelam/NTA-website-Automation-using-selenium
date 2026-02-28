
from pathlib import Path
"""
files is downloading with .crdownload so we need to change to .pdf
"""
path = r"F:\2026_jee_applications"

input_folder_path = Path(path)

for each_file in input_folder_path.iterdir():

    if each_file.suffix == ".crdownload":
        new_name = each_file.as_posix().replace(".crdownload","")
        each_file.rename(new_name)
    
    elif each_file.suffix == ".tmp":
        new_name = each_file.as_posix().replace(".tmp",".pdf")
        each_file.rename(new_name)