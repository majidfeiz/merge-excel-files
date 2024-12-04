# importing the required modules
import glob
import pandas as pd

# specifying the path to excel files
path = "/Users/sermajid/Desktop/Project/Python/MergeExcelFiles/Por/"

# excel files in the path
file_list = glob.glob(path + "*.xlsx")

# list of excel files we want to merge.
# pd.read_excel(file_path) reads the excel
# data into pandas dataframe.
excl_list = []

for file in file_list:
    try:
        df = pd.read_excel(file, engine='openpyxl')
        excl_list.append(df)
    except Exception as e:
        print(f"Error reading {file}: {e}")

# create a new dataframe to store the 
# merged excel file.
if excl_list:
    excl_merged = pd.concat(excl_list, ignore_index=True)

    # exports the dataframe into excel file with
    # specified name.
    excl_merged.to_excel('result2.xlsx', index=False)
else:
    print("No valid Excel files found to merge.")
