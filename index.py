import pandas as pd
from glob import glob 

# load the excel file

file_path =  "./data/Qual Analysis Download_ Gr11 End of Year 2024.xlsx"
data = pd.read_excel(file_path)

#summarize the data
print(data.describe());
sheets = ["Gr11 End of Year", "HCT 2024 Survey"]
data = pd.read_excel(file_path,sheet_name= None)


# #Load only the first 105 rows.
data_subset = pd.read_excel(file_path, nrows=105)
print(data_subset.head(250))

# combine multiple Excel files from a folder
path = "./data/"
files = sorted(glob(path + "/*.xlsx"))
print( f"Files to process: {files}")

# combining both excel sheets to one
data = pd.read_excel(file_path , sheet_name=None)
data_combined_ = pd.concat(data.values(), ignore_index=True)

if files:
  data_combined = pd.concat([pd.read_excel(f, index_col=0, parse_dates=[0]) for f in files])
  print(data_combined.head(100))
  
else:
    print("No files found to process.") 


