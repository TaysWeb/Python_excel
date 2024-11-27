import pandas as pd

# load the excel file

file_path =  r"C:\Users\Tara Snell\Documents\Python_data\data\Qual Analysis Download_ Gr11 End of Year 2024.xlsx "
data = pd.read_excel(file_path)

data.info();

# Load only the first 10,000 rows.
# data = pd.read_excel(file_path, nrows = 105)  

# total = data['NumericColumn'].sum()
# print(f"Total: {total}")

# print(data.head())