
import pandas as pd

# from glob import glob 


# load the excel file
file_path =  "./data/Qual Analysis Download_ Gr11 End of Year 2024.xlsx"
data = pd.read_excel(file_path)

data = pd.read_excel(file_path,sheet_name= "Gr11 End of Year")

print(data)

pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None)

#Calculatig the mean for a column
data.columns = data.columns.str.strip().str.replace('\n', ' ').str.replace(r'\s+', ' ', regex=True)

cleaned_data = data.dropna( how='all')

        
    
print(data["6.1. How would you rate the one-on-one session with your coach (tick N/A if you did not receive these sessions)"].dtype)

data["6.1. How would you rate the one-on-one session with your coach (tick N/A if you did not receive these sessions)"] = pd.to_numeric(
    data["6.1. How would you rate the one-on-one session with your coach (tick N/A if you did not receive these sessions)"], errors='coerce'
)

        
mean_value =data["6.1. How would you rate the one-on-one session with your coach (tick N/A if you did not receive these sessions)"].mean()
print(f"Mean Value: {mean_value}")




# error handling 
if "6.1. How would you rate the one-on-one session with your coach (tick N/A if you did not receive these sessions)" in data.columns:
    mean_value = data["5.6. Having Someone supportive to talk to about my emotions"].mean()
    print(f"Mean Value: {mean_value}")
else:
    print("Column not found. Available columns are:")
    print(data.columns.tolist())



#summarize the data
print(data.describe())
# drop rows with missing data
cleaned_data = data.dropna( how='all')

#Load only the first 500 rows.
data_subset = pd.read_excel(file_path, nrows=104)
print(data_subset.head(104))

# print(data.columns)
# print(data.index)



#########   ANALYSIS CODE   #######
 
 
# Extract columns CA to CF for Question 6.2
question_6_data = data.iloc[:, 78:83]  # Adjusted based on column positions (CA to CF)

# Keeping original comments and analyzing top 3 responses based on frequency or sentiment
top_responses = question_6_data.apply(pd.Series.value_counts).head(3)

# Calculating a summary: Transforming qualitative comments into ratings (if applicable) and summarizing
# Assuming an approach to calculate a "rating" for qualitative comments is feasible or ratings are numeric
summary_stats = question_6_data.describe(include="all")

# Display results
question_6_data.head(), top_responses, summary_stats











# combine multiple Excel files from a folder
# path = "./data/"
# files = sorted(glob(path + "/*.xlsx"))
# print( f"Files to process: {files}")

# combining both excel sheets to one
# data = pd.read_excel(file_path , sheet_name=None)
# data_combined_ = pd.concat(data.values(), ignore_index=True)

# if files:
#   data_combined = pd.concat([pd.read_excel(f, index_col=0, parse_dates=[0]) for f in files])
#   print(data_combined.head(1000))
  
# else:
#     print("No files found to process.") 


