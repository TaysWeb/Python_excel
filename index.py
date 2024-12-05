import pandas as pd
# from glob import glob

# File paths
file_path = "./data/Qual Analysis Download_ Gr11 End of Year 2024.xlsx"
file_path_csv = "data/Qual Analysis Download_ Gr11 End of Year 2024 - Gr11 End of Year (1).csv"

# Load Excel data
data = pd.read_excel(file_path, sheet_name="Gr11 End of Year")

# Display all rows and columns for debugging
pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None)

# Clean and format column names
data.columns = data.columns.str.strip().str.replace('\n', ' ').str.replace(r'\s+', ' ', regex=True)

# Drop rows with all NaN values
data = data.dropna(how='all')

# Define target columns for percentage calculation
percentage_columns = [
    "6.1. How would you rate the one-on-one session with your coach (tick N/A if you did not receive these sessions)",
    "6.3 How would you rate the one-on-one counselling sessions you had with a Life Choices counsellor?",
   "7.4_How likely are you to recommend Leader's Quest to a friend? (NPS) 0-6 = Detractors 7-8 = Passives 9-10 = Promoters"
]

# Ensure the target columns exist and are numeric
for col in percentage_columns:
    if col in data.columns:
        data[col] = pd.to_numeric(data[col], errors='coerce')
    else:
        print(f"Column '{col}' not found.")

# Calculate percentages for each column
percentages = data[percentage_columns].apply(lambda x: x / x.sum() * 100, axis=0)
print("Percentages per column:")
print(percentages)

# Calculate the average percentage across columns
average_percentages = percentages.mean(axis=1)
data["Average Percentage"] = average_percentages

average_percentages.plot()

print("Data with average percentages:")
print(data[["Average Percentage"]])

# Check and calculate mean for a specific column (optional)
target_column = "6.1. How would you rate the one-on-one session with your coach (tick N/A if you did not receive these sessions)"
if target_column in data.columns:
    mean_value = data[target_column].mean()
    print(f"Mean Value: {mean_value}")
else:
    print(f"Column '{target_column}' not found. Available columns:")
    print(data.columns.tolist())

# Summarize the data
print(data.describe())

# Load only the first 104 rows
data_subset = pd.read_excel(file_path, sheet_name="Gr11 End of Year", nrows=104)
print(data_subset.head())

# Analyze Question 6.2
question_6_data = data.iloc[:, 78:83]  # Replace with dynamic column selection if possible
top_responses = question_6_data.apply(pd.Series.value_counts).head(3)
summary_stats = question_6_data.describe(include="all")

print("Top Responses:")
print(top_responses)
print("Summary Statistics:")
print(summary_stats)

# Load CSV data for further analysis
try:
    data_q6 = pd.read_csv(file_path_csv)
    print(data_q6.head())
except FileNotFoundError:
    print(f"CSV file not found at: {file_path_csv}")
except Exception as e:
    print(f"Error loading CSV: {e}")

# Combine Excel files in a folder (if needed)
# path = "./data/"
# files = sorted(glob(path + "/*.xlsx"))
# if files:
#     combined_data = pd.concat([pd.read_excel(f) for f in files], ignore_index=True)
#     print(combined_data.head(1000))
# else:
#     print("No Excel files found to process.")
