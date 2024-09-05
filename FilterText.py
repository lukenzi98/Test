
import pandas as pd

# Load the Excel file
excel_file = r"C:\Users\e5723514\OneDrive - FIS\Luka\Py projects\Host_Compliance.xlsx"

# Load Sheet 1 and Sheet 2 into DataFrames
sheet1 = pd.read_excel(excel_file, sheet_name='Sheet1')
sheet2 = pd.read_excel(excel_file, sheet_name='Sheet2')

# Convert data in the first column of Sheet 2 to uppercase
sheet2_values = sheet2.iloc[:, 0].str.upper().tolist()

# Convert the 3rd column of Sheet 1 to uppercase and filter based on the list from Sheet 2
filtered_data = sheet1[sheet1.iloc[:, 2].str.upper().isin(sheet2_values)]

# Convert the entire DataFrame to uppercase
filtered_data = filtered_data.applymap(lambda s: s.upper() if type(s) == str else s)

# Save the filtered data to a new sheet or a new file
filtered_data.to_excel('filtered_output.xlsx', index=False, sheet_name='FilteredData')

print("Filtered data has been saved to 'filtered_output.xlsx'.")


