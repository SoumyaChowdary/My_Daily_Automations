import pandas as pd
import openpyxl
import os

# Load the template and the data file
template_path = 'Path_to_template_file'
data_path = 'Path_to_file_that_contains_gading_information_and_Student_info'

# Read the data from the Excel files
data = pd.read_excel(data_path)

# Load the template workbook
template_wb = openpyxl.load_workbook(template_path)

# Directory to save the new Excel files
output_dir = 'Folder_to_save_the_generated_excel_files'
os.makedirs(output_dir, exist_ok=True)

# Loop through each row in the data file
for index, row in data.iterrows():
    # Create a new workbook from the template
    new_wb = openpyxl.load_workbook(template_path)
    sheet = new_wb.active
    
    # Update the specific cells with the information from the data file
    sheet['C2'] = f"ID: {row['ID']}"
    sheet['E1'] = f"Name: {row['Name']}"
    sheet['E2'] = f"Email: {row['Email']}"
    #update marks
    sheet['D13'] = row['A']
    sheet['D14'] = row['B']
    sheet['D15'] = row['C']
    sheet['D16'] = row['D']
    sheet['D17'] = row['E']

    #update comments
    sheet['E13'] = row['AA']
    sheet['E14'] = row['BB']
    sheet['E15'] = row['CC']
    sheet['E16'] = row['DD']
    sheet['E17'] = row['EE']
    
    # Save the new workbook with the name provided in the data file
    output_file = os.path.join(output_dir, f"{row['Name'].replace(' ', '_')}.xlsx")
    new_wb.save(output_file)
    # break

print("Files generated successfully.")
