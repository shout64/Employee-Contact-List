import pandas as pd
import openpyxl as ox

# Read CSV, update column names and sort by Name
df = pd.read_csv(r'C:\members.csv')
df = df.rename(columns={'displayName': 'Name', 'title': 'Title', 'department': 'Department', 'physicalDeliveryOfficeName': 'Office Location', 'mail': 'Email', 'telephoneNumber': 'Phone Number'})
df = df.sort_values('Name')

# Create Excel Doc
df.to_excel(r'C:\Contact List.xlsx', sheet_name='Contacts', index=False)

# Adjust column width
workbook = ox.load_workbook(r'C:\Contact List.xlsx')
worksheet = workbook['Contacts']
worksheet.column_dimensions['A'].width = 25
worksheet.column_dimensions['B'].width = 65
worksheet.column_dimensions['C'].width = 48
worksheet.column_dimensions['D'].width = 40
worksheet.column_dimensions['E'].width = 40
worksheet.column_dimensions['F'].width = 20

workbook.save(r'C:\Contact List.xlsx')
