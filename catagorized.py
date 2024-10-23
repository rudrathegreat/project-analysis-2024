import pandas as pd
import re

# Get categories
def extract_and_capitalize(text):
    if not isinstance(text, str):
        text = str(text)
    return ''.join(re.findall('[a-zA-Z]+', text)).upper()

# audit data files
excel_files = ['Audit_1.xlsx', 'Audit_2.xlsx', 'Audit_3.xlsx', 'Audit_4.xlsx', 'Audit_5.xlsx', 'Audit_6.xlsx', 'Audit_7.xlsx']

for file in excel_files:
    xl = pd.ExcelFile(file)
    
    sheet_name_template = file.split('.')[0]
    sheet_name = sheet_name_template.replace('_', ' ')
    
    # read data
    df = xl.parse(sheet_name)

    #create categories
    df['Categories'] = df['Litter Code'].apply(extract_and_capitalize)
    
    # save file
    with pd.ExcelWriter(file, mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("Categories column with capitalized letters added to the Excel files.")
