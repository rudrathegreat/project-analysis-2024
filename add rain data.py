import pandas as pd

# fix dates
def convert_date(date_str):
    return pd.to_datetime(f"2024-{date_str[3:]}-{date_str[:2]}")

# List rain data files
csv_files = ['flemingtonBOM_modified.csv', 'flemingtonMelbW_modified.csv']

# List audit data files
excel_files = ['Audit_1.xlsx', 'Audit_2.xlsx', 'Audit_3.xlsx', 'Audit_4.xlsx', 'Audit_5.xlsx', 'Audit_6.xlsx', 'Audit_7.xlsx']

for file in excel_files:
    # Read audit file
    xl = pd.ExcelFile(file)
    
    sheet_name = file.split('.')[0]
    site_audit_sheet = f'Site audit {sheet_name.split('_')[1]}' 
    
    # get audit data
    site_data = xl.parse(site_audit_sheet)
    
    # get dates
    def parse_date(date):
        try:
            return pd.to_datetime(date)
        except:
            try:
                return pd.to_datetime(date, format='%d-%m-%Y')
            except:
                return pd.to_datetime(date, format='%m/%d/%Y')

    date_of_audit = parse_date(site_data['Date of audit'].iloc[0])
    last_audited_str = site_data['Last audited'].iloc[0]
    
    # audit 1 last audit
    if last_audited_str == 'unknown':
        last_audited = date_of_audit - pd.Timedelta(weeks=3)
    else:
        last_audited = parse_date(last_audited_str)
    
    # write data
    for csv_file in csv_files:
        csv_data = pd.read_csv(csv_file)
        csv_data['Date'] = csv_data['Date'].apply(convert_date)
        filtered_data = csv_data[(csv_data['Date'] > last_audited) & (csv_data['Date'] <= date_of_audit)]
        new_sheet_name = f'{csv_file}'
        with pd.ExcelWriter(file, mode='a', engine='openpyxl') as writer:
            filtered_data.to_excel(writer, sheet_name=new_sheet_name, index=False)

        print(f"Processed {file} with {csv_file}")

print("All files processed!")
