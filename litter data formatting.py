import pandas as pd

def read_litter_audit_sheets(file_paths, sheets_to_read):
    dataframes = {}

    for file_path in file_paths:
        file_data = {}
        for sheet_name in sheets_to_read[file_path]:
            if sheet_name.startswith('Site audit'):
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                df = modify_site_audit_df(df)
            else:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            file_data[sheet_name] = df
        dataframes[file_path] = file_data

    return dataframes

def modify_site_audit_df(df):
    # Get info
    transept_length = df.iloc[1:, 1]
    transept_area = df.iloc[1:, 2]
    weather = df.iloc[0, 6]  # Cell G1
    vegetation = df.iloc[1, 6]  # Cell G2
    date_of_audit = df.iloc[2, 6]  # Cell G3
    time_of_audit = df.iloc[3, 6]  # Cell G4
    last_audited = df.iloc[4, 6]  # Cell G5

    # Add data to all rows
    weather_list = [weather] * len(transept_length)
    vegetation_list = [vegetation] * len(transept_length)
    date_of_audit_list = [date_of_audit] * len(transept_length)
    time_of_audit_list = [time_of_audit] * len(transept_length)
    last_audited_list = [last_audited] * len(transept_length)

    # Column names
    df.columns = df.iloc[0]
    df = df[1:]

    # Create dataframe
    modified_df = pd.DataFrame({
        'Transept length': transept_length,
        'Transept Area': transept_area,
        'Weather': weather_list,
        'Vegetation': vegetation_list,
        'Date of audit': date_of_audit_list,
        'Time of Audit': time_of_audit_list,
        'Last audited': last_audited_list
    })

    # Drop unneccesary data
    modified_df = modified_df.drop(modified_df.index[5])
    modified_df = modified_df.iloc[:6]
    modified_df = modified_df.reset_index(drop=True)

    return modified_df

def save_to_excel(dataframes):
    for file_path, sheets in dataframes.items():
        for sheet_name, df in sheets.items():
            if sheet_name.startswith('Audit'):
                site_audit_sheet = sheet_name.replace('Audit', 'Site audit')
                if site_audit_sheet in sheets:
                    output_file = f"{sheet_name.replace(' ', '_')}.xlsx"
                    with pd.ExcelWriter(output_file) as writer:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        sheets[site_audit_sheet].to_excel(writer, sheet_name=site_audit_sheet, index=False)

file_paths = ['Litter Audit Data 5.xlsx', 'Litter Audit Data 1_2_3_4.xlsx', 'Litter Audit Data 6_7.xlsx']
sheets_to_read = {
    'Litter Audit Data 5.xlsx': ['Audit 5', 'Site audit 5'],
    'Litter Audit Data 1_2_3_4.xlsx': ['Audit 1', 'Site audit 1', 'Audit 2', 'Site audit 2', 'Audit 3', 'Site audit 3', 'Audit 4', 'Site audit 4'],
    'Litter Audit Data 6_7.xlsx': ['Audit 6', 'Site audit 6', 'Audit 7', 'Site audit 7']
}
 
dataframes = read_litter_audit_sheets(file_paths, sheets_to_read)

# Save to excel
save_to_excel(dataframes)
