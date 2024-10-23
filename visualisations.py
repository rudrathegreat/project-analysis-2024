import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.gridspec import GridSpec
import numpy as np

# Plot stylesheet
# plt.style.use('seaborn-v0_8-whitegrid')

# List of Excel files
excel_files = ['Audit_1.xlsx', 'Audit_2.xlsx', 'Audit_3.xlsx', 'Audit_4.xlsx', 'Audit_5.xlsx', 'Audit_6.xlsx', 'Audit_7.xlsx']

# Category colours
colors = {
    'EPS': '#1f77b4',  # blue
    'ME': '#ff7f0e',  # orange
    'PL': '#2ca02c',  # green
    'OT': '#d62728',  # red
}

category_order = ['EPS', 'ME', 'OT', 'PL']

# Category names
category_full_names = {
    'EPS': 'Expanded polystyrene',
    'ME': 'Metal',
    'OT': 'Other',
    'PL': 'Plastic'
}

audit_categories_data = {}

rainfall_df = []

for key in category_full_names.keys():
    audit_categories_data[key] = []

for file in excel_files:
    fig = plt.figure(figsize=(16, 10))
    gs = GridSpec(2, 3, figure=fig, width_ratios=[1, 1, 0.5], height_ratios=[1, 1])

    # Subplots
    ax1 = fig.add_subplot(gs[0, 0])
    ax2 = fig.add_subplot(gs[0, 1])
    ax3 = fig.add_subplot(gs[1, 0])
    ax4 = fig.add_subplot(gs[1, 1])
    ax_legend = fig.add_subplot(gs[:, 2])
    
    ax_legend.axis('off')

    #Read file
    xl = pd.ExcelFile(file)
    sheet_name_template = file.split('.')[0]
    sheet_name = sheet_name_template.replace('_', ' ')
    site_audit_sheet = f"Site audit {sheet_name.split(' ')[1]}"
    
    #create dataframes
    df = xl.parse(sheet_name)
    audit_df = xl.parse(site_audit_sheet)

    # Formatting
    audit_df.columns = audit_df.columns.str.replace('Transept length', 'Transept length (m)')
    audit_df.columns = audit_df.columns.str.replace('Transept Area', 'Transept area (m$^2$)')
    audit_df['Date of audit'] = pd.to_datetime(audit_df['Date of audit'], format='%Y-%m-%d', errors='coerce').dt.date
    audit_df['Last audited'] = pd.to_datetime(audit_df['Last audited'], format='%Y-%m-%d', errors='coerce').dt.date
    audit_df['Date of audit'] = audit_df['Date of audit'].fillna('unknown')
    audit_df['Last audited'] = audit_df['Last audited'].fillna('unknown')
    df['Volume (L)'] = pd.to_numeric(df['Volume (ml)'], errors='coerce') / 1000
    df['Weight (kg)'] = pd.to_numeric(df['Weight (g)'], errors='coerce') / 1000
    
    # Category order
    df['Categories'] = pd.Categorical(df['Categories'], categories=category_order, ordered=True)
    
    # Category colours pt2
    df['Color'] = df['Categories'].map(colors)

    # No. items bar chart
    df_grouped = df.groupby('Categories', observed=True)['No.of Items'].sum().reindex(category_order).reset_index()
    ax1.bar(df_grouped['Categories'], df_grouped['No.of Items'], alpha=0.8, color=[colors.get(cat, '#7f7f7f') for cat in df_grouped['Categories']])
    ax1.set_title('Categories vs No. of Items')
    ax1.set_xlabel('Categories')
    ax1.set_ylabel('No. of Items')
    ax1.set_ylim(0, 1000)
    ax1.grid(axis='y')

    for i,category in enumerate(df_grouped['Categories']):
        audit_categories_data[category].append(df_grouped['No.of Items'][i])
    
    # Volume bar chart
    df_grouped_volume = df.groupby('Categories', observed=True)['Volume (L)'].sum().reindex(category_order).reset_index()
    ax2.bar(df_grouped_volume['Categories'], df_grouped_volume['Volume (L)'], alpha=0.8, color=[colors.get(cat, '#7f7f7f') for cat in df_grouped_volume['Categories']])
    ax2.set_title('Categories vs Volume (L)')
    ax2.set_xlabel('Categories')
    ax2.set_ylabel('Volume (L)')
    ax2.set_ylim(0, 25)
    ax2.grid(axis='y')
    
    # Weight bar chart
    df_grouped_weight = df.groupby('Categories', observed=True)['Weight (kg)'].sum().reindex(category_order).reset_index()
    ax3.bar(df_grouped_weight['Categories'], df_grouped_weight['Weight (kg)'], alpha=0.8, color=[colors.get(cat, '#7f7f7f') for cat in df_grouped_weight['Categories']])
    ax3.set_title('Categories vs Weight (kg)')
    ax3.set_xlabel('Categories')
    ax3.set_ylabel('Weight (kg)')
    ax3.set_ylim(0, 10)
    ax3.grid(axis='y')
    
    # Rainfall data
    rain_df = xl.parse('flemingtonMelbW_modified.csv')
    
    # Rainfall date formatting
    rain_df['Date'] = pd.to_datetime(rain_df['Date'])
    
    # Cumulative sum
    rain_df['CumulativeRainfall'] = rain_df[' 2024'].cumsum()
    
    # Format x-ticks to remove the year
    rain_df['DateFormatted'] = rain_df['Date'].dt.strftime('%b %d')
    
    # Rainfall line graph
    ax4.plot(rain_df['Date'], rain_df['CumulativeRainfall'])

    rainfall_df.append(np.array(rain_df['CumulativeRainfall'])[-1])

    ax4.set_title('Cumulative Rainfall in 2024')
    ax4.set_xlabel('Date')
    ax4.set_ylabel('Cumulative Rainfall (mm)')
    ax4.set_ylim(0, 20)
    ax4.grid(axis='y')
    
    # Rename x-ticks for audit dates
    if file != 'Audit_1.xlsx':
        xticks_labels = ['Last audit'] + rain_df['DateFormatted'][1:-1].tolist() + ['Audit date']
    else:
        xticks_labels = rain_df['DateFormatted'].tolist()[:-1] + ['Audit date']

    ax4.set_xticks(rain_df['Date'])
    ax4.set_xticklabels(xticks_labels, rotation=45)

    # Title
    file_label = file.replace('.xlsx', '').replace('_', ' ')
    ax_legend.set_title(file_label, fontsize=14)

    # Legend
    handles = [plt.Rectangle((0,0),1,1, color=colors[key]) for key in category_order]
    labels = [f"{key} = {category_full_names[key]}" for key in category_order]
    # ax_legend.legend(handles, labels, loc='upper center', frameon=False)
    ax_legend.legend(handles, labels, frameon=True)

    #Transept table
    # transect_data = audit_df[['Transept length (m)', 'Transept area (m$^2$)']].head(5).to_numpy()
    # transect_table = pd.DataFrame(transect_data, columns=['Transept length (m)', 'Transept area (m$^2$)'], index=range(1, 6))
    # ax_legend.table(cellText=transect_table.values, colLabels=transect_table.columns, rowLabels=transect_table.index, loc='center', cellLoc='center', fontsize=14)

    # # Audit info table
    # audit_info_data = audit_df.iloc[0][['Weather', 'Vegetation', 'Date of audit', 'Time of Audit', 'Last audited']].to_frame().transpose()
    # audit_info_table = pd.DataFrame(audit_info_data.values).transpose()
    # audit_info_table.columns = ['Audit Info']
    # audit_info_table.index = ['Weather', 'Vegetation', 'Date of audit', 'Time of Audit', 'Last audited']
    # ax_legend.table(cellText=audit_info_table.values, colLabels=['Audit Info'], rowLabels=audit_info_table.index, loc='center', cellLoc='center', bbox=[0.0, 0, 0, 0])

    # Save images
    plt.tight_layout()
    plt.savefig(f'visualizations/{sheet_name_template}_visualization.png')
    plt.close(fig)

print("Visualizations saved to the 'visualizations' folder.")

plt.clf()

print(audit_categories_data.keys())

total_volume = np.zeros(len(audit_categories_data[category]))

for category in audit_categories_data.keys():
    print(category)
    audit_data = np.array(audit_categories_data[category])
    difference_between_audits = np.diff(audit_data)

    total_volume += audit_data

    p = plt.plot(rainfall_df[2:], difference_between_audits[1:], '.', label=category, markersize=9)

    color = p[0].get_color()

    coefficients = np.polyfit(rainfall_df[2:], difference_between_audits[1:], 1)
    slope, intercept = coefficients

    plt.plot(rainfall_df[2:], ((slope * np.array(rainfall_df[2:])) + intercept), color=color, label='{} fit'.format(category), alpha=0.3)

difference_between_audits = np.diff(total_volume)
plt.plot(rainfall_df[2:], difference_between_audits[1:], '.', label='Total', color='black', markersize=9)
coefficients = np.polyfit(rainfall_df[2:], difference_between_audits[1:], 1)
slope, intercept = coefficients
plt.plot(rainfall_df[2:], ((slope * np.array(rainfall_df[2:])) + intercept), color='black', label='{} fit'.format('Total'), alpha=0.3)

plt.xlabel('Cumulative Rainfall Between Previous Audit and Current Audit')
plt.ylabel('Difference in Volume of Litter')
plt.title('Correlation Between Amount of Rainfall and Accumulation of Plastics')

plt.legend()
plt.show()

total = 0
for category in audit_categories_data.keys():
    audit_data = np.array(audit_categories_data[category])
    all_litter_items = np.sum(audit_data)
    total += all_litter_items
    print(category, all_litter_items)
print('Total', total)