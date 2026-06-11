
import pandas as pd
import matplotlib.pyplot as plt

# Load Excel file
file_path = "C:\\Users\\ACarruther\\OneDrive - NORIT AMERICAS INC\\Documents\\CNR150 analysis\\CNR150 Trial August 25 analysis_v3.xlsm"  # Replace with your actual file path
sheet_name = 'Kiln v Quality'  # Replace with your sheet name if different
good_pallet_sheet = 'Sheet2'
# Read the data
df = pd.read_excel(file_path, sheet_name=sheet_name)
df_gp = pd.read_excel(file_path, sheet_name=good_pallet_sheet)

# Display the first few rows to understand the structure
# print(df.head())

# Show available columns
# print("Columns in the DataFrame:", df.columns.tolist())

# Plot each numeric column as both scatter and line plot
# numeric_cols = df.select_dtypes(include='number').columns

phases = df['Phase'].unique()



params = {
    'Acid':['F11 Fresh Acid', 'F28 Acid Bleed', 'Total Dry Acid:Raw material'],
    'A Kiln':['A kiln Gas per mix (Nm3/mix)', 'A kiln Back box', 'A kiln Temperature'],
    'C kiln':['C kiln Gas per mix (Nm3/mix)', 'C kiln Back box', 'C kiln Temperature']
}

qualities = ['BWC', 'Weight Ads (%)', 'Density', 'Volatiles', 'wt retained (%)', 'Pressure Drop', '>2.36mm']

limits = {
    'BWC':14.8,
    'Density':340,
    'Weight Ads (%)':50,
    'Volatiles': 15
}

for phase in phases:
    phase_df = df[df['Phase']== phase]
    for col in phase_df.select_dtypes(include='number').columns:
        fig, ax1 = plt.subplots()
        ax1.plot(phase_df['Time in C kiln'], phase_df[col], label=col)
        ax1.set_xlabel('Time in C kiln')
        ax1.set_ylabel(col)
        plt.title(phase+" "+col)
        ax2= ax1.twinx()
        ax2.scatter(phase_df['Time in C kiln'], phase_df['BWC'])
        ax2.plot(phase_df['Time in C kiln'])
        ax2.set_ylabel('BWC')
        plt.tight_layout()
        plt.show()

# for col in numeric_cols:
#     plt.figure(figsize=(10, 5))
#     # Scatter plot
#     plt.subplot(1, 2, 1)
#     plt.scatter(df['Time in C kiln'], df[col], color='blue')
#     plt.title(f'Scatter Plot of {col}')
#     plt.xlabel('Time in C kiln')
#     plt.ylabel(col)
#     # Line plot
#     plt.subplot(1, 2, 2)
#     plt.plot(df['Time in C kiln'], df[col], color='green', marker='o')
#     plt.title(f'Line Plot of {col}')
#     plt.xlabel('Time in C kiln')
#     plt.ylabel(col)
#     plt.tight_layout()
#     plt.show()