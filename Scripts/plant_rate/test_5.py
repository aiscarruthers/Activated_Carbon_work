import pandas as pd

# Sample DataFrames
df1 = pd.DataFrame({'timestamp': ['2023-11-27 12:00:00', '2023-11-27 13:00:00', '2023-11-27 14:00:00', '2023-11-27 15:00:00'],
                    'value_df1': [1, 2, 3, 4]})

df2 = pd.DataFrame({'timestamp': ['2023-11-27 11:45:00', '2023-11-27 13:10:00', '2023-11-27 14:30:00'],
                    'value_df2': ['A', 'B', 'C']})

# Convert the 'timestamp' columns to datetime format
df1['timestamp'] = pd.to_datetime(df1['timestamp'])
df2['timestamp'] = pd.to_datetime(df2['timestamp'])

# Perform an asof merge
merged_df = pd.merge_asof(df2, df1, on='timestamp', direction='nearest')

# Display the merged DataFrame
print(merged_df)
