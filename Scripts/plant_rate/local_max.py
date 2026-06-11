import pandas as pd
from scipy.signal import find_peaks

# Sample DataFrame
# Replace this with your actual DataFrame
data = {'time': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
        'value': [500, 1200, 800, 1500, 900, 1100, 1300, 950, 1050, 100]}

df = pd.DataFrame(data)

# Set values lower than 1000 to zero
df['value'] = df['value'].apply(lambda x: 0 if x < 1000 else x)

# Find local maxima
peaks, _ = find_peaks(df['value'])

# Create a new DataFrame with local maxima
maxima_df = df.loc[peaks]

# Print the results
print("Original DataFrame:")
print(df)

print("\nDataFrame with Local Maxima:")
print(maxima_df)
