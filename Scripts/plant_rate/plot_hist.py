import matplotlib.pyplot as plt
import matplotlib.mlab as mlab
import numpy as np
from scipy.stats import norm
import pandas as pd

def plot_histogram(csv_file, column_name, sp, usl, lsl):
    # Load the CSV file into a pandas DataFrame
    df = pd.read_csv(csv_file)

    # Extract the desired column data
    column_data = df[column_name]

    # Calculate average, UCL, and LCL
    average = column_data.mean()
    std_dev = column_data.std()
    ucl = average + 3 * std_dev
    lcl = average - 3 * std_dev

    # Plot the histogram
    plt.hist(column_data, bins=2)  # You can adjust the number of bins as needed
    
    # Fit a Gaussian curve to the data
    
    mu, sigma = norm.fit(column_data)
    x = np.linspace(lcl - std_dev, ucl + std_dev, 100)
    y = norm.pdf(x, mu, sigma)
    count = len(column_data)
    y_scaled = y * count
    # Plot the Gaussian curve
    plt.plot(x, y_scaled, 'r-', linewidth=2, label='Gaussian Fit')

    plt.axvline(average, color='r', linestyle='solid', linewidth=1, label='Average = {}'.format(str(round(average, 4))))
    plt.axvline(ucl, color='r', linestyle='dashed', linewidth=1, label='UCL = {}'.format(str(round(ucl, 4))))
    plt.axvline(lcl, color='r', linestyle='dashed', linewidth=1, label='LCL = {}'.format(str(round(lcl, 4))))
    
    
    sigma_spec = (sp - lsl)/3

    # Plot the Gaussian curve representing the specification limits
    x_spec = np.linspace(lsl, usl, 100)
    y_spec = norm.pdf(x_spec, sp, sigma_spec)
    y_spec_scaled = y_spec * count
    
    # plt.plot(x_spec, y_spec_scaled, 'y--', linewidth=2, label='Control plan Limits')

    # Configure plot labels and title

    plt.xlabel(column_name)
    plt.ylabel('Frequency')
    # plt.title('Histogram of ' + column_name)
    
    # Add vertical lines for average, UCL, and LCL
    plt.axvline(sp, color='y', linestyle='solid', label='Target = {}'.format(str(sp)) )
    # plt.axvline(lsl, color='y', linestyle='dashed', linewidth=1, label='LSL = {}'.format(str(round(lsl, 4))))
    # plt.axvline(usl, color='y', linestyle='dashed', linewidth=1, label='USL = {}'.format(str(round(usl, 4))))

    # Display the legend
    plt.legend()
    # Display the histogram
    plt.show()

# csv_file_path = 'OSF_hist.csv'  # Replace with the actual path to your CSV file
# column_name_to_plot = 'OSF'  # Replace with the column name you want to plot
# plot_histogram("OSF_hist.csv", "OSF", 505, 510, 490)
# plot_histogram("OSF_hist_1.csv", "OSF", 505, 510, 490)
plot_histogram("acid_hist.csv", "Acid", 467, 450, 490)
plot_histogram("acid_hist_1.csv", "Acid", 467, 450, 490)
# plot_histogram("D01_hist.csv", "D01", 1540, 1550, 1530)
# plot_histogram("D01_hist_1.csv", "D01", 1540, 1550, 1530)



