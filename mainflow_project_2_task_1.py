# import neccessary libraries
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import zscore

# Load the Excel dataset
file_path = r"C:\Users\sraja\Downloads\archive\Global_Superstore2.xlsx"
data = pd.read_excel(file_path, engine='openpyxl')

# Clean Data: Handle missing values (numeric columns only)
numeric_columns = data.select_dtypes(include=['number']).columns
data[numeric_columns] = data[numeric_columns].fillna(data[numeric_columns].mean())

# Remove duplicates
data = data.drop_duplicates()

# Detect and handle outliers using Z-scores (numeric columns only)
data = data[(np.abs(zscore(data[numeric_columns])) < 3).all(axis=1)]

# Perform Statistical Analysis
print("Descriptive Statistics:")
print(data.describe())

# Visualize Data
# Histogram for Sales
plt.hist(data['Sales'], bins=30, edgecolor='k')
plt.title("Sales Distribution")
plt.xlabel("Sales")
plt.ylabel("Frequency")
plt.show()

# Boxplot for Profit
plt.boxplot(data['Profit'])
plt.title("Profit Boxplot")
plt.ylabel("Profit")
plt.show()

# 4.3 Correlation Heatmap
# Select numeric columns and compute the correlation matrix
corr_matrix = data.select_dtypes(include=['number']).corr()

# Plot the heatmap
sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', fmt=".2f")
plt.title("Correlation Heatmap")
plt.show()

