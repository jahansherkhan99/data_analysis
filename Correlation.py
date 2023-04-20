import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

df = pd.read_csv('Data_Assessment.csv')

# Create Correlation Matrix
corr_matrix = df.corr()

# Display Correlation Matrix in form of Heatmap
sns.heatmap(corr_matrix, annot=True, cmap='coolwarm')
plt.title('Correlation Matrix Heatmap')
plt.show()