import pandas as pd
from sklearn.metrics.pairwise import cosine_similarity

# Step 1: Load the dataset
file_path = 'Dataset Ransomware.xlsx'
gang_profile_df = pd.read_excel(file_path, sheet_name='Ransomware Gang Profile')

# Step 2: Clean and Prepare the Data
# Extract the relevant columns
gangs = gang_profile_df['Gang name']
ttp_data = gang_profile_df.iloc[:, 7:]  # TTPs start from the 7th column onward

# Replace NaN values with 0s
ttp_data_cleaned = ttp_data.fillna(0)

# Create a binary matrix where rows are gangs and columns are TTPs
binary_matrix_cleaned = pd.get_dummies(ttp_data_cleaned.stack()).groupby(level=0).sum()

# Step 3: Calculate Cosine Similarity
# Compute the cosine similarity between gangs
similarity_matrix_cleaned = pd.DataFrame(cosine_similarity(binary_matrix_cleaned),
                                         index=gangs, columns=gangs)

# Step 4: Display the Similarity Matrix in the Output
pd.set_option('display.max_rows', None)  # Show all rows
pd.set_option('display.max_columns', None)  # Show all columns
print(similarity_matrix_cleaned)
