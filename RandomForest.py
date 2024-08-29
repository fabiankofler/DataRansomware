#Application in the Code:
#Training: The RandomForestRegressor model is trained on the historical ransomware data, using features like country and year.
#Prediction: It predicts the number of ransomware attacks for each country in the year 2025.
#This model is well-suited for the task due to its ability to handle non-linear data and provide robust predictions based on multiple decision trees.

import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error
from sklearn.preprocessing import LabelEncoder
import numpy as np

# Step 1: Load the dataset
file_path = 'Dataset Ransomware.xlsx'
dataset_df = pd.read_excel(file_path, sheet_name='Dataset')

# Convert 'date' column to datetime
dataset_df['date'] = pd.to_datetime(dataset_df['date'], errors='coerce')

# Step 2: Analyze the historical attack data by country
country_attack_trend = dataset_df.groupby(['Victim Country', dataset_df['date'].dt.year]).size().unstack(fill_value=0)

# Step 3: Prepare the dataset for modeling
# Reshape the data for modeling (each country-year pair as a sample)
data = country_attack_trend.stack().reset_index().rename(columns={0: 'attack_count'})

# Encode the 'Victim Country' using LabelEncoder
le_country = LabelEncoder()
data['Victim Country'] = le_country.fit_transform(data['Victim Country'])

# Step 4: Feature and Target Selection
X = data[['Victim Country', 'date']]  # Features: Country and Year
y = data['attack_count']  # Target: Number of attacks

# Step 5: Split data into training and testing sets
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Step 6: Train a Random Forest Regressor
regressor = RandomForestRegressor(random_state=42)
regressor.fit(X_train, y_train)

# Step 7: Make predictions and evaluate the model
y_pred = regressor.predict(X_test)
rmse = np.sqrt(mean_squared_error(y_test, y_pred))
print(f"Model RMSE: {rmse}")

# Step 8: Predict future attacks for the year 2025
X_future = X_test.copy()
X_future['date'] = 2025
y_future_pred = regressor.predict(X_future)

# Decode the country names from the LabelEncoder
X_future['Victim Country'] = le_country.inverse_transform(X_future['Victim Country'])

# Add predictions to the DataFrame
X_future['predicted_attacks'] = y_future_pred

# Aggregate predictions by country (e.g., by averaging)
aggregated_predictions = X_future.groupby('Victim Country')['predicted_attacks'].mean().reset_index()

# Sort aggregated predictions by the predicted number of attacks
aggregated_predictions_sorted = aggregated_predictions.sort_values(by='predicted_attacks', ascending=False)

# Display the top 10 countries predicted for 2025
top_predicted_countries = aggregated_predictions_sorted.head(10)
print(top_predicted_countries)

