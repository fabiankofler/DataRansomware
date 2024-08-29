import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error
from sklearn.preprocessing import LabelEncoder

# Step 1: Load the dataset
file_path = 'Dataset Ransomware.xlsx'
dataset_df = pd.read_excel(file_path, sheet_name='Dataset')

# Convert 'date' column to datetime
dataset_df['date'] = pd.to_datetime(dataset_df['date'], errors='coerce')

# Step 2: Prepare the dataset for regression
# Group by sector and year to calculate the number of attacks per sector per year
sector_attack_trend = dataset_df.groupby(['Victim sectors', dataset_df['date'].dt.year]).size().unstack(fill_value=0)

# Reshape the data for modeling (each sector-year pair as a sample)
data = sector_attack_trend.stack().reset_index().rename(columns={0: 'attack_count'})

# Encode the 'Victim sectors' using LabelEncoder
le_sector = LabelEncoder()
data['Victim sectors'] = le_sector.fit_transform(data['Victim sectors'])

# Step 3: Feature and Target Selection
X = data[['Victim sectors', 'date']]  # Features: Sector and Year
y = data['attack_count']  # Target: Number of attacks

# Step 4: Split data into training and testing sets
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Step 5: Train a Random Forest Regressor
regressor = RandomForestRegressor(random_state=42)
regressor.fit(X_train, y_train)

# Step 6: Make predictions and evaluate the model
y_pred = regressor.predict(X_test)
rmse = np.sqrt(mean_squared_error(y_test, y_pred))
print(f"Model RMSE: {rmse}")

# Step 7: Predict future attacks for the year 2025
X_future = X_test.copy()
X_future['date'] = 2025
y_future_pred = regressor.predict(X_future)

# Decode the sector names from the LabelEncoder
X_future['Victim sectors'] = le_sector.inverse_transform(X_future['Victim sectors'])

# Add predictions to the DataFrame
X_future['predicted_attacks'] = y_future_pred

# Aggregate predictions by sector (e.g., by averaging)
aggregated_predictions = X_future.groupby('Victim sectors')['predicted_attacks'].mean().reset_index()

# Sort aggregated predictions by the predicted number of attacks
aggregated_predictions_sorted = aggregated_predictions.sort_values(by='predicted_attacks', ascending=False)

# Display the top sectors predicted for 2025
top_predicted_sectors = aggregated_predictions_sorted.head(10)
print(top_predicted_sectors)

