from sklearn.model_selection import train_test_split
from sklearn.preprocessing import LabelEncoder
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import classification_report, confusion_matrix, accuracy_score

# Pulizia dei dati: Riempire i valori mancanti
df['gang'].fillna('Unknown', inplace=True)

# Codifica delle variabili categoriali
label_encoders = {}
for column in ['victim', 'gang', 'Victim Country', 'Victim sectors']:
    le = LabelEncoder()
    df[column] = le.fit_transform(df[column])
    label_encoders[column] = le

# Selezione delle caratteristiche e della variabile target
X = df[['victim', 'Victim Country', 'Victim sectors']]
y = df['gang']

# Suddivisione del dataset
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, random_state=42)

# Addestramento del modello Random Forest
rf_model = RandomForestClassifier(n_estimators=100, random_state=42)
rf_model.fit(X_train, y_train)

# Previsione sui dati di test
y_pred = rf_model.predict(X_test)

# Valutazione del modello
print("Accuracy:", accuracy_score(y_test, y_pred))
print("Confusion Matrix:\n", confusion_matrix(y_test, y_pred))
print("Classification Report:\n", classification_report(y_test, y_pred))
