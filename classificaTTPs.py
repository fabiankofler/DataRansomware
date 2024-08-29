import pandas as pd
from itertools import combinations

# Caricamento del file Excel
file_path = 'Dataset Ransomware.xlsx'
ransomware_gang_profile_df = pd.read_excel(file_path, sheet_name='Ransomware Gang Profile')
ransomware_ttp_df = pd.read_excel(file_path, sheet_name='RANSOMWARE TTP')

# Step 1: Classifica delle TTP più utilizzate dal foglio "Ransomware TTP"
ttp_data = ransomware_ttp_df.iloc[1:, [1, 7]].copy()  # Colonna 1 contiene il nome della gang, colonna 7 contiene i TTP
ttp_data.columns = ['Gang name', 'TTPS']

# Separare i TTP in liste e rimuovere i valori mancanti
ttp_data['TTPS'] = ttp_data['TTPS'].str.split(',').apply(
    lambda x: [item.strip() for item in x] if isinstance(x, list) else [])

# Contare le occorrenze di ogni TTP
all_ttps = [ttp for sublist in ttp_data['TTPS'] for ttp in sublist if ttp]  # Crea una lista con tutte le TTP
ttp_counts = pd.Series(all_ttps).value_counts().reset_index()
ttp_counts.columns = ['TTP', 'Count']

# Identificare le gang che usano ogni TTP
gangs_by_ttp = ttp_data.explode('TTPS').groupby('TTPS')['Gang name'].apply(list).reset_index()

# Unire i conteggi delle TTP con le gang che le usano
ttp_summary = pd.merge(ttp_counts, gangs_by_ttp, left_on='TTP', right_on='TTPS').drop('TTPS', axis=1)

# Formattare l'output della classifica delle TTP
ttp_rankings = [
    f"{index + 1}. {row['TTP']}: Usata da {row['Count']} Gang: {', '.join(row['Gang name'])}"
    for index, row in ttp_summary.iterrows()
]

# Step 2: Classifica della similarità tra gang di ransomware
ttp_similarity_data = ransomware_ttp_df.iloc[1:, :].copy()  # Salta la prima riga che è l'intestazione
ttp_similarity_data.columns = ransomware_ttp_df.iloc[0]  # Imposta le intestazioni corrette

gang_ttp_grouped = ttp_similarity_data.groupby("Gang name")["TTPS"].apply(lambda x: set(x.dropna())).reset_index()

similarity_list = []
for (i, row1), (j, row2) in combinations(gang_ttp_grouped.iterrows(), 2):
    common_ttps = row1["TTPS"].intersection(row2["TTPS"])
    total_ttps = row1["TTPS"].union(row2["TTPS"])
    if len(total_ttps) > 0:  # Evita divisioni per zero
        similarity_score = len(common_ttps) / len(total_ttps) * 100
        if similarity_score > 0:  # Considera solo similarità non nulle
            similarity_list.append(
                (row1["Gang name"], row2["Gang name"], similarity_score, common_ttps)
            )

similarity_list.sort(key=lambda x: x[2], reverse=True)

similarity_rankings = [
    f"{gang1} = {gang2} ({round(score, 2)}%): {', '.join(sorted(map(str, ttps)))}"
    for gang1, gang2, score, ttps in similarity_list
]

# Salvataggio di entrambi i risultati in un singolo file CSV
with open('ransomware_analysis.csv', 'w') as file:
    file.write("Classifica delle TTP più utilizzate dalle gang ransomware:\n")
    for line in ttp_rankings[:157]:
        file.write(line + "\n")

    file.write("\nSimilarità tra gang ransomware basate sulle TTP:\n")
    for line in similarity_rankings[:21609]:
        file.write(line + "\n")

# Visualizza i risultati nella console
print("Classifica delle TTP più utilizzate dalle gang ransomware:")
# il Numero dentro ttp_rankings = il numero di TTP nel Foglio ID TTP
for line in ttp_rankings[:157]:
    print(line)

print("\nSimilarità tra gang ransomware basate sulle TTP:")
# il Numero dentro similarity_rankings = il numero di gang x il numero di gang. Esempio: 147 gang in totale, quindi: 147*147= 21609.
for line in similarity_rankings[:21609]:
    print(line)

# Indica che il file CSV è stato creato
"File CSV 'ransomware_analysis.csv' creato con successo."
