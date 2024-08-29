import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Definire la funzione per determinare il range di dipendenti
def employee_range(number):
    if pd.isna(number):
        return "N/A"
    if number <= 0:
        return "N/A"
    if number <= 5:
        return '1-5'
    elif number <= 10:
        return '6-10'
    elif number <= 15:
        return '11-15'
    elif number <= 20:
        return '16-20'
    elif number <= 30:
        return '21-30'
    elif number <= 40:
        return '31-40'
    elif number <= 50:
        return '41-50'
    elif number <= 75:
        return '51-75'
    elif number <= 100:
        return '76-100'
    elif number <= 150:
        return '101-150'
    elif number <= 200:
        return '151-200'
    elif number <= 300:
        return '201-300'
    elif number <= 400:
        return '301-400'
    elif number <= 500:
        return '401-500'
    elif number <= 750:
        return '501-750'
    elif number <= 1000:
        return '751-1000'
    elif number <= 2000:
        return '1001-2000'
    elif number <= 3000:
        return '2001-3000'
    elif number <= 4000:
        return '3001-4000'
    elif number <= 5000:
        return '4001-5000'
    elif number <= 7500:
        return '5001-7500'
    elif number <= 10000:
        return '7501-10000'
    elif number <= 15000:
        return '10001-15000'
    elif number <= 20000:
        return '15001-20000'
    else:
        return '20001+'

def carica_e_modifica_excel():
    root = tk.Tk()
    root.withdraw()

    # Leggere il file Excel
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        messagebox.showerror("Errore", "Nessun file selezionato")
        return

    try:
        df = pd.read_excel(file_path)

        # Controllare se la colonna 'Number of employees' esiste
        if 'Number of employees' not in df.columns:
            messagebox.showerror("Errore", "La colonna 'Number of employees' non esiste nel file selezionato")
            return

        # Creare una nuova colonna con il range di dipendenti
        df['Number of employees'] = df['Number of employees'].apply(employee_range)

        # Creare un nuovo dataframe con tutte le colonne originali più la nuova colonna
        df_modified = df.copy()
        df_modified['Number of employees'] = df['Number of employees']

        # Chiedere all'utente dove salvare il file Excel modificato
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path:
            messagebox.showerror("Errore", "Nessun percorso di salvataggio selezionato")
            return

        # Salvare il file Excel modificato
        df_modified.to_excel(save_path, index=False)
        messagebox.showinfo("Successo", "File Excel aggiornato con successo!")

    except Exception as e:
        messagebox.showerror("Errore", f"Si è verificato un errore durante l'elaborazione del file: {e}")

if __name__ == "__main__":
    carica_e_modifica_excel()
