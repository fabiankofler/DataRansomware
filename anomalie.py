import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class RansomwareAnalyzer:
    def __init__(self, master):
        self.master = master
        master.title("Ransomware Analyzer")

        self.label = ttk.Label(master, text="Analizza Anomalie negli Attacchi Ransomware")
        self.label.pack(pady=10)

        self.upload_button = ttk.Button(master, text="Carica Dataset", command=self.load_dataset)
        self.upload_button.pack(pady=10)

        self.analyze_button = ttk.Button(master, text="Analizza", command=self.analyze, state=tk.DISABLED)
        self.analyze_button.pack(pady=10)

        self.export_button = ttk.Button(master, text="Esporta Risultati", command=self.export_results, state=tk.DISABLED)
        self.export_button.pack(pady=10)

        self.result_text = tk.Text(master, height=10, width=50)
        self.result_text.pack(pady=10)

        # Gestire l'evento di chiusura della finestra
        self.master.protocol("WM_DELETE_WINDOW", self.on_exit)

    def load_dataset(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            try:
                self.df = pd.read_excel(file_path, sheet_name='Dataset')
                messagebox.showinfo("Caricamento", "Dataset caricato con successo!")
                self.analyze_button.config(state=tk.NORMAL)
                self.export_button.config(state=tk.NORMAL)
            except Exception as e:
                messagebox.showerror("Errore", f"Errore nel caricamento del file: {str(e)}")

    def analyze(self):
        self.result_text.delete(1.0, tk.END)

        self.df['date'] = pd.to_datetime(self.df['date'])
        self.time_series = self.df['date'].dt.date.value_counts().sort_index()

        self.analyze_temporal_anomalies()
        self.analyze_country_anomalies()
        self.analyze_sector_anomalies()

        self.plot_time_series()

    def analyze_temporal_anomalies(self):
        mean_attacks = self.time_series.mean()
        std_attacks = self.time_series.std()
        threshold = mean_attacks + 2 * std_attacks
        anomalous_days = self.time_series[self.time_series > threshold]

        self.result_text.insert(tk.END, "Picchi Temporali Inusuali:\n")
        self.result_text.insert(tk.END, anomalous_days.to_string())
        self.result_text.insert(tk.END, "\n\n")

    def analyze_country_anomalies(self):
        country_distribution = self.df['Victim Country'].value_counts()
        mean_country = country_distribution.mean()
        std_country = country_distribution.std()
        threshold_country = mean_country + 2 * std_country
        anomalous_countries = country_distribution[country_distribution > threshold_country]

        self.result_text.insert(tk.END, "Paesi con Attacchi Inusuali:\n")
        self.result_text.insert(tk.END, anomalous_countries.to_string())
        self.result_text.insert(tk.END, "\n\n")

    def analyze_sector_anomalies(self):
        sector_distribution = self.df['Victim sectors'].value_counts()
        mean_sector = sector_distribution.mean()
        std_sector = sector_distribution.std()
        threshold_sector = mean_sector + 2 * std_sector
        anomalous_sectors = sector_distribution[sector_distribution > threshold_sector]

        self.result_text.insert(tk.END, "Settori con Attacchi Inusuali:\n")
        self.result_text.insert(tk.END, anomalous_sectors.to_string())

    def plot_time_series(self):
        fig, ax = plt.subplots(figsize=(14, 7))
        ax.plot(self.time_series.index, self.time_series.values, marker='o')
        ax.set_title('Numero di Attacchi Ransomware nel Tempo')
        ax.set_xlabel('Data')
        ax.set_ylabel('Numero di Attacchi')
        ax.grid(True)
        plt.xticks(rotation=45)

        canvas = FigureCanvasTkAgg(fig, master=self.master)
        canvas.draw()
        canvas.get_tk_widget().pack()

    def export_results(self):
        export_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                   filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if export_path:
            try:
                with open(export_path, 'w') as f:
                    f.write(self.result_text.get(1.0, tk.END))
                messagebox.showinfo("Esportazione", "Risultati esportati con successo!")
            except Exception as e:
                messagebox.showerror("Errore", f"Errore durante l'esportazione: {str(e)}")

    def on_exit(self):
        # Puoi aggiungere qui qualsiasi pulizia necessaria prima di chiudere
        self.master.quit()
        self.master.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    analyzer = RansomwareAnalyzer(root)
    root.mainloop()
