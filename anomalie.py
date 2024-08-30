import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from fpdf import FPDF
import logging

# Set up logging
logging.basicConfig(filename='app.log', level=logging.INFO)

class RansomwareAnalyzer:
    def __init__(self, master):
        self.master = master
        master.title("Ransomware Analyzer")
        master.configure(bg='#282c34')

        self.setup_ui()

    def setup_ui(self):
        # Label for the application
        self.label = ttk.Label(self.master, text="Analyze Ransomware Attack Anomalies", foreground='#ffffff', background='#282c34')
        self.label.pack(pady=10)

        # Upload dataset button
        self.upload_button = ttk.Button(self.master, text="Load Dataset", command=self.load_dataset, style="TButton")
        self.upload_button.pack(pady=10)

        # Analyze data button
        self.analyze_button = ttk.Button(self.master, text="Analyze", command=self.analyze, state=tk.DISABLED, style="TButton")
        self.analyze_button.pack(pady=10)

        # Export results button
        self.export_button = ttk.Button(self.master, text="Export Results", command=self.export_results, state=tk.DISABLED, style="TButton")
        self.export_button.pack(pady=10)

        # Dropdown for theme selection
        self.theme_var = tk.StringVar(value='Dark')
        self.theme_menu = ttk.OptionMenu(self.master, self.theme_var, 'Dark', 'Dark', 'Light', command=self.change_theme)
        self.theme_menu.pack(pady=10)

        # Text area for displaying results
        self.result_text = tk.Text(self.master, height=10, width=50, bg='#1e2127', fg='#abb2bf', insertbackground='#ffffff')
        self.result_text.pack(pady=10)

        # Handle window close event
        self.master.protocol("WM_DELETE_WINDOW", self.on_exit)

        # Configure button styles
        style = ttk.Style()
        style.configure("TButton", background='#61afef', foreground='#282c34', font=('Helvetica', 10, 'bold'))

    def load_dataset(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            try:
                self.df = pd.read_excel(file_path, sheet_name='Dataset')
                required_columns = ['date', 'Victim Country', 'Victim sectors']
                for col in required_columns:
                    if col not in self.df.columns:
                        raise ValueError(f"Missing required column: {col}")
                messagebox.showinfo("Load", "Dataset loaded successfully!")
                self.analyze_button.config(state=tk.NORMAL)
                self.export_button.config(state=tk.NORMAL)
                logging.info("Dataset loaded successfully.")
            except Exception as e:
                logging.error(f"Error loading file: {e}")
                messagebox.showerror("Error", f"Error loading file: {str(e)}")

    def analyze(self):
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "Processing...\n")
        self.master.update_idletasks()  # Refresh the GUI

        try:
            self.df['date'] = pd.to_datetime(self.df['date'])
            self.time_series = self.df['date'].dt.date.value_counts().sort_index()

            temporal_results = self.analyze_temporal_anomalies()
            country_results = self.analyze_country_anomalies()
            sector_results = self.analyze_sector_anomalies()

            self.display_results(temporal_results, country_results, sector_results)
            self.plot_time_series()

            logging.info("Analysis completed successfully.")
        except Exception as e:
            logging.error(f"Error during analysis: {e}")
            messagebox.showerror("Error", f"Error during analysis: {str(e)}")

    def analyze_temporal_anomalies(self):
        mean_attacks = self.time_series.mean()
        std_attacks = self.time_series.std()
        threshold = mean_attacks + 2 * std_attacks
        anomalous_days = self.time_series[self.time_series > threshold]
        return f"Unusual Temporal Peaks:\n{anomalous_days.to_string()}"

    def analyze_country_anomalies(self):
        country_distribution = self.df['Victim Country'].value_counts()
        mean_country = country_distribution.mean()
        std_country = country_distribution.std()
        threshold_country = mean_country + 2 * std_country
        anomalous_countries = country_distribution[country_distribution > threshold_country]
        return f"Countries with Unusual Attacks:\n{anomalous_countries.to_string()}"

    def analyze_sector_anomalies(self):
        sector_distribution = self.df['Victim sectors'].value_counts()
        mean_sector = sector_distribution.mean()
        std_sector = sector_distribution.std()
        threshold_sector = mean_sector + 2 * std_sector
        anomalous_sectors = sector_distribution[sector_distribution > threshold_sector]
        return f"Sectors with Unusual Attacks:\n{anomalous_sectors.to_string()}"

    def display_results(self, *results):
        self.result_text.delete(1.0, tk.END)
        for result in results:
            self.result_text.insert(tk.END, result)
            self.result_text.insert(tk.END, "\n\n")
        self.result_text.insert(tk.END, "Analysis Complete\n")

    def plot_time_series(self):
        # Close any existing figures to prevent memory issues
        plt.close('all')

        # Create a new figure and plot
        fig, ax = plt.subplots(figsize=(14, 7))
        ax.plot(self.time_series.index, self.time_series.values, marker='o', color='#61afef')
        ax.set_title('Number of Ransomware Attacks Over Time', color='#61afef')
        ax.set_xlabel('Date', color='#ffffff')
        ax.set_ylabel('Number of Attacks', color='#ffffff')
        ax.grid(True, color='#abb2bf')
        plt.xticks(rotation=45, color='#abb2bf')
        plt.yticks(color='#abb2bf')
        fig.patch.set_facecolor('#282c34')
        ax.set_facecolor('#1e2127')

        # Embed the plot in the tkinter window
        canvas = FigureCanvasTkAgg(fig, master=self.master)
        canvas.draw()
        canvas.get_tk_widget().pack()

        # Close the figure after embedding to free memory
        plt.close(fig)

    def export_results(self):
        export_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                   filetypes=[("Text files", "*.txt"),
                                                              ("CSV files", "*.csv"),
                                                              ("PDF files", "*.pdf"),
                                                              ("Excel files", "*.xlsx"),
                                                              ("All files", "*.*")])
        if export_path:
            try:
                if export_path.endswith('.xlsx'):
                    self.df.to_excel(export_path, index=False)
                elif export_path.endswith('.txt') or export_path.endswith('.csv'):
                    with open(export_path, 'w') as f:
                        f.write(self.result_text.get(1.0, tk.END))
                elif export_path.endswith('.pdf'):
                    pdf = FPDF()
                    pdf.add_page()
                    pdf.set_font("Arial", size=12)
                    lines = self.result_text.get(1.0, tk.END).splitlines()
                    for line in lines:
                        pdf.cell(200, 10, txt=line, ln=True)
                    pdf.output(export_path)
                messagebox.showinfo("Export", "Results exported successfully!")
                logging.info("Results exported successfully.")
            except Exception as e:
                logging.error(f"Error during export: {e}")
                messagebox.showerror("Error", f"Error during export: {str(e)}")

    def change_theme(self, theme):
        if theme == 'Dark':
            self.master.configure(bg='#282c34')
            self.label.config(background='#282c34', foreground='#ffffff')
            self.result_text.config(bg='#1e2127', fg='#abb2bf')
        elif theme == 'Light':
            self.master.configure(bg='#f0f0f0')
            self.label.config(background='#f0f0f0', foreground='#000000')
            self.result_text.config(bg='#ffffff', fg='#000000')

    def on_exit(self):
        logging.info("Application closed by user.")
        self.master.quit()
        self.master.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    analyzer = RansomwareAnalyzer(root)
    root.mainloop()
