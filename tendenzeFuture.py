import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from statsmodels.tsa.seasonal import seasonal_decompose
from statsmodels.tsa.arima.model import ARIMA
from tkinter import Tk, Button, Toplevel, Text, Scrollbar, VERTICAL, RIGHT, Y, END, Frame, Menu, messagebox, filedialog
from tkinter import ttk

# Funzione per caricare i dati e preparare le variabili globali
def load_data(file_path):
    global data, attacks_per_year, attacks_per_month, monthly_attacks, decomposition, fit_model, forecast

    # Simulare un'operazione lenta con una progress bar
    progress['value'] = 0
    root.update_idletasks()

    # Caricare il dataset
    data = pd.read_excel(file_path)

    progress['value'] = 20
    root.update_idletasks()

    # Convertire la colonna 'date' in formato datetime
    data['date'] = pd.to_datetime(data['date'])

    progress['value'] = 40
    root.update_idletasks()

    # Creare colonne per anno e mese
    data['year'] = data['date'].dt.year
    data['month'] = data['date'].dt.month

    progress['value'] = 60
    root.update_idletasks()

    # Contare il numero di attacchi per anno e per mese
    attacks_per_year = data['year'].value_counts().sort_index()
    attacks_per_month = data.groupby(['year', 'month']).size().unstack().fillna(0)

    progress['value'] = 80
    root.update_idletasks()

    # Aggregare i dati per mese utilizzando 'ME' per Month End
    monthly_attacks = data.set_index('date').resample('ME').size()

    # Decomposizione delle serie temporali utilizzando il modello additivo
    decomposition = seasonal_decompose(monthly_attacks, model='additive')

    progress['value'] = 90
    root.update_idletasks()

    # Creare un modello ARIMA
    model = ARIMA(monthly_attacks, order=(1, 1, 1))
    fit_model = model.fit()

    # Prevedere i prossimi 12 mesi
    forecast = fit_model.forecast(steps=12)

    progress['value'] = 100
    root.update_idletasks()

# Funzione per selezionare il file
def select_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if file_path:
        load_data(file_path)

# Funzioni per visualizzare i grafici
def plot_attacks_per_year():
    plt.figure(figsize=(10, 6))
    sns.lineplot(x=attacks_per_year.index, y=attacks_per_year.values, marker="o")
    plt.title("Number of Ransomware Attacks per Year")
    plt.xlabel("Year")
    plt.ylabel("Number of Attacks")
    plt.grid(True)
    plt.show()

def plot_attacks_per_month():
    plt.figure(figsize=(14, 8))
    sns.heatmap(attacks_per_month, annot=True, fmt=".0f", cmap="YlGnBu", cbar=True)
    plt.title("Monthly Distribution of Ransomware Attacks")
    plt.xlabel("Month")
    plt.ylabel("Year")
    plt.show()

def plot_decomposition():
    decomposition.plot()
    plt.show()

def plot_forecast():
    plt.figure(figsize=(10, 6))
    plt.plot(monthly_attacks.index, monthly_attacks, label='Observed')
    plt.plot(forecast.index, forecast, label='Forecast', color='red')
    plt.title('Forecast of Ransomware Attacks for the Next 12 Months')
    plt.xlabel('Date')
    plt.ylabel('Number of Attacks')
    plt.legend()
    plt.grid(True)
    plt.show()

def show_documentation():
    doc_window = Toplevel(root)
    doc_window.title("Documentation")
    doc_window.geometry("700x400")

    doc_text = Text(doc_window, wrap='word', width=80, height=20)
    scrollbar = Scrollbar(doc_window, orient=VERTICAL, command=doc_text.yview)
    doc_text['yscrollcommand'] = scrollbar.set

    scrollbar.pack(side=RIGHT, fill=Y)
    doc_text.pack(fill='both', expand=True)

    documentation = (
        "1. Number of Ransomware Attacks per Year:\n"
        "This diagram shows the total number of ransomware attacks per year. It helps to identify trends over "
        "the years and observe any significant increases or decreases in attacks.\n\n"

        "2. Monthly Distribution of Ransomware Attacks:\n"
        "This heatmap displays the distribution of ransomware attacks by month for each year. It helps to "
        "identify if there are specific months that are more prone to attacks and observe seasonal trends.\n\n"

        "3. Decomposition of Time Series:\n"
        "This plot shows the decomposition of the time series data into three components: Trend, Seasonality, "
        "and Residuals. The trend component reveals the general direction of attacks over time, the seasonal "
        "component shows recurring patterns, and the residual component captures random fluctuations.\n\n"

        "4. Forecast of Ransomware Attacks for the Next 12 Months:\n"
        "This plot provides a forecast of ransomware attacks for the next 12 months using the ARIMA model. "
        "It helps to anticipate future trends and prepare for potential increases in attacks."
    )

    doc_text.insert(END, documentation)
    doc_text.config(state='disabled')

def about():
    messagebox.showinfo("About", "Ransomware Attacks Analysis Tool\nVersion 1.0\nDeveloped by [Fabian Kofler]\nfabiankofler5@gmail.com")

# Creare la GUI
def create_gui():
    global root, progress
    root = Tk()
    root.title("Ransomware Attacks Analysis")
    root.geometry("800x500")  # Imposta la dimensione iniziale della finestra

    # Creare una barra dei menu
    menubar = Menu(root)
    root.config(menu=menubar)

    # Creare un sottomenu "File"
    file_menu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="Open", command=select_file)
    file_menu.add_command(label="Exit", command=root.quit)

    # Creare un sottomenu "Help"
    help_menu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="Documentation", command=show_documentation)
    help_menu.add_command(label="About", command=about)

    # Creare una progress bar
    progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
    progress.pack(pady=20)

    # Creare un frame per i bottoni
    button_frame = Frame(root)
    button_frame.pack(pady=20, fill="x")

    # Creare i bottoni e posizionarli con pack()
    ttk.Button(button_frame, text="Show Attacks per Year", command=plot_attacks_per_year).pack(side="left", padx=10,
                                                                                               pady=10, expand=True)
    ttk.Button(button_frame, text="Show Attacks per Month", command=plot_attacks_per_month).pack(side="left", padx=10,
                                                                                                 pady=10, expand=True)
    ttk.Button(button_frame, text="Show Decomposition", command=plot_decomposition).pack(side="left", padx=10, pady=10,
                                                                                         expand=True)
    ttk.Button(button_frame, text="Show Forecast", command=plot_forecast).pack(side="left", padx=10, pady=10,
                                                                               expand=True)

    root.mainloop()

# Creare e avviare la GUI
create_gui()
