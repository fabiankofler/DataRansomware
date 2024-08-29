from io import BytesIO
import mplcursors
import pandas as pd
import os
import geopandas as gpd
import matplotlib.pyplot as plt
from tkinter import *
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter import filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import sys
import re
import plotly.express as px
from PIL import Image, ImageTk
import plotly.io as pio

# Variabile per controllare lo stato di riproduzione
is_playing = False

# Function to load the dataset
def load_data(file_path):
    try:
        data = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit()

    required_columns = ['id', 'date', 'Victim Country', 'Victim sectors', 'gang', 'Number of employees',
                        'Ransom Currency', 'Ransom Paid']
    missing_columns = [col for col in required_columns if col not in data.columns]
    if missing_columns:
        print(f"Missing required columns in Excel file: {missing_columns}")
        sys.exit()

    data_cleaned = data.drop(columns=['id', 'Ransom Currency', 'Ransom Paid'])
    data_cleaned = data_cleaned.dropna(subset=['Victim Country', 'Victim sectors', 'gang'])

    data_cleaned['date'] = pd.to_datetime(data_cleaned['date'], errors='coerce')
    data_cleaned = data_cleaned.dropna(subset=['date'])

    data_cleaned['Victim Country'] = data_cleaned['Victim Country'].str.strip().str.lower().replace({
        'usa': 'united states',
        'uk': 'united kingdom',
        'south\xa0korea': 'south korea',
        'non disponibile': None,
        'potugal': 'portugal',
        'ivory\xa0coast': 'ivory coast',
        'india': 'india'
    })
    data_cleaned['Victim sectors'] = data_cleaned['Victim sectors'].str.strip().str.lower()

    data_cleaned['Original Victim Country'] = data_cleaned['Victim Country']

    def clean_number_of_employees(value):
        if pd.isna(value):
            return None
        value = str(value).strip()
        if '-' in value:
            parts = value.split('-')
            return (int(parts[0]) + int(parts[1])) // 2
        elif any(char in value for char in ['<', '>']):
            return int(re.sub(r'[^\d]', '', value))
        elif value.isdigit():
            return int(value)
        else:
            return None

    data_cleaned['Number of employees'] = data_cleaned['Number of employees'].apply(clean_number_of_employees).astype(
        'Int64')

    data_cleaned = pd.get_dummies(data_cleaned, columns=['Victim Country', 'Victim sectors'], drop_first=True)

    return data_cleaned


# Load the dataset
file_path = 'Dataset Ransomware.xlsx'
data_cleaned = load_data(file_path)

# Function to handle select all gangs
def toggle_select_all():
    if select_all_var.get():
        gangs_listbox.select_set(0, END)  # Select all gangs
    else:
        gangs_listbox.select_clear(0, END)  # Deselect all gangs


# Function to get filtered data
def get_filtered_data():
    if all_time_var.get():
        # If "All Time" is selected, use the full dataset
        filtered_data = data_cleaned.copy()
    else:
        # Apply date range filter as usual
        start_date = pd.to_datetime(start_cal.get_date())
        end_date = pd.to_datetime(end_cal.get_date())
        filtered_data = data_cleaned[(data_cleaned['date'] >= start_date) & (data_cleaned['date'] <= end_date)]

    selected_country = country_combobox.get()
    selected_gangs = [gangs_listbox.get(i) for i in gangs_listbox.curselection()]

    if selected_country != "All":
        country_column = f'Victim Country_{selected_country.lower()}'
        if country_column in filtered_data.columns:
            filtered_data = filtered_data[filtered_data[country_column] == 1]

    if selected_gangs:
        filtered_data = filtered_data[filtered_data['gang'].isin(selected_gangs)]

    return filtered_data, selected_country, selected_gangs


# Function to update the chart
def update_chart():
    global current_fig
    filtered_data, selected_country, active_gangs = get_filtered_data()

    plt.style.use('ggplot')
    fig, ax = plt.subplots(figsize=(16, 8))

    if not filtered_data.empty and active_gangs:
        colors = plt.colormaps['tab10'].resampled(len(active_gangs))

        plots = []
        data_mapping = []

        for i, gang in enumerate(active_gangs):
            gang_data = filtered_data[filtered_data['gang'] == gang]
            if not gang_data.empty:
                time_series_data = gang_data.groupby(gang_data['date'].dt.date).size().reset_index(name='attack_count')
                plot, = ax.plot(time_series_data['date'], time_series_data['attack_count'], marker='o', linestyle='-',
                                color=colors(i), linewidth=0.5, markersize=4, label=gang)
                plots.append(plot)
                data_mapping.append(time_series_data)

        ax.set_title(f'Time Series of Ransomware Attacks - {selected_country}', fontsize=20, fontweight='bold')
        ax.set_xlabel('Date', fontsize=16)
        ax.set_ylabel('Number of Attacks', fontsize=16)
        ax.grid(True, which='both', linestyle='--', linewidth=0.7, alpha=0.7)
        ax.set_facecolor('#f5f5f5')
        ax.tick_params(axis='both', which='major', labelsize=14)

        fig.autofmt_xdate(rotation=45)

        cursor = mplcursors.cursor(plots, hover=True)
        cursor.connect("add", lambda sel: sel.annotation.set_text(
            f"Gang: {active_gangs[plots.index(sel.artist)]}\nDate: {data_mapping[plots.index(sel.artist)]['date'].iloc[int(sel.index)]}\nAttacks: {data_mapping[plots.index(sel.artist)]['attack_count'].iloc[int(sel.index)]}"
        ))
    else:
        ax.text(0.5, 0.5, 'No data available for this selection',
                horizontalalignment='center', verticalalignment='center', transform=ax.transAxes, fontsize=16)
        ax.set_title('No data available', fontsize=20)
        ax.axis('off')

    for widget in chart_frame.winfo_children():
        widget.destroy()

    chart = FigureCanvasTkAgg(fig, master=chart_frame)
    chart.draw()
    chart.get_tk_widget().pack(fill=BOTH, expand=True)

    current_fig = fig


# Funzione per trovare la prossima data valida con attacchi
def find_next_valid_dates(start_date, end_date, direction):
    while True:
        filtered_data, _, _ = get_filtered_data_with_dates(start_date, end_date)
        if not filtered_data.empty:
            return start_date, end_date

        # Aggiorna le date per cercare nel periodo successivo o precedente
        start_date += pd.DateOffset(months=direction)
        end_date += pd.DateOffset(months=direction)

        # Imposta i limiti per evitare di andare oltre i dati disponibili
        min_date = data_cleaned['date'].min()
        max_date = data_cleaned['date'].max()

        if start_date < min_date:
            start_date = min_date
            end_date = start_date + pd.DateOffset(months=1)
        elif end_date > max_date:
            end_date = max_date
            start_date = end_date - pd.DateOffset(months=1)

        # Se si raggiungono i limiti, fermarsi
        if start_date == min_date and direction < 0:
            return min_date, min_date + pd.DateOffset(months=1)
        if end_date == max_date and direction > 0:
            return max_date - pd.DateOffset(months=1), max_date

# Funzione per aggiornare la mappa con i dati filtrati
def update_world_map(time_window):
    global current_start_date, current_end_date

    # Aggiorna la finestra temporale
    current_start_date += pd.DateOffset(months=time_window)
    current_end_date += pd.DateOffset(months=time_window)

    # Trova il prossimo intervallo di date valide con attacchi
    current_start_date, current_end_date = find_next_valid_dates(current_start_date, current_end_date, time_window)

    # Filtra i dati con le nuove date
    filtered_data, _, _ = get_filtered_data_with_dates(current_start_date, current_end_date)

    # Se non ci sono dati disponibili, visualizza un messaggio inform
    if filtered_data.empty:
        canvas.itemconfig(image_item, image=img_placeholder)
        canvas.image = img_placeholder
        return

        # Conta il numero di attacchi per paese nei dati filtrati
    country_attacks = filtered_data.groupby('Original Victim Country').size().reset_index(name='attacks')

    # Crea la mappa con Plotly
    fig = px.choropleth(
        country_attacks,
        locations="Original Victim Country",
        locationmode="country names",
        color="attacks",
        hover_name="Original Victim Country",
        hover_data=["attacks"],
        color_continuous_scale=px.colors.sequential.Viridis,
        title=f"Global Distribution of Ransomware Attacks ({current_start_date.date()} - {current_end_date.date()})",
        labels={'attacks': 'Numero di Attacchi'},
    )

    fig.update_layout(
        title_font=dict(size=24, color='darkblue', family="Arial Black"),
        geo=dict(
            showframe=False,
            showcoastlines=True,
            coastlinecolor="Black",
            projection_type='equirectangular',
            bgcolor="aliceblue"
        ),
        coloraxis_colorbar=dict(
            title="Attacks",
            ticks="outside",
            lenmode="fraction",
            len=0.8,
            thicknessmode="pixels",
            thickness=15,
            tickfont=dict(color='darkblue', size=15),
            titlefont=dict(color='darkblue', size=20)
        )
    )

    # Converti la figura di Plotly in un'immagine e aggiornala nella finestra
    img_bytes = pio.to_image(fig, format="png", width=1200, height=600)
    img = Image.open(BytesIO(img_bytes))
    img_tk = ImageTk.PhotoImage(img)

    canvas.itemconfig(image_item, image=img_tk)
    canvas.image = img_tk


# Funzione per filtrare i dati con l'intervallo temporale attuale
def get_filtered_data_with_dates(start_date, end_date):
    filtered_data = data_cleaned[(data_cleaned['date'] >= start_date) & (data_cleaned['date'] <= end_date)]
    return filtered_data, None, None


# Funzione per mostrare la mappa del mondo con navigazione temporale nella stessa finestra
def show_world_map_with_navigation():
    global current_start_date, current_end_date, canvas, image_item, img_placeholder

    # Utilizza le date selezionate dall'utente
    current_start_date = pd.to_datetime(start_cal.get_date())
    current_end_date = pd.to_datetime(end_cal.get_date())

    # Crea una nuova finestra per la mappa del mondo
    world_map_window = Toplevel(root)
    world_map_window.title("Mappa Mondiale degli Attacchi Ransomware")

    # Canvas per visualizzare l'immagine
    canvas = Canvas(world_map_window, width=1200, height=600)
    canvas.pack(fill=BOTH, expand=True)

    # Crea un placeholder per l'immagine
    img_placeholder = ImageTk.PhotoImage(Image.new('RGB', (1000, 600), color=(255, 255, 255)))
    image_item = canvas.create_image(0, 0, anchor=NW, image=img_placeholder)

    # Frame per i bottoni di navigazione temporale sotto la mappa
    navigation_frame = Frame(world_map_window)
    navigation_frame.pack(fill=X)

    # Aggiungi bottoni per navigare nel tempo
    Button(navigation_frame, text="<< Previous", command=lambda: update_world_map(-1)).pack(side=LEFT, padx=10,
                                                                                              pady=10)
    Button(navigation_frame, text="Next >>", command=lambda: update_world_map(1)).pack(side=LEFT, padx=10,
                                                                                             pady=10)

    # Aggiungi bottoni per il controllo del play automatico
    Button(navigation_frame, text="Play", command=play_map).pack(side=LEFT, padx=10, pady=10)
    Button(navigation_frame, text="Stop", command=stop_play).pack(side=LEFT, padx=10, pady=10)

    # Mostra la mappa iniziale con il range di date selezionato
    update_world_map(0)



# Funzione per avviare la riproduzione automatica
def play_map():
    global is_playing
    if not is_playing:
        is_playing = True
        start_play()


# Funzione che gestisce la riproduzione automatica
def start_play():
    global current_start_date, current_end_date, is_playing

    if not is_playing:
        return

    # Aggiorna la mappa al prossimo intervallo
    update_world_map(1)

    # Verifica se abbiamo raggiunto la fine del dataset
    max_date = data_cleaned['date'].max()
    if current_end_date >= max_date:
        is_playing = False
        return

    # Imposta il timer per l'aggiornamento successivo
    root.after(3000, start_play)  # 3000 ms = 3 secondi


# Funzione per fermare la riproduzione automatica
def stop_play():
    global is_playing
    is_playing = False


# Funzione per esportare i dati
def export_data():
    filtered_data, selected_country, selected_gangs = get_filtered_data()
    filtered_data = filtered_data.loc[:, (filtered_data != False).any(axis=0)]

    # Select only columns that exist in the filtered DataFrame
    relevant_columns = ['date', 'gang', 'Original Victim Country', 'Victim sectors', 'Number of employees']
    columns_to_save = [col for col in relevant_columns if col in filtered_data.columns]

    # Add the country column if "All" is selected
    if selected_country == "All":
        filtered_data = filtered_data[columns_to_save]
    else:
        filtered_data = filtered_data[[col for col in columns_to_save if col != 'Original Victim Country']]

    file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                             filetypes=[("CSV files", "*.csv")],
                                             title="Save file")
    if file_path:
        filtered_data.to_csv(file_path, index=False)


# Funzione per esportare il grafico come PDF
def export_graph_pdf():
    if current_fig:
        file_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                 filetypes=[("PDF files", "*.pdf")],
                                                 title="Save Graph as PDF")
        if file_path:
            current_fig.savefig(file_path)


# Funzione per abilitare/disabilitare la selezione delle date
def toggle_all_time():
    if all_time_var.get():
        start_cal.config(state=DISABLED)
        end_cal.config(state=DISABLED)
    else:
        start_cal.config(state=NORMAL)
        end_cal.config(state=NORMAL)


# Funzione per chiudere la finestra e terminare l'applicazione
def on_closing():
    root.destroy()  # Chiude la finestra di Tkinter
    sys.exit()  # Termina l'applicazione


# Funzione per mostrare le gang in una nuova finestra
def show_gangs_window():
    filtered_data, _, active_gangs = get_filtered_data()

    if not filtered_data.empty:
        colors = plt.colormaps.get_cmap('tab20').resampled(len(active_gangs))

        gang_window = Toplevel(root)
        gang_window.title("Active Gangs in the Chart")

        max_columns = 5

        for i, gang in enumerate(sorted(active_gangs)):
            row = i // max_columns
            column = i % max_columns

            hex_color = rgba_to_hex(colors(i))

            color_label = Label(gang_window, bg=hex_color, width=2, height=1)
            color_label.grid(row=row, column=column * 2, padx=5, pady=5)

            gang_label = Label(gang_window, text=gang, font=("Arial", 12))
            gang_label.grid(row=row, column=column * 2 + 1, padx=5, pady=5)

    else:
        gang_window = Toplevel(root)
        gang_window.title("Active Gangs in the Chart")
        Label(gang_window, text="No gangs found for the selected date range.", font=("Arial", 12)).pack(anchor=CENTER,
                                                                                                        pady=20)


# Funzione per convertire i colori da RGBA a HEX
def rgba_to_hex(color):
    r = int(color[0] * 255)
    g = int(color[1] * 255)
    b = int(color[2] * 255)
    return f'#{r:02x}{g:02x}{b:02x}'


# Creazione dell'interfaccia utente
root = Tk()
root.title("Ransomware Attack Visualization")

# Configurazione della chiusura della finestra
root.protocol("WM_DELETE_WINDOW", on_closing)

# Frame per i controlli a sinistra
controls_frame = Frame(root)
controls_frame.pack(side=LEFT, fill=Y, padx=10, pady=10)
chart_frame = Frame(root)
chart_frame.pack(side=RIGHT, fill=BOTH, expand=True)

# Selezione del range temporale
Label(controls_frame, text="Start Date:").pack(anchor=W, pady=2)
start_cal = DateEntry(controls_frame, selectmode='day')
start_cal.pack(anchor=W, pady=2)

Label(controls_frame, text="End Date:").pack(anchor=W, pady=2)
end_cal = DateEntry(controls_frame, selectmode='day')
end_cal.pack(anchor=W, pady=2)

# Opzione All Time
all_time_var = BooleanVar(value=False)
all_time_checkbox = Checkbutton(controls_frame, text="All Time", variable=all_time_var, command=toggle_all_time)
all_time_checkbox.pack(anchor=W, pady=2)

# Selezione del paese
Label(controls_frame, text="Select Country:").pack(anchor=W, pady=2)
countries = ["All"] + sorted(
    [col.split('_')[-1].capitalize() for col in data_cleaned.columns if col.startswith('Victim Country_')])
country_combobox = ttk.Combobox(controls_frame, values=countries, state='readonly')
country_combobox.current(0)
country_combobox.pack(anchor=W, pady=2)

# Selezione della gang (lista multi-selezione con opzione Select All)
select_all_var = BooleanVar(value=False)
select_all_checkbox = Checkbutton(controls_frame, text="Select All", variable=select_all_var, command=toggle_select_all)
select_all_checkbox.pack(anchor=W, pady=2)

gangs = sorted(data_cleaned['gang'].unique())
gangs_listbox = Listbox(controls_frame, selectmode=MULTIPLE, height=8, exportselection=0)
for gang in gangs:
    gangs_listbox.insert(END, gang)
gangs_listbox.pack(anchor=W, pady=2, fill=X)

# Bottone per mostrare le gang in una nuova finestra
show_gangs_button = Button(controls_frame, text="Show Gangs", command=show_gangs_window)
show_gangs_button.pack(anchor=W, pady=10)

# Bottone per mostrare il grafico
show_chart_button = Button(controls_frame, text="Show Chart", command=update_chart)
show_chart_button.pack(anchor=W, pady=10)

# Bottone per mostrare la mappa del mondo
world_map_button = Button(controls_frame, text="Show World Map", command=show_world_map_with_navigation)
world_map_button.pack(anchor=W, pady=10)

# Bottone per esportare i dati
export_button = Button(controls_frame, text="Export CSV", command=export_data)
export_button.pack(anchor=W, pady=10)

# Bottone per esportare il grafico come PDF
export_pdf_button = Button(controls_frame, text="Export Chart as PDF", command=export_graph_pdf)
export_pdf_button.pack(anchor=W, pady=10)

# Avvia l'applicazione
root.mainloop()