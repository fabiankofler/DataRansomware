import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QListWidget, QTableView, QHeaderView, \
    QStatusBar, QWidget, QLabel, QLineEdit, QAbstractItemView
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem
import pandas as pd


class DataFrameViewer(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Dataset Ransomware GUI")
        self.setGeometry(100, 100, 1000, 600)

        # Applicare stili con CSS
        self.setStyleSheet("""
            QMainWindow {
                background-color: #2E3440;
            }
            QLabel {
                color: #88C0D0;
                font-weight: bold;
            }
            QListWidget, QLineEdit {
                background-color: #3B4252;
                color: #ECEFF4;
                padding: 5px;
                border-radius: 5px;
            }
            QPushButton {
                background-color: #5E81AC;
                color: #ECEFF4;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #81A1C1;
            }
            QStatusBar {
                background-color: #4C566A;
                color: #ECEFF4;
            }
            QTableView {
                background-color: #ECEFF4;
                color: #2E3440;
                gridline-color: #D8DEE9;
                selection-background-color: #5E81AC;
            }
            QHeaderView::section {
                background-color: #4C566A;
                color: #ECEFF4;
                padding: 5px;
                border: 1px solid #D8DEE9;
            }
        """)

        # Layout principale
        self.main_layout = QVBoxLayout()

        # Layout superiore per i filtri
        self.filter_layout = QHBoxLayout()

        # Filtro per la data
        self.date_label = QLabel("Date Filter:", self)
        self.filter_layout.addWidget(self.date_label)
        self.date_filter = QListWidget(self)
        self.date_filter.setSelectionMode(QAbstractItemView.MultiSelection)
        self.filter_layout.addWidget(self.date_filter)

        # Filtro per la gang
        self.gang_label = QLabel("Gang Filter:", self)
        self.filter_layout.addWidget(self.gang_label)
        self.gang_filter = QListWidget(self)
        self.gang_filter.setSelectionMode(QAbstractItemView.MultiSelection)
        self.filter_layout.addWidget(self.gang_filter)

        # Filtro per il settore
        self.sector_label = QLabel("Sector Filter:", self)
        self.filter_layout.addWidget(self.sector_label)
        self.sector_filter = QListWidget(self)
        self.sector_filter.setSelectionMode(QAbstractItemView.MultiSelection)
        self.filter_layout.addWidget(self.sector_filter)

        # Filtro per il paese della vittima
        self.country_label = QLabel("Victim Country Filter:", self)
        self.filter_layout.addWidget(self.country_label)
        self.country_filter = QListWidget(self)
        self.country_filter.setSelectionMode(QAbstractItemView.MultiSelection)
        self.filter_layout.addWidget(self.country_filter)

        # Barra di ricerca generale
        self.search_bar = QLineEdit(self)
        self.search_bar.setPlaceholderText("Search")
        self.filter_layout.addWidget(self.search_bar)

        self.search_button = QPushButton("Search", self)
        self.search_button.clicked.connect(self.search_data)
        self.filter_layout.addWidget(self.search_button)

        self.reset_button = QPushButton("Reset", self)
        self.reset_button.clicked.connect(self.reset_view)
        self.filter_layout.addWidget(self.reset_button)

        # Aggiungi il layout dei filtri al layout principale
        self.main_layout.addLayout(self.filter_layout)

        # Tabella per visualizzare i dati
        self.table_view = QTableView(self)
        self.main_layout.addWidget(self.table_view)

        # Contenitore centrale
        self.container = QWidget()
        self.container.setLayout(self.main_layout)
        self.setCentralWidget(self.container)

        # Barra di stato
        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)

        # Caricamento del dataset
        self.df = None
        self.df_gui = None

        # Carica il file all'avvio
        self.load_data()

    def load_data(self):
        file_path = 'Dataset Ransomware.xlsx'  # Percorso al file caricato
        if file_path:
            try:
                self.df = pd.read_excel(file_path, sheet_name='Dataset')

                # Verifica che la colonna 'Date' esista
                if 'date' not in self.df.columns:
                    raise ValueError("La colonna 'Date' non esiste nel dataset.")

                # Assicurati che la colonna 'Date' sia nel formato datetime
                self.df['date'] = pd.to_datetime(self.df['date'], errors='coerce')

                # Converte la colonna 'number of employees' in stringa per evitare che sia interpretata come data
                if 'number of employees' in self.df.columns:
                    self.df['number of employees'] = self.df['number of employees'].astype(str)

                self.df_gui = self.df.copy()  # Fai una copia del dataframe per la visualizzazione
                self.df_gui['id'] = self.df_gui['id'].fillna(0).astype(int)  # Rimuovi i decimali dalla colonna 'id'
                self.populate_filters()
                self.display_data()
                self.status_bar.showMessage(f"Loaded file: {file_path}")
            except Exception as e:
                self.status_bar.showMessage(f"Error loading file: {str(e)}")
                print(f"Error: {e}")  # Debug: stampa l'errore in console

    def populate_filters(self):
        # Popola le liste con i valori unici del dataset
        self.date_filter.addItem("All")
        self.date_filter.addItems(sorted(self.df_gui['date'].dropna().dt.strftime('%Y-%m-%d').unique()))

        self.gang_filter.addItem("All")
        self.gang_filter.addItems(sorted(self.df_gui['gang'].dropna().unique().astype(str)))

        self.sector_filter.addItem("All")
        self.sector_filter.addItems(sorted(self.df_gui['Victim sectors'].dropna().unique().astype(str)))

        self.country_filter.addItem("All")
        self.country_filter.addItems(sorted(self.df_gui['Victim Country'].dropna().unique().astype(str)))

    def get_selected_items(self, list_widget):
        """Restituisce un elenco di elementi selezionati da un QListWidget."""
        selected_items = [item.text() for item in list_widget.selectedItems()]
        return selected_items if "All" not in selected_items else []

    def search_data(self):
        # Filtro generale
        query = self.search_bar.text().lower()
        filtered_df = self.df.copy()

        # Filtro per la data
        selected_dates = self.get_selected_items(self.date_filter)
        if selected_dates:
            filtered_df = filtered_df[filtered_df['date'].dt.strftime('%Y-%m-%d').isin(selected_dates)]

        # Filtro per la gang
        selected_gangs = self.get_selected_items(self.gang_filter)
        if selected_gangs:
            filtered_df = filtered_df[filtered_df['gang'].isin(selected_gangs)]

        # Filtro per il settore
        selected_sectors = self.get_selected_items(self.sector_filter)
        if selected_sectors:
            filtered_df = filtered_df[filtered_df['Victim sectors'].isin(selected_sectors)]

        # Filtro per il paese della vittima
        selected_countries = self.get_selected_items(self.country_filter)
        if selected_countries:
            filtered_df = filtered_df[filtered_df['Victim Country'].isin(selected_countries)]

        # Filtro generale
        if query:
            filtered_df = filtered_df[filtered_df.apply(lambda row: query in row.to_string().lower(), axis=1)]

        self.df_gui = filtered_df
        self.display_data()

        if filtered_df.empty:
            self.status_bar.showMessage("No matching records found")
        else:
            self.status_bar.showMessage("Search completed")

    def reset_view(self):
        self.df_gui = self.df.copy()
        self.df_gui['id'] = self.df_gui['id'].fillna(0).astype(int)
        self.display_data()
        self.status_bar.showMessage("View reset to full dataset")

    def display_data(self):
        model = QStandardItemModel(self.df_gui.shape[0], self.df_gui.shape[1])
        model.setHorizontalHeaderLabels(self.df_gui.columns)

        for row in range(self.df_gui.shape[0]):
            for column in range(self.df_gui.shape[1]):
                item = QStandardItem(str(self.df_gui.iat[row, column]))
                item.setEditable(False)  # Rendi l'elemento non modificabile
                model.setItem(row, column, item)

        self.table_view.setModel(model)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_view.setEditTriggers(
            QTableView.NoEditTriggers)  # Disabilita la modifica tramite doppio clic o tasto invio


def main():
    app = QApplication(sys.argv)
    viewer = DataFrameViewer()
    viewer.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
