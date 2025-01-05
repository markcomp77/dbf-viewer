import os
import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableView, QVBoxLayout, QWidget, QComboBox, QLabel,
    QPushButton, QHBoxLayout, QTextEdit
)
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from dbfread import DBF
import pandas as pd

# Mapowanie CP790 -> UTF-8
CP790_TO_UTF8 = {
    0x8F: 'Ą', 0x86: 'ą', 0x95: 'Ć', 0x8D: 'ć',
    0x90: 'Ę', 0x91: 'ę', 0x9C: 'Ł', 0x92: 'ł',
    0xA5: 'Ń', 0xA4: 'ń', 0xA3: 'Ó', 0xA2: 'ó',
    0x98: 'Ś', 0x9E: 'ś', 0xA1: 'Ż', 0xA7: 'ż',
    0xA0: 'Ź', 0xA6: 'ź'
}

def decode_cp790(data):
    return ''.join(CP790_TO_UTF8.get(byte, chr(byte)) for byte in data.encode('latin1'))

def encode_cp790(data):
    reverse_map = {v: k for k, v in CP790_TO_UTF8.items()}
    return ''.join(chr(reverse_map.get(char, ord(char))) for char in data)

def get_unique_filename(base_path, extension):
    counter = 1
    file_path = f"{base_path}.{extension}"
    while os.path.exists(file_path):
        file_path = f"{base_path}_{counter}.{extension}"
        counter += 1
    return file_path

class DBFViewer(QMainWindow):
    def __init__(self, data_dir):
        super().__init__()

        self.setWindowTitle("DBF Viewer Enhanced")
        self.resize(800, 600)

        # Central widget and layout
        central_widget = QWidget()
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # List of databases
        self.db_list = QTextEdit()
        self.db_list.setReadOnly(True)
        main_layout.addWidget(self.db_list)

        # Actions
        actions_layout = QHBoxLayout()
        main_layout.addLayout(actions_layout)

        self.combo_box = QComboBox()
        actions_layout.addWidget(self.combo_box)

        self.view_button = QPushButton("View in Sheet")
        self.view_button.clicked.connect(self.on_view_button_clicked)
        actions_layout.addWidget(self.view_button)

        self.export_csv_button = QPushButton("Export to CSV")
        self.export_csv_button.clicked.connect(self.on_export_csv_clicked)
        actions_layout.addWidget(self.export_csv_button)

        self.export_xlsx_button = QPushButton("Export to XLSX")
        self.export_xlsx_button.clicked.connect(self.on_export_xlsx_clicked)
        actions_layout.addWidget(self.export_xlsx_button)

        # TableView setup
        self.table_view = QTableView()
        main_layout.addWidget(self.table_view)

        # Load DBF file list
        self.data_dir = data_dir
        self.load_dbf_files()

    def load_dbf_files(self):
        self.dbf_files = sorted([f for f in os.listdir(self.data_dir) if f.lower().endswith('.dbf')])
        self.combo_box.addItems(self.dbf_files)
        self.populate_db_list()

    def populate_db_list(self):
        summary = ""
        for dbf_file in self.dbf_files:
            file_path = os.path.join(self.data_dir, dbf_file)
            try:
                dbf_data = DBF(file_path, encoding="latin1")
                headers = dbf_data.field_names
                records = list(dbf_data)[:2]

                summary += f"{dbf_file} (Size: {os.path.getsize(file_path)} bytes)\n"
                summary += f"Fields: {', '.join(headers)}\n"
                for record in records:
                    summary += f"Record: {record}\n"
                summary += "-" * 40 + "\n"
            except Exception as e:
                summary += f"Error reading {dbf_file}: {e}\n"

        self.db_list.setText(summary)

    def load_dbf(self, dbf_file):
        try:
            dbf_data = DBF(dbf_file, load=True, encoding="latin1")
            headers = dbf_data.field_names

            # Prepare the model
            model = QStandardItemModel()
            model.setHorizontalHeaderLabels(headers)

            for record in dbf_data:
                row = [QStandardItem(decode_cp790(str(record[field]))) for field in headers]
                model.appendRow(row)

            # Set the model in the QTableView
            self.table_view.setModel(model)
        except Exception as e:
            print(f"Error loading DBF file: {e}")

    def on_view_button_clicked(self):
        selected_file = os.path.join(self.data_dir, self.combo_box.currentText())
        self.load_dbf(selected_file)

    def export_to_file(self, base_file, file_type):
        selected_file = os.path.join(self.data_dir, base_file)
        base_name, _ = os.path.splitext(base_file)
        output_file = get_unique_filename(os.path.join(self.data_dir, base_name), file_type)
        try:
            dbf_data = DBF(selected_file, load=True, encoding="latin1")
            data = [
                {field: decode_cp790(str(record[field])) for field in dbf_data.field_names}
                for record in dbf_data
            ]
            df = pd.DataFrame(data)
            if file_type == 'csv':
                df.to_csv(output_file, index=False)
            elif file_type == 'xlsx':
                df.to_excel(output_file, index=False, engine='openpyxl')
            print(f"Exported to {output_file}")
        except Exception as e:
            print(f"Error exporting file: {e}")

    def on_export_csv_clicked(self):
        base_file = self.combo_box.currentText()
        self.export_to_file(base_file, 'csv')

    def on_export_xlsx_clicked(self):
        base_file = self.combo_box.currentText()
        self.export_to_file(base_file, 'xlsx')

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Path to your DATA directory
    data_dir_path = "DATA"

    viewer = DBFViewer(data_dir_path)
    viewer.show()

    sys.exit(app.exec_())
