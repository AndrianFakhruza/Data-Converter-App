import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QLineEdit, QMessageBox, QComboBox, QScrollArea
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import pandas as pd
import matplotlib.pyplot as plt

class InputDataPemeriksaan(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Input Data Pemeriksaan')
        self.setStyleSheet("background-color: #ffffff;")
        self.setGeometry(100, 100, 480, 800)  
        self.setFixedSize(480, 800)  
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()

        title_label = QLabel("Input Data Pemeriksaan")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("font-family: 'Roboto', sans-serif; font-size: 24px; color: #4285f4; text-transform: uppercase;")
        layout.addWidget(title_label)

        self.month_label = QLabel("Pilih Bulan:")
        self.month_label.setStyleSheet("font-family: 'Roboto', sans-serif; font-size: 16px; color: #333333;")
        layout.addWidget(self.month_label)

        self.month_combo_box = QComboBox()
        self.month_combo_box.setMaxVisibleItems(12)  
        self.month_combo_box.setMinimumWidth(200)
        self.month_combo_box.setStyleSheet(
            "font-family: 'Roboto', sans-serif; font-size: 14px; background-color: #ffffff; border: 1px solid #cccccc; border-radius: 4px; padding: 5px 10px;"
            "QComboBox::drop-down { border: none; }"
            "QComboBox::down-arrow { image: url(down_arrow.png); width: 16px; height: 16px; }"
        )
        self.month_combo_box.addItems(["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"])
        layout.addWidget(self.month_combo_box)

        scroll_area = QScrollArea()
        scroll_area.setStyleSheet("QScrollArea { border: none; }"
                                  "QScrollBar:vertical { background: #f0f0f0; width: 12px; border-radius: 6px; margin: 0px 2px; }"
                                  "QScrollBar::handle:vertical { background: #cccccc; min-height: 20px; border-radius: 6px; }"
                                  "QScrollBar::handle:vertical:hover { background: #999999; }"
                                  "QScrollBar::sub-line:vertical, QScrollBar::add-line:vertical { height: 0px; }"
                                  "QScrollBar::up-arrow:vertical, QScrollBar::down-arrow:vertical { width: 0px; }"
                                  "QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: none; }")
        scroll_area.setWidgetResizable(True)
        self.scroll_content = QWidget(scroll_area)
        scroll_area.setWidget(self.scroll_content)
        layout_scroll = QVBoxLayout(self.scroll_content)

        self.nama_Pemeriksaan_list = [
            "SPILIS", "HIV", "HbSAg", "Asam Urat", "WIDAL", "TBC", "Kadar Gula Darah", 
            "Golongan Darah", "DARAH RUTIN", "URINALISA", "CHOLESTEROL", "MALARIA", 
            "TRIGLISERIDA", "DBD", "SGOT", "SGPT", "T.PROTEIN", "R.FAKTOR", "CREATININT", 
            "UREA", "CAMPAK", "ALBUMIN", "BIL.DIRECT", "Rapid Antigen"
        ]

        self.nama_Pemeriksaan_list = sorted([nama.upper() for nama in self.nama_Pemeriksaan_list])

        self.labels = []
        self.line_edits = []

        for nama_Pemeriksaan in self.nama_Pemeriksaan_list:
            label = QLabel(nama_Pemeriksaan)
            label.setStyleSheet("font-family: 'Roboto', sans-serif; font-size: 14px; color: #333333;")
            line_edit = QLineEdit()
            line_edit.setStyleSheet("font-family: 'Roboto', sans-serif; font-size: 14px; background-color: #ffffff; border: 1px solid #cccccc; border-radius: 4px; padding: 5px 10px;")
            line_edit.returnPressed.connect(self.focus_next_input)
            layout_scroll.addWidget(label)
            layout_scroll.addWidget(line_edit)
            self.labels.append(label)
            self.line_edits.append(line_edit)

        layout.addWidget(scroll_area)

        self.submit_button = QPushButton('Submit')
        self.submit_button.setStyleSheet("QPushButton { font-family: 'Roboto', sans-serif; font-size: 16px; background-color: #4caf50; color: #ffffff; border: none; border-radius: 4px; padding: 10px 20px; }"
                                          "QPushButton:hover { background-color: #388e3c; }"
                                          "QPushButton:pressed { background-color: #1b5e20; }")
        self.submit_button.clicked.connect(self.submit_data)
        layout.addWidget(self.submit_button)

        self.setLayout(layout)

    def focus_next_input(self):
        current_line_edit = self.sender()
        current_index = self.line_edits.index(current_line_edit)
        next_index = current_index + 1

        if next_index < len(self.line_edits):
            self.line_edits[next_index].setFocus()
        else:
            self.submit_data()

    def submit_data(self):
        data = []
        for label, line_edit in zip(self.labels, self.line_edits):
            nama_Pemeriksaan = label.text()
            jumlah = line_edit.text()
            if jumlah.isdigit():
                data.append({'Nama Pemeriksaan': nama_Pemeriksaan, 'Jumlah': int(jumlah)})
            else:
                QMessageBox.warning(self, 'Input Error', f'Invalid input for {nama_Pemeriksaan}. Please enter a valid number.')
                return  

        if data:
            selected_month = self.month_combo_box.currentText().upper()
            self.simpan_ke_excel(data, f'GRAFIK DATA JUMLAH PEMERIKSAAN LABORATORIUM BULAN {selected_month} 2024.xlsx')
            self.buat_grafik(data, f'GRAFIK DATA JUMLAH PEMERIKSAAN LABORATORIUM BULAN {selected_month} 2024', f'GRAFIK DATA JUMLAH PEMERIKSAAN LABORATORIUM BULAN {selected_month} 2024.png')

    def simpan_ke_excel(self, data, nama_file):
        df = pd.DataFrame(data)
        df.to_excel(nama_file, index=False)
        QMessageBox.information(None, 'Information', f'Data berhasil disimpan ke {nama_file}')

    def buat_grafik(self, data, title, nama_file):
        df = pd.DataFrame(data)
        df = df.sort_values(by='Jumlah', ascending=False)
        plt.figure(figsize=(11.69, 8.27))
        bars = plt.bar(df['Nama Pemeriksaan'], df['Jumlah'], color='blue')

        for bar in bars:
            yval = bar.get_height()
            plt.text(bar.get_x() + bar.get_width()/2, yval, int(yval), ha='center', va='bottom', fontsize=10)

        plt.title(title)
        plt.xlabel('Nama Pemeriksaan')
        plt.ylabel('Jumlah Pemeriksaan Lab')
        plt.xticks(rotation=45, ha='right', fontsize=8)
        plt.ylim(0, max(df['Jumlah']) + 10)
        plt.tight_layout()
        plt.savefig(nama_file)
        plt.close()
        QMessageBox.information(None, 'Information', f'Grafik berhasil disimpan ke {nama_file}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('icon.png'))
    window = InputDataPemeriksaan()
    window.show()
    sys.exit(app.exec_())