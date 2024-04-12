from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QProgressBar, QErrorMessage
from interface import Ui_MainWindow
from pathlib import Path
import sys
import pandas as pd
from openpyxl import load_workbook
from webster import compute_webster
import os
import re

class WebsterWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.pushButton.clicked.connect(self.vehicle_open)
        self.ui.pushButton_2.clicked.connect(self.pedestrian_open)
        self.ui.pushButton_3.clicked.connect(self.subarea_open)
        self.ui.pushButton_4.clicked.connect(self.start)

    def vehicle_open(self):
        self.vehicle_path = QFileDialog.getOpenFileName(self, "Open File", "", "Excel Files (*.xlsm)")[0]
        if self.vehicle_path: self.ui.lineEdit.setText(self.vehicle_path)

    def pedestrian_open(self):
        self.pedestrian_path = QFileDialog.getOpenFileName(self, "Open File", "", "Excel Files (*.xlsm)")[0]
        if self.pedestrian_path: self.ui.lineEdit_2.setText(self.pedestrian_path)

    def subarea_open(self):
        self.subarea_directory = QFileDialog.getExistingDirectory(self, "Open Directory")
        if self.subarea_directory: self.ui.lineEdit_3.setText(self.subarea_directory)

    def start(self) -> None:
        try:
            vehicle_path = Path(self.vehicle_path)
            pedestrian_path = Path(self.pedestrian_path)
            subarea_folder = Path(self.subarea_directory)
        except Exception as inst:
            error_message = QErrorMessage(self)
            return error_message.showMessage("Primero debe abrir todos los archivos")
        
        path_datos = subarea_folder / "DATOS.xlsx"

        try:
            df = pd.read_excel(path_datos, header=0, usecols="A:G", nrows=11)
        except Exception as inst:
            error_message = QErrorMessage(self)
            return error_message.showMessage("No se encontro el archivo DATOS.xlsx")

        df.index = df.iloc[:,0].astype(int)
        df =df.iloc[:,1:]
        mapping = {'SI': True, 'NO': False}
        df['Protegido'] = df['Protegido'].replace(mapping).infer_objects(copy=False)

        try:
            df2 = pd.read_excel(path_datos, header=0, usecols="I:L", nrows=11).dropna()
        except Exception as inst:
            error_message = QErrorMessage(self)
            return error_message.showMessage("No se encontro el archivo DATOS.xlsx")
        
        rr_time_id = df2.loc[df2['Todo Rojo'].idxmax()]['Caso']
        min_green_id = rr_time_id

        intervals = []
        wb = load_workbook(vehicle_path, read_only=True, data_only=True)
        ws = wb['Histograma']
        for j in range(3):
            peak_hour = ws.cell(18+j*7,3).value
            hour = int((int(peak_hour[:2]) + int(peak_hour[3:5])/60)*4)
            intervals.append(slice(hour, hour+4))
        
        wb.close()

        path_template = r".\tools\WEBSTER.xlsx"

        code = os.path.split(vehicle_path)[1]
        code = code[:5]

        vehicular_folder = vehicle_path
        for _ in range(2):
            vehicular_folder = os.path.split(vehicular_folder)[0]
        
        vehicular_folder = Path(vehicular_folder)
        atipico_folder = vehicular_folder / "Atipico"
        atipico_excels = os.listdir(atipico_folder)
        pattern = r"([A-Z]+-[0-9]+)"
        for atipico_excel in atipico_excels:
            coincidence = re.search(pattern, atipico_excel)
            if coincidence:
                atipico_excel_path = atipico_folder / atipico_excel
                break

        wb = load_workbook(atipico_excel_path, read_only=True, data_only=True)
        ws = wb['Histograma']
        for j in range(3):
            peak_hour = ws.cell(18+j*7,3).value
            hour = int((int(peak_hour[:2]) + int(peak_hour[3:5])/60)*4)
            intervals.append(slice(hour, hour+4))
        
        wb.close()

        try:
            wb_WEBSTER = load_workbook(path_template, read_only=False, data_only=False)
        except Exception as inst:
            error_message = QErrorMessage(self)
            return error_message.showMessage("No se encontro el archivo de template WEBSTER.xlsx")

        for i in range(13):
            #Factores
            if i+1 == 1 or i+1 == 9:
                factor = 0.30
            elif i+1 == 2 or i+1 == 8 or i+1 == 13:
                factor = 0.50
            else:
                factor = 1.00

            #Horas puntas para los intervalos
            if 1 <= i+1 <= 3:
                interval = intervals[0]
            elif i+1 == 4:
                interval = slice(8*4, 9*4)
            elif i+1 == 5:
                interval = intervals[1]
            elif i+1 == 6:
                interval = slice(14*4, 15*4)
            elif 7 <= i+1 <= 8:
                interval = intervals[2]
            elif 9 <= i+1 <= 10:
                interval = intervals[3]
            elif i+1 == 11:
                interval = intervals[4]
            elif 12 <= i+1:
                interval = intervals[5]

            try:
                compute_webster([vehicle_path, atipico_excel_path], pedestrian_path, min_green_id, rr_time_id, interval, df, factor, wb_WEBSTER, i)
            except Exception as inst:
                error_message = QErrorMessage(self)
                raise inst
                return error_message.showMessage("Error en calcular Webster")
            self.ui.progressBar.setValue(i)

        wb_WEBSTER.save(subarea_folder / f"WEBSTER_{code}.xlsx")
        wb_WEBSTER.close()

        self.ui.label.setText("Done!")

def main():
    app = QApplication(sys.argv)
    app.processEvents()
    window = WebsterWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()