from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QErrorMessage, QMessageBox
from interface import Ui_MainWindow
from pathlib import Path
import sys
import pandas as pd
from openpyxl import load_workbook
from webster import compute_webster
import os
import shutil
from create_sigs import start_creating_sigs
from src.utils import *
import warnings
import logging
from tqdm import tqdm

LOGGER = logging.getLogger(__name__)
LOGGER.setLevel(logging.DEBUG)
f = logging.Formatter("%(asctime)s-%(levelname)s: %(message)s")

class WebsterWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.open_pushButton.clicked.connect(self.subarea_open)
        self.ui.get_pushButton.clicked.connect(self.get_intersections)
        self.ui.start_pushButton.clicked.connect(self.start)
        self.ui.datos_pushButton.clicked.connect(self.create_datos)
        self.ui.sigs_pushButton.clicked.connect(self.multiply_sigs)

    def subarea_open(self):
        self.subarea_directory = QFileDialog.getExistingDirectory(self, "Open Directory")
        if self.subarea_directory:
            self.ui.subarea_lineEdit.setText(self.subarea_directory)
            logPath = os.path.join(self.subarea_directory, "logs")
            if not os.path.exists(logPath):
                os.mkdir(logPath)

            fh = logging.FileHandler(os.path.join(logPath, "Distro.log"), mode='w')
            fh.setFormatter(f)
            LOGGER.addHandler(fh)

    def get_intersections(self):
        try:
            error_message = QErrorMessage(self)
            self.listCodes = get_codes(self.subarea_directory, error_message)
        except AttributeError as e:
            error_message = QErrorMessage(self)
            return error_message.showMessage("There is no subarea folder found")
        except Exception as e:
            error_message = QErrorMessage(self)
            return error_message.showMessage(str(e))
        
        codesText = ", ".join(self.listCodes)
        self.ui.get_lineEdit.setText(codesText)

    def start(self) -> None:      
        selectedCode = self.ui.selected_lineEdit.text()

        if not selectedCode:
            error_message = QErrorMessage(self)
            return error_message.showMessage("There is no code selected")
        
        #Get vehicle path:
        pathParts = self.subarea_directory.split("/")
        projectParts = pathParts[:-2]
        subareaName = pathParts[-1]
        projectParts = "\\".join(projectParts)

        subareaFolder = os.path.join(
            projectParts,
            "7. Informacion de Campo",
            subareaName,
        )

        excel_by_agent = {
            "Vehicular": {
                "Tipico": None,
                "Atipico": None,
            },
            "Peatonal": {
                "Tipico": None,
                "Atipico": None,
            }
        }
        
        excel_by_agent, intervals = get_dict_by_agent(subareaFolder, excel_by_agent, selectedCode)
        
        #Lectura del excel de datos:
        path_datos = os.path.join(
            self.subarea_directory,
            f"DATOS_{selectedCode}.xlsx", #TODO: Cambiar a uno con el nombre del código de la subarea.
        )

        try:
            dfTurns = pd.read_excel(path_datos, sheet_name="DATA", header=0, usecols="A:G", nrows=51, skiprows=27).dropna()
            dfTurns['Fase'] = dfTurns['Fase'].apply(process_elem)
            dfLanes = pd.read_excel(path_datos, sheet_name="DATA", header=0, usecols="I:L", nrows=51, skiprows=27).dropna()
            dfLanes["Destino.1"] = pd.to_numeric(dfLanes["Destino.1"], errors="coerce")
            dfLanes["Destino.1"] = dfLanes["Destino.1"].astype("Int64")
            dfLanes["Origen.1"] = dfLanes["Origen.1"].astype("Int64")
            dfPhases = pd.read_excel(path_datos, sheet_name="DATA", header=0, usecols="A:E", nrows=11).dropna()
        except Exception as inst:
            error_message = QErrorMessage(self)
            return error_message.showMessage("No se encontro el archivo DATOS.xlsx")
        
        #rr_time_id = df2.loc[df2['Todo Rojo'].idxmax()]['Caso']
        #min_green_id = rr_time_id
        #Abriendo template de webster
        path_template = r".\tools\WEBSTER.xlsx"

        try:
            wb_WEBSTER = load_workbook(path_template, read_only=False, data_only=False)
        except Exception as inst:
            error_message = QErrorMessage(self)
            return error_message.showMessage("No se encontro el archivo de template WEBSTER.xlsx")

        wbVehicleTipico = load_workbook(
            excel_by_agent["Vehicular"]["Tipico"],
            read_only=True,
            data_only=True,
        )

        wbVehicleAtipico = load_workbook(
            excel_by_agent["Vehicular"]["Atipico"],
            read_only=True,
            data_only=True,
        )
        
        wbPedestrianTipico = load_workbook(
            excel_by_agent["Peatonal"]["Tipico"],
            read_only=True,
            data_only=True,
        )

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

            print(f"Cargando: {i+1}/13")
            try:
                compute_webster(
                    wbVehicleTipico,
                    wbVehicleAtipico,
                    wbPedestrianTipico,
                    dfTurns, dfLanes, dfPhases, #Dataframes enviados
                    interval, factor, wb_WEBSTER, i, #Datos según cada escenario
                    LOGGER,
                    )
            except Exception as inst:
                error_message = QErrorMessage(self)
                return error_message.showMessage("Error en calcular Webster")

        wb_WEBSTER.save(
            os.path.join(
                self.subarea_directory,
                f"WEBSTER_{selectedCode.upper()}.xlsx", #TODO: Change after fix compute_webster
                )
            )
        wbVehicleTipico.close()
        wbVehicleAtipico.close() 
        wbPedestrianTipico.close()
        wb_WEBSTER.close()

        self.ui.label.setText("Done!")

        #TODO: Falta ahora corregir la función de cálculo de Webster, ya se esta enviando la información.

    def create_datos(self) -> None:
        origin_route = r".\tools\DATOS.xlsx"
        
        if self.ui.get_lineEdit.text == "":
            error_message = QMessageBox(self)
            return error_message.setText("Selected code is empty. Choose a code first!")
        
        destiny_route = os.path.join(
            self.subarea_directory,
            f"DATOS_{self.ui.selected_lineEdit.text().upper()}.xlsx",
        )

        shutil.copy2(origin_route, destiny_route)
        info_message = QMessageBox(self)
        info_message.setIcon(QMessageBox.Information)
        info_message.setWindowTitle("Info")
        info_message.setText(f"Se ha creado el archivo DATOS_{self.ui.selected_lineEdit.text().upper()}.xlsx")
        return info_message.show()

    def multiply_sigs(self) -> None:
        """ Create SIG from a unique sig file in ./Tipico/HPM folder. """
        code = os.path.split(self.vehicle_path)[1]
        code = code[:5]
        hpm_excel_path = Path(self.subarea_directory) / "Propuesto" / "Tipico" / "HPM" / f"{code}.sig"
        
        for scenario in ["HPMAD","HPN","HPT","HVM","HVMAD","HVN","HVT"]:
            destiny_route = Path(self.subarea_directory) / "Propuesto" / "Tipico" / scenario / f"{code}.sig"
            try:
                shutil.copy2(hpm_excel_path, destiny_route)
            except Exception as inst:
                error_message = QErrorMessage(self)
                error_message.showMessage("No se encontro el archivo HPM.sig")

        for scenario in ["HPM","HPN","HPT","HVMAD","HVN"]:
            destiny_route = Path(self.subarea_directory) / "Propuesto" / "Atipico" / scenario / f"{code}.sig"
            try:
                shutil.copy2(hpm_excel_path, destiny_route)
            except Exception as inst:
                error_message = QErrorMessage(self)
                error_message.showMessage("No se encontro el archivo HPM.sig")

        self.ui.label.setText("Copied!")

    def create_sig_webster(self) -> None:
        """ Modify all sigs according to Webster times. """
        code = os.path.split(self.vehicle_path)[1][:5]
        subarea_folder = Path(self.subarea_directory)

        subarea_content = os.listdir(subarea_folder)
        for file in subarea_content:
            if f"WEBSTER_{code}.xlsx" == file:
                try:
                    start_creating_sigs(
                        webs_xlsx = subarea_folder / file,
                        code_int = code,
                    )
                except Exception as inst:
                    error_message = QErrorMessage(self)
                    return error_message.showMessage(inst)
                break

def main():
    warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
    app = QApplication(sys.argv)
    app.processEvents()
    window = WebsterWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()