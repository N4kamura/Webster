import xml.etree.ElementTree as ET
import os
from openpyxl import load_workbook
import re
import numpy as np
import pandas as pd
import win32com.client as win32
from dataclasses import dataclass
from pathlib import Path

def process_elem(item: str | int) -> list:
    if isinstance(item, str):
        return list(map(int, item.split(',')))
    else:
        return [item]

def get_codes(subareaPath, error_message):
    listContent = os.listdir(subareaPath)

    #Filters to obtain .inpx file (skeleton)
    listContent = [file for file in listContent if file.endswith(".inpx") and "(SA)" in file]

    #Check if only one .inpx file is found
    if len(listContent) > 1:
        return error_message.showMessage("More than one .inpx file found")

    vissimFile = listContent[0]
    vissimPath = os.path.join(subareaPath, vissimFile)

    #Obtaining codes from inpx file (skeleton)
    tree = ET.parse(vissimPath)
    network_tag = tree.getroot()
    listCodes = []
    for node_tag in network_tag.findall("./nodes/node"):
        for uda_tag in node_tag.findall("./uda"):
            code = uda_tag.get("value")
            listCodes.append(code)

    tree = None

    return listCodes

def _get_interval_from_excel(excelPath) -> slice:
    wb = load_workbook(excelPath, read_only=True, data_only=True)
    ws = wb['Histograma']
    intervals = []
    for j in range(3):
        peakHour = ws.cell(18+j*7,3).value
        hour = int((int(peakHour[:2]) + int(peakHour[3:5])/60)*4)
        intervals.append(slice(hour, hour + 4))
    
    wb.close()
    
    return intervals

def get_dict_by_agent(subareaFolder: str, excel_by_agent: dict, selectedCode: str) -> list[dict, list[slice]]:
    pattern = r"([A-Z]+-[0-9]+)"
    intervals = []
    for agentName in ["Vehicular", "Peatonal"]:
        for tipicidad in ["Tipico", "Atipico"]:
            agentFolder = os.path.join(
                subareaFolder,
                agentName,
                tipicidad,
            )

            agentContentFolder = os.listdir(agentFolder)
            for agentFile in agentContentFolder:
                if re.search(pattern, agentFile):
                    codeFound = re.search(pattern, agentFile).group(1)
                    if selectedCode == codeFound:
                        agentFilePath = os.path.join(
                            agentFolder,
                            agentFile,
                        )
                        if agentName == "Vehicular":
                            intervalObtained = _get_interval_from_excel(agentFilePath)
                            intervals.extend(intervalObtained)
                        excel_by_agent[agentName][tipicidad] = agentFilePath
                        break

    return excel_by_agent, intervals

def flows(vehicle_types: list,
          wb: load_workbook,
          FACTOR: float,
          interval: slice,
          ) -> np.array:
    origin_slices = [
        slice("E12", "E21"),
        slice("K12", "K21"),
        slice("E24", "E33"),
        slice("K24", "K33"),
    ]

    destiny_slices = [
        slice("F12", "F21"),
        slice("L12", "L21"),
        slice("F24", "F33"),
        slice("L24", "L33"),
    ]

    giro_slices = [
        slice("G12", "G21"),
        slice("M12", "M21"),
        slice("G24", "G33"),
        slice("M24", "M33"),
    ]

    num_giros = [0,0,0,0]
    num_veh_classes = len(vehicle_types)

    ws = wb['Inicio']
    for j, turn in enumerate(giro_slices):
        aux = []
        for row in ws[turn]:
            aux.append(row[0].value)
        try:
            quant = aux.index(None)
        except ValueError:
            quant = len(aux)
        num_giros[j] = quant

    hojas = ["N", "S", "E", "O"]

    list_destination    = []
    list_origin         = []
    list_vr_name        = []

    list_flow = []
    for i_giro in range(len(hojas)):
        ws = wb["Inicio"]
        slice_origin = origin_slices[i_giro]
        slice_destiny = destiny_slices[i_giro]
        slice_giros = giro_slices[i_giro]
        num_giro_i = num_giros[i_giro]

        list_origin.extend(
            [row[0].value for row in ws[slice_origin]][:num_giro_i]
        )

        list_destination.extend(
            [row[0].value for row in ws[slice_destiny]][:num_giro_i]
        )

        list_vr_name.extend(
            [row[0].value for row in ws[slice_giros]][:num_giro_i]
        )

        ws = wb[hojas[i_giro]]

        list_A = [[cell.value for cell in row] for row in ws["K16":"HB111"]]
        A = np.array(list_A, dtype="float")
        A[np.isnan(A)] = 0
        A = A*FACTOR

        list_flow.append(
            np.array(
                [
                    A[interval, (10*veh_type):(10*veh_type+num_giro_i)]
                    for veh_type in range(num_veh_classes)
                ]
            )
        )

        array_flow = np.concatenate(list_flow, axis=-1)

    return array_flow, list_origin, list_destination, list_vr_name

def pedestrian_flows(wb: load_workbook, interval: slice):
    ws = wb['Inicio']

    movimiento_slices = [
        slice("G13", "G22"),
        slice("K13", "K22"),
        slice("G25", "G34"),
        slice("K25", "K34"),
    ]

    types_ped_slices = [
        slice("W4", "W8"),
        slice("Y4", "Y8"),
        slice("AA4", "AA8"),
    ]

    num_ped_classes = 0
    for type_ped_slice in types_ped_slices:
        num_ped_classes += [row[0].value for row in ws[type_ped_slice]].index(None)

    num_giros = [0,0,0,0]
    for j, turn in enumerate(movimiento_slices):
        aux = []
        for row in ws[turn]:
            aux.append(row[0].value)
        try: quant = aux.index(None)
        except ValueError: quant = len(aux)
        num_giros[j] = quant

    list_movimientos = []
    list_ped = []

    for i_giro in range(len(num_giros)):
        ws = wb['Inicio']
        slice_movimiento = movimiento_slices[i_giro]
        num_giro_i = num_giros[i_giro]

        list_movimientos.extend([row[0].value for row in ws[slice_movimiento]][:num_giro_i])

        ws = wb['Data Peatonal']

        list_A = [[cell.value for cell in row] for row in ws["L20":"UY83"]]
        A = np.array(list_A, dtype="float")
        A[np.isnan(A)] = 0
        A = np.concatenate((np.zeros((24, A.shape[1])), A), axis = 0)

        list_ped.append(
            np.array(
                [
                    A[interval, (10*ped_type + 140*i_giro):(10*ped_type + num_giro_i + 140*i_giro)]
                    for ped_type in range(num_ped_classes)
                ]
            )
        )

        array_ped_flow = np.concatenate(list_ped, axis=-1)

        str_movements = [str(no) for no in list_movimientos]

    list_sums = []
    for ped_type in range(num_ped_classes):
        check_status = [False for _ in range(len(str_movements))]
        ped_flow = array_ped_flow[ped_type]
        list_sum_peds = []
        for i, giro in enumerate(str_movements):
            flow_cross = 0
            if check_status[i]: continue
            check_status[i] = True
            flow_cross += sum(ped_flow[:,i])
            for j, giro_inv in enumerate(str_movements):
                if giro[::-1] == giro_inv:
                    check_status[j] = True
                    flow_cross += sum(ped_flow[:,j])
                    break
            list_sum_peds.append(flow_cross)
        list_sums.append(list_sum_peds)

    sum_by_peds = [sum(x) for x in zip(*list_sums)]

    max_ped_flow = max(sum_by_peds)

    return max_ped_flow

def compute_flows(origin: int, dfTurns: pd.DataFrame, direction: int, dfFlows: pd.DataFrame, array_flow):
    flow = 0
    listTurn = dfTurns[(dfTurns["Origen"] == origin) & (dfTurns["Giro"] == direction)].index.tolist()
    for leftTurnIndex in listTurn:
        for veh_type in range(len(array_flow)):
            flow += sum(array_flow[veh_type][:,leftTurnIndex])
    if flow == None:
        dfFlows.at[origin, direction] = 0
    else:
        dfFlows.at[origin, direction] = flow

def duplicate_name_sheets(excelPath: str, listCodes: list, finalPath: str) -> None:
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
    except Exception as inst:
        excel.Application.Quit()
        raise inst

    try:
        #Starting the excel
        workBook = excel.Workbooks.Open(excelPath)
        mainSheet = workBook.Sheets('DATA') #I uniformized everything with this name

        #Create new sheets
        for code in listCodes:
            mainSheet.Copy(After=mainSheet)
            newSheet = workBook.Sheets[mainSheet.Index + 1]
            newSheet.Name = code

        #Delete sheets
        mainSheet.Delete()
        #workBook.Worksheets(2).Delete()
    except Exception as inst:
        excel.Application.Quit()
        raise inst

    #End with the excel
    finalPath = os.path.normpath(finalPath)

    try:
        workBook.SaveAs(finalPath)
        workBook.Close(SaveChanges = True)
    except Exception as inst:
        excel.Application.Quit()
        raise inst

def data2excel(subareaFolder: str, destiny_route: str) -> None:
    """ Create an excel with data from counting files. """
    #Find field files:
    pathDirectory = Path(subareaFolder)
    subareaName = pathDirectory.name
    projectPath = pathDirectory.parents[1] #Va alrevés
    fieldPath = projectPath / "7. Informacion de Campo" / subareaName / "Vehicular"
    typicalPath = fieldPath / "Tipico"
    excelVehicles = os.listdir(typicalPath)
    pattern = r"([A-Z]+-[0-9]+)"

    @dataclass
    class excelData:
        origins: list
        destinations: list
        names: list

    dictInfo = {}

    for excel in excelVehicles:
        excelPath = typicalPath / excel
        code = re.search(pattern, excel).group(1)
        wb = load_workbook(excelPath, read_only=True, data_only=True)
        ws = wb["Inicio"]
        slicesOrigin = [
            slice("E12", "E22"),
            slice("K12", "K22"),
            slice("E24", "E34"),
            slice("K24", "K34"),
        ]

        slicesDestination = [
            slice("F12", "F22"),
            slice("L12", "L22"),
            slice("F24", "F34"),
            slice("L24", "L34"),
        ]

        slicesNames = [
            slice("G12", "G22"),
            slice("M12", "M22"),
            slice("G24", "G34"),
            slice("M24", "M34"),
        ]
        
        listOrigins = []
        listDestinations = []
        listNames = []

        for sliceOrigin, sliceDestination, sliceName in zip(slicesOrigin, slicesDestination, slicesNames):
            content = [row[0].value for row in ws[sliceOrigin] if row[0].value is not None]
            listOrigins.extend(content)
            content = [row[0].value for row in ws[sliceDestination] if row[0].value is not None]
            listDestinations.extend(content)
            content = [row[0].value for row in ws[sliceName] if row[0].value is not None]
            listNames.extend(content)

        wb.close()

        data = excelData(
            origins = listOrigins,
            destinations = listDestinations,
            names = listNames
        )

        dictInfo[code] = data

    #Writing data in excel

    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(destiny_route)

    for nameSheet, dataList in dictInfo.items():
        ws = wb.Sheets[nameSheet]
        for i in range(len(dataList.origins)):
            ws.Cells(29+i, 1).Value = dataList.origins[i]
            ws.Cells(29+i, 2).Value = dataList.destinations[i]
            ws.Cells(29+i, 4).Value = dataList.names[i]

        originList = list(set(dataList.origins))
        
        for j, origin in enumerate(originList):
            ws.Cells(29+j, 10).Value = origin

    wb.Save()
    wb.Close()