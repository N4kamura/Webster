import xml.etree.ElementTree as ET
from pathlib import Path
from openpyxl import load_workbook
import os

NAMES_TIP = [
    "HPMAD",
    "HVMAD",
    "HPM",
    "HVM",
    "HPT",
    "HVT",
    "HPN",
    "HVN",
]

NAMES_ATI= [
    "HVMAD",
    "HPM",
    "HPT",
    "HPN",
    "HVN",
]

def _change_sig(
        list_green,
        sig_path,
        ) -> None:
    #################################
    # Modifying sigs for each phase #
    #################################

    tree = ET.parse(sig_path)
    sc_tag = tree.getroot()
    interstages = sc_tag.find("./stageProgs/stageProg/interstages")
    for interstage, green in zip(interstages.findall("./interstage"), list_green):
        interstage.attrib["begin"] = str(green*1000)
    
    ET.indent(tree, "    ")
    tree.write(sig_path, encoding = "utf-8", xml_declaration = True)

def _get_greens(
        webs_xlsx: str | Path, #Route of excel.
        ) -> dict:
    """ Modify sigs for each scenario. """
    ##############################
    # Computing phases and times #
    ##############################

    wb = load_workbook(webs_xlsx, read_only=True, data_only=True)
    ws = wb['WEBSTER']

    TIMES = [[elem.value for elem in row] for row in ws[slice("U2","AI14")]]
    wb.close()

    programs_dict = {}
    for index, row in enumerate(TIMES):
        program = []
        for i in range(len(row)):
            if i%3 == 0:
                program.append(row[i:i+3])
        programs_dict[index+1] = program
        
    program_0 = programs_dict[1]
    for i, phases in enumerate(program_0):
        if sum(phases) == 0:
            no_phases = i
            break

    ###################################
    # Computing greens times for sigs #
    ###################################

    begin_dict = {}
    for key, value in programs_dict.items():
        row = value[:no_phases]
        greens = []
        aux = 0
        for i in range(len(row)):
            if i==0:
                greens.append(row[i][0])
            else:
                aux = aux + sum(row[i-1])
                greens.append(aux+row[i][0])

        begin_dict[key] = greens

    return begin_dict

def start_creating_sigs(
        webs_xlsx: str | Path, #Route of excel.
        code_int: str, #Code of intersectio
        ) -> None:

    begin_dict = _get_greens(webs_xlsx)

    #######################
    # Finding sigs folder #
    #######################

    folder = os.path.dirname(webs_xlsx)
    propuesto_folder = Path(folder) / "Propuesto"
    for tipicidad in ["Tipico", "Atipico"]:
        list_files = os.listdir(propuesto_folder / tipicidad)
        list_files = [file for file in list_files if not file.endswith(".ini")]
        for scenario in list_files:
            sig_files = os.listdir(propuesto_folder / tipicidad / scenario)
            sig_files = [file for file in sig_files if file.endswith(".sig")]
            for sig in sig_files:
                if code_int == sig[:-4]:
                    sig_path = propuesto_folder / tipicidad / scenario / sig
                    if tipicidad == "Tipico":
                        for i, NAME in enumerate(NAMES_TIP):
                            if NAME == scenario:
                                list_green = begin_dict[i+1]

                    elif tipicidad == "Atipico":
                        for i, NAME in enumerate(NAMES_ATI):
                            if NAME == scenario:
                                list_green = begin_dict[i+1]
                    _change_sig(list_green, sig_path)
                    break
