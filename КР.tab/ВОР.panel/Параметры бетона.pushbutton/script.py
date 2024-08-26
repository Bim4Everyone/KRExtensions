# -*- coding: utf-8 -*-
import os
import clr
import datetime
from System.Collections.Generic import *

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from System.Runtime.InteropServices import Marshal

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep_libs.bim4everyone import *

import pyevent
from pyrevit import EXEC_PARAMS, revit
from pyrevit.forms import *
from pyrevit import script

import Autodesk.Revit.DB
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application

mark_B_param_name = "обр_ФОП_Марка бетона B"
mark_F_param_name = "обр_ФОП_Марка бетона F"
mark_W_param_name = "обр_ФОП_Марка бетона W"
material_type_param_name = "ФОП_ТИП_Тип материала"


# Класс для хранения инфы по видам работ из Excel
class RevitElement:
    def __init__(self, elem):
        self.elem = elem  # экземпляр элемента из Revit
        self.chapter = chapter  # наименование главы
        self.title_of_work = title_of_work  # наименование работы
        self.unit_of_measurement = unit_of_measurement  # единица измерения

    def get_data_from_name(self):
        print("Hello")


def get_elements():
    elems = (FilteredElementCollector(doc, doc.ActiveView.Id)
             .WhereElementIsNotElementType()
             .ToElements())

    if len(elems) == 0:
        output = script.output.get_output()
        output.close()
        alert("На активном виде не найдено ни одного элемента", exitscript=True)
    return elems

def sort_elements(elems):
    elems_dict = {}
    for elem in elems:
        if "B" not in elem.Name:
            continue
        name = elem.Name
        if elems_dict.has_key(name):
            elems_dict[name].append(elem)
        else:
            elems_dict[name] = [elem]
    return elems_dict


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    print("Здравствуйте! Данный плагин предназначен для записи значений в параметры бетона:")
    print("- обр_ФОП_Марка бетона B")
    print("- обр_ФОП_Марка бетона F")
    print("- обр_ФОП_Марка бетона W")
    print("- ФОП_ТИП_Тип материала")



    print("Собираю элементы на активном виде...")
    elements = get_elements()
    elems_dict = sort_elements(elements)

    for k in elems_dict.keys():
        print(str(k))

    print(str(len(elems_dict.keys())))

script_execute()