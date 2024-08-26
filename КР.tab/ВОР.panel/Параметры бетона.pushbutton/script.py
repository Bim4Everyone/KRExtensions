# -*- coding: utf-8 -*-
import os
import clr
import datetime
from System.Collections.Generic import *
import re

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


# Класс для хранения информации по типу Revit
class RevitElementType:
    def __init__(self, elem_type_name, elems_list):
        self.elem_type_name = elem_type_name  # имя типоразмера из Revit
        self.elems_list = elems_list  # экземпляры типоразмера из Revit
        type_id = elems_list[0].GetTypeId()
        self.elem_type = doc.GetElement(type_id)  # типоразмер из Revit
        self.value_b = 0.0  # значение для "обр_ФОП_Марка бетона B"
        self.value_f = 0.0  # значение для "обр_ФОП_Марка бетона F"
        self.value_w = 0.0  # значение для "обр_ФОП_Марка бетона W"
        self.has_errors = False  # метка, что при получении значений возникли ошибки

    def analyze_element_type_name(self):
        pat_B = 'B'
        pat_F = 'F'
        pat_W = 'W'
        base_pattern = '[0-9,.]*'

        value_b = "0"
        value_f = "0"
        value_w = "0"

        if "B" in self.elem_type_name:
            try:
                search_b = re.findall(pat_B + base_pattern, self.elem_type_name)[0]  # B30
                value_b = re.findall(base_pattern, search_b)[1]  # 30
                value_b = value_b.replace(",", ".")
                self.value_b = float(value_b)
            except:
                self.has_errors = True

        if "F" in self.elem_type_name:
            try:
                search_f = re.findall(pat_F + base_pattern, self.elem_type_name)[0]  # F150
                value_f = re.findall(base_pattern, search_f)[1]  # 150
                value_f = value_f.replace(",", ".")
                self.value_f = float(value_f)
            except:
                self.has_errors = True

        if "W" in self.elem_type_name:
            try:
                search_w = re.findall(pat_W + base_pattern, self.elem_type_name)[0]  # W6
                value_w = re.findall(base_pattern, search_w)[1]  # 6
                value_w = value_w.replace(",", ".")
                self.value_w = float(value_w)
            except:
                self.has_errors = True

    def write_values(self):
        try:
            print(self.elem_type)
            self.elem_type.GetParam(mark_B_param_name).Set(self.value_b)
            self.elem_type.GetParam(mark_F_param_name).Set(self.value_f)
            self.elem_type.GetParam(mark_W_param_name).Set(self.value_w)
        except:
            self.has_errors = True


def get_elements():
    elems = (FilteredElementCollector(doc, doc.ActiveView.Id)
             .WhereElementIsNotElementType()
             .ToElements())

    if len(elems) == 0:
        output = script.output.get_output()
        output.close()
        alert("На активном виде не найдено ни одного элемента", exitscript=True)
    return elems


def filter_elements(elems):
    temp = []
    for elem in elems:
        if "(" not in elem.Name and ")" not in elem.Name:
            continue
        print(elem.Name)
        temp.append(elem)
    return temp


def sort_elements(elems):
    elems_dict = {}
    for elem in elems:
        name = elem.Name
        if elems_dict.has_key(name):
            elems_dict[name].append(elem)
        else:
            elems_dict[name] = [elem]

    revit_elem_types = []
    for key in elems_dict.keys():
        revit_elem_types.append(RevitElementType(key, elems_dict[key]))

    return revit_elem_types


def analyze_element_types(elem_types):
    for elem_type in elem_types:
        elem_type.analyze_element_type_name()


def write_values(elem_types):
    for elem_type in elem_types:
        elem_type.write_values()


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
    elements = filter_elements(elements)
    revit_elem_types = sort_elements(elements)

    analyze_element_types(revit_elem_types)

    with revit.Transaction("BIM: Заполнение параметров бетона"):
        write_values(revit_elem_types)


script_execute()
