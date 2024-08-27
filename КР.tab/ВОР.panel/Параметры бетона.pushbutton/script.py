# -*- coding: utf-8 -*-
import os
import clr
import datetime
from System.Collections.Generic import *
import re
import inspect

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


class ReportItemForType:
    def __init__(self, type_name, value_b, value_f, value_w, inst_report_items):
        self.type_name = type_name  # имя типоразмера из Revit
        self.value_b = value_b  # значение марки бетона B
        self.value_f = value_f  # значение марки бетона F
        self.value_w = value_w  # значение марки бетона W
        self.inst_report_items = inst_report_items  # отчеты по экземплярам типоразмера из Revit


class ReportItemForInst:
    def __init__(self, inst_id, material_type_value):
        self.inst_id = inst_id  # id экземпляра
        self.material_type_value = material_type_value  # значение "ФОП_ТИП_Тип материала"


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

    def write_values_in_type(self, report_for_insts):
        self.elem_type.GetParam(mark_B_param_name).Set(self.value_b)
        self.elem_type.GetParam(mark_F_param_name).Set(self.value_f)
        self.elem_type.GetParam(mark_W_param_name).Set(self.value_w)
        """
        try:
            self.elem_type.GetParam(mark_B_param_name).Set(self.value_b)
            self.elem_type.GetParam(mark_F_param_name).Set(self.value_f)
            self.elem_type.GetParam(mark_W_param_name).Set(self.value_w)

            return ReportItemForType(
                self.elem_type_name,
                str(self.value_b),
                str(self.value_f),
                str(self.value_w),
                report_for_insts
            )
        except:
            self.has_errors = True
            return ReportItemForType(
                self.elem_type_name,
                "Ошибка",
                "Ошибка",
                "Ошибка",
                [])
"""
    def write_values_in_instance(self):
        report_for_insts = []
        try:
            for elem in self.elems_list:
                value = "B" + self.get_simple_str_value(self.value_b)
                elem.GetParam(material_type_param_name).Set(value)

                report_for_insts.append(ReportItemForInst(elem.Id, value))
        except:
            self.has_errors = True
        return report_for_insts

    def get_simple_str_value(self, value):
        if value % 1 == 0:
            return str(int(value))
        else:
            return str(value)

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
    errors = []
    # При выборке нельзя использовать фильтрацию по категориям, поэтому в elems много
    # элементов не нужных типов, отсеиваем их поиском параметра на экземпляре + по имени
    for elem in elems:
        try:
            if "(ЖБ" not in elem.Name and "(Б" not in elem.Name:
                continue
            elem.GetParam(material_type_param_name)
            temp.append(elem)
        except:
            errors.append(elem)

    if len(temp) == 0:
        output = script.output.get_output()
        output.close()
        alert("На активном виде не найдено ни одного элемента", exitscript=True)
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
    report = []
    for elem_type in elem_types:
        report_for_insts = elem_type.write_values_in_instance()

        #report_part = elem_type.write_values_in_type(report_for_insts)
        elem_type.write_values_in_type(report_for_insts)

        #report.append(report_part)
    return report


def get_report(type_report_list):
    report = []
    output = script.output.get_output()
    for type_report in type_report_list:
        report_part = []
        for inst_report in type_report.inst_report_items:
            report_item = ["", "", "", "", "", ""]
            report_item[4] = output.linkify(inst_report.inst_id)
            report_item[5] = inst_report.material_type_value

            report_part.append(report_item)

        first_report_part = report_part[0]
        first_report_part[0] = type_report.type_name
        first_report_part[1] = str(type_report.value_b)
        first_report_part[2] = str(type_report.value_f)
        first_report_part[3] = str(type_report.value_w)

        report = report + report_part
    return report

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
    print("Найдено элементов для работы: " + str(len(elements)))
    revit_elem_types = sort_elements(elements)
    print("Найдено типоразмеров: " + str(len(revit_elem_types)))
    analyze_element_types(revit_elem_types)



    print("Выполняю запись...")
    with revit.Transaction("BIM: Заполнение параметров бетона"):
        type_report_list = write_values(revit_elem_types)


    """
    report = get_report(type_report_list)
    output = script.output.get_output()
    output.print_table(table_data=report,
                       title="Отчет работы плагина",
                       columns=[
                           "Имя типоразмера",
                           "Марка B",
                           "Марка F",
                           "Марка W",
                           "ID экземпляра",
                           "Тип материала"]
                       )

"""
script_execute()
