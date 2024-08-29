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


class ReportItem:
    def __init__(self, type_name, value_b, value_f, value_w, material_type, count_of_insts):
        self.type_name = type_name  # имя типоразмера из Revit
        self.value_b = value_b  # значение марки бетона B
        self.value_f = value_f  # значение марки бетона F
        self.value_w = value_w  # значение марки бетона W
        self.material_type = material_type  # значение типа материала бетона
        self.count_of_insts = count_of_insts  # кол-во экземпляров типоразмера из Revit


# Класс для хранения информации по типоразмеру из Revit
class RevitElementType:
    def __init__(self, elem_type_name, elems_list):
        self.elem_type_name = elem_type_name  # имя типоразмера из Revit
        self.elems_list = elems_list  # экземпляры типоразмера из Revit

        type_id = elems_list[0].GetTypeId()
        self.elem_type = doc.GetElement(type_id)  # типоразмер из Revit

        # Задаем значения по умолчанию
        self.value_b = 0.0  # значение для "обр_ФОП_Марка бетона B"
        self.value_f = 0.0  # значение для "обр_ФОП_Марка бетона F"
        self.value_w = 0.0  # значение для "обр_ФОП_Марка бетона W"
        self.material_type = "B0"  # значение для "ФОП_ТИП_Тип материала"
        self.has_errors = False  # метка, что при получении значений возникли ошибки

    def analyze_element_type_name(self):
        pat_B = r'[BВ]'
        pat_F = 'F'
        pat_W = 'W'
        base_pattern = r'[0-9,.]*'

        value_b = "0"
        value_f = "0"
        value_w = "0"

        # Полное имя "ВН_Перекрытие-200 (ЖБ В25 F150 W4)"
        # Отбираем "(ЖБ В25 F150 W4)"
        material_part_of_name = re.findall('\(.*\)', self.elem_type_name)
        if not material_part_of_name:
            self.has_errors = True
            return

        material_part_of_name = material_part_of_name[0]

        # Букву "B" можно написать на русском и на английском
        if re.findall('[BВ]', material_part_of_name):
            try:
                search_b = re.findall(pat_B + base_pattern, material_part_of_name)[0]  # B30
                value_b = re.findall(base_pattern, search_b)[1]  # 30 из ['', '30', '']
                value_b = value_b.replace(",", ".")  # 30 или 7.5
                self.material_type = "B" + value_b  # B30 или B7.5

                value_b = float(value_b)  # 30.0 или 7.5
                self.value_b = value_b
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

    def write_values_in_params(self):
        try:
            self.elem_type.GetParam(mark_B_param_name).Set(self.value_b)
            self.elem_type.GetParam(mark_F_param_name).Set(self.value_f)
            self.elem_type.GetParam(mark_W_param_name).Set(self.value_w)
            self.write_values_in_instance()

            return ReportItem(
                self.elem_type_name,
                str(self.value_b),
                str(self.value_f),
                str(self.value_w),
                self.material_type,
                str(len(self.elems_list))
            )
        except:
            self.has_errors = True
            return ReportItem(
                self.elem_type_name,
                "Ошибка",
                "Ошибка",
                "Ошибка",
                "Ошибка",
                str(len(self.elems_list))
            )

    def write_values_in_instance(self):
        try:
            for elem in self.elems_list:
                elem.GetParam(material_type_param_name).Set(self.material_type)
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
    errors = []
    # При выборке нельзя использовать фильтрацию по категориям (по ТЗ), поэтому в elems много
    # элементов не нужных типов, отсеиваем их поиском параметра на экземпляре + по имени
    for elem in elems:
        try:
            if not re.findall("(\(ЖБ|\(Б)( B| В)", elem.Name):
                continue
            elem.GetParam(material_type_param_name)
            # Отбираем элементы, которые прошли фильтрацию
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
        # Распределяем элементы группируя по имени типа
        if elems_dict.has_key(name):
            elems_dict[name].append(elem)
        else:
            elems_dict[name] = [elem]
    # Теперь ключи - имена типов, значения - списки экзепляров элементов

    revit_elem_types = []
    for key in elems_dict.keys():
        revit_elem_types.append(RevitElementType(key, elems_dict[key]))
    return revit_elem_types


# Проводим анализ каждого типа: забираем из имени значения B,W,F, заполняем поля класса-оболочки
def analyze_element_types(elem_types):
    for elem_type in elem_types:
        elem_type.analyze_element_type_name()


# Записываем значения в параметры бетона на типе и в Тип материала на экземпляре
# на основе информации из имени типа
def write_values(elem_types):
    report = []
    for elem_type in elem_types:
        report_part = elem_type.write_values_in_params()
        report.append(report_part)
    return report


# Преобразуем данные в вид, подходящий для таблиц pyRevit
def get_report(type_report_list):
    report = []
    for type_report in type_report_list:
        report_item = ["", "", "", "", "", ""]
        report_item[0] = type_report.type_name
        report_item[1] = type_report.value_b
        report_item[2] = type_report.value_f
        report_item[3] = type_report.value_w
        report_item[4] = type_report.material_type
        report_item[5] = type_report.count_of_insts

        report.append(report_item)
    return report


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    print("Здравствуйте! Данный плагин предназначен для записи значений в параметры бетона:")
    print("- обр_ФОП_Марка бетона B")
    print("- обр_ФОП_Марка бетона F")
    print("- обр_ФОП_Марка бетона W")
    print("- ФОП_ТИП_Тип материала")

    print("Будут отобраны элементы только с корректными наименованиями:")
    print("- \"НН_Перекрытие-240 (ЖБ B30 F150 W6)\"")
    print("- \"Ф_Подготовка-70 (Б B7.5)\"")

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

    report = get_report(type_report_list)
    '''
    # При печати таблицы встречается ошибка, когда таблица по неустановленной причине печаться не хочет
    # Решается добавлением в коллекцию еще одной строки
    # Чтобы не отвлекать пользователя в доп строке содержится один невидимый символ (это не пробел)
    '''
    report.append(["⠀", "", "", "", "", ""])

    output = script.output.get_output()
    output.print_table(table_data=report[:],
                       title="Отчет работы плагина",
                       columns=[
                           "Имя типоразмера",
                           "Марка B",
                           "Марка F",
                           "Марка W",
                           "Тип материала",
                           "Кол-во"]
                       )


script_execute()
