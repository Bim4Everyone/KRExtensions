# -*- coding: utf-8 -*-
import Autodesk.Revit.DB
import clr
import datetime

from System.Collections.Generic import *

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

from System.Windows.Input import ICommand

import pyevent
from pyrevit import EXEC_PARAMS, revit
from pyrevit.forms import *
from pyrevit import script

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep_libs.bim4everyone import *

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from System.Runtime.InteropServices import Marshal

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application

CHAPTER_PARAMETER = "ФОП_МТР_Наименование главы"
WORK_TITLE_PARAMETER = "ФОП_МТР_Наименование работы"
UNIT_PARAMETER = "ФОП_МТР_Единица измерения"
CALCULATION_TYPE_PARAMETER = "ФОП_МТР_Тип подсчета"

CALCULATION_TYPE_DICT = {
    "м": 1,
    "м²": 2,
    "м³": 3,
    "шт.": 4}

report_no_work_code = []
report_classifier_code_not_found = []
report_edited = []
report_not_edited = []
report_errors = []


# Класс для хранения инфы по видам работ из Excel
class Work:
    def __init__(self, code, chapter, title_of_work, unit_of_measurement):
        self.code = code  # код работы
        self.chapter = chapter  # наименование главы
        self.title_of_work = title_of_work  # наименование работы
        self.unit_of_measurement = unit_of_measurement  # единица измерения


# Класс-оболочка для хранения информации
class RevitMaterial:
    def __init__(self, keynote, material, work):
        self.keynote = keynote  # ключевая заметка материала из Revit
        self.material = material  # материал из Revit
        self.work = work  # работа, параметры которой нужно назначить


def read_from_excel(path):
    excel = Excel.ApplicationClass()
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        workbook = excel.Workbooks.Open(path)
        ws_1 = workbook.Worksheets(1)
        row_end_1 = ws_1.Cells.Find("*", System.Reflection.Missing.Value,
                                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                    Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                    False, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row

        d = {}

        code = ""
        chapter = ""
        title_of_work = ""
        unit_of_measurement = ""

        for i in range(2, row_end_1 + 1):
            unit_of_measurement = ws_1.Cells(i, 3).Text

            if not unit_of_measurement:
                chapter = ws_1.Cells(i, 2).Text
                continue
            else:
                code = ws_1.Cells(i, 1).Text
                title_of_work = ws_1.Cells(i, 2).Text
                work = Work(code, chapter, title_of_work, unit_of_measurement)
                d[code] = work

    finally:
        excel.ActiveWorkbook.Close(False)
        Marshal.ReleaseComObject(ws_1)
        Marshal.ReleaseComObject(workbook)
        Marshal.ReleaseComObject(excel)
    return d


def get_calculation_type_value(unit_value):
    if CALCULATION_TYPE_DICT.has_key(unit_value):
        return CALCULATION_TYPE_DICT[unit_value]
    else:
        return "Ошибка"


def set_param(param, value, edited):
    if param.AsValueString() == str(value):
        return edited
    else:
        param.Set(value)
        edited = True
        return edited


def set_classifier_parameters(revit_materials):
    with revit.Transaction("BIM: Заполнение параметров классификатора"):
        try:
            for revit_material in revit_materials:
                edited = False
                material = revit_material.material
                work = revit_material.work

                edited = set_param(
                    material.GetParam(CHAPTER_PARAMETER),
                    work.chapter,
                    edited)

                edited = set_param(
                    material.GetParam(WORK_TITLE_PARAMETER),
                    work.title_of_work,
                    edited)

                edited = set_param(
                    material.GetParam(UNIT_PARAMETER),
                    work.unit_of_measurement,
                    edited)

                calculation_type = get_calculation_type_value(work.unit_of_measurement)
                edited = set_param(
                    material.GetParam(CALCULATION_TYPE_PARAMETER),
                    calculation_type,
                    edited)

                if edited:
                    report_edited.append(["ИЗМЕНЕН", revit_material.keynote, material.Name])
                else:
                    report_not_edited.append(["БЕЗ ИЗМЕНЕНИЙ", revit_material.keynote, material.Name])
        except:
            report_errors.append(["ОШИБКА ПРИ ЗАПИСИ", revit_material.keynote, material.Name])


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    excel_path = pick_file()
    if excel_path:
        TaskDialog.Show("Example Title", excel_path)

    # Добавить проверку на корректность пути
    # alert("Не указан путь к ФОП или файлу Excel", exitscript=True)
    dict_from_excel = read_from_excel(excel_path)

    materials = (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_Materials)
                 .ToElements())
    print("Найдено материалов: " + str(len(materials)))

    revit_materials = []

    for material in materials:

        keynote = material.GetParam(BuiltInParameter.KEYNOTE_PARAM).AsString()

        # Отсеиваем ситуации, когда у материала не указана Ключевая заметка (код работы)
        if not keynote:
            report_no_work_code.append(["НЕТ КОДА РАБОТЫ", "", material.Name])
            continue

        # Отсеиваем ситуации, когда Классификатор не содержит указанный в материале код
        if not dict_from_excel.has_key(keynote):
            report_classifier_code_not_found.append(["НЕ НАЙДЕН КОД", keynote, material.Name])
            continue

        revit_materials.append(RevitMaterial(keynote, material, dict_from_excel[keynote]))

    set_classifier_parameters(revit_materials)

    report = (report_errors
              + report_edited
              + report_not_edited
              + report_no_work_code
              + report_classifier_code_not_found)
    output = script.output.get_output()
    output.print_table(table_data=report,
                       title="Отчет работы плагина",
                       columns=["Статус⠀⠀⠀⠀⠀⠀⠀⠀", "Код работы", "Имя материала"],
                       formats=['', '', ''])


script_execute()
