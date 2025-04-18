# -*- coding: utf-8 -*-
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

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application
active_view = doc.ActiveView
output = script.output.get_output()

filed_name_for_start = "ВИДИМАЯ ЧАСТЬ"
filed_name_for_end = "СКРЫТАЯ ЧАСТЬ"


def get_schedule_field(schedule):
    """
    Возвращает список всех полей переданной спецификации
    """
    definition = schedule.Definition
    schedule_field_ids = definition.GetFieldOrder()

    schedule_fields = []
    for fieldId in schedule_field_ids:
        schedule_fields.append(definition.GetField(fieldId))
    return schedule_fields


def show_fields(schedule_fields):
    """
    Показывает указанные поля спецификации
    """
    for field in schedule_fields:
        field.IsHidden = False


def hide_fields(fields_for_hide):
    """
    Скрывает указанные поля спецификации
    """
    for field in fields_for_hide:
        field.IsHidden = True



def get_boundary_indexes(schedule_fields):
    """
    Возвращает начальный и конечный индексы столбцов пограничных полей,
    между которыми следует анализировать поля спеки на нуль
    """
    # Ищем специальные столбцы
    start_column_index = -1
    end_column_index = -1
    for field in schedule_fields:
        column_name = field.GetName()
        if column_name.Contains(filed_name_for_start):
            start_column_index = field.FieldIndex
        elif column_name.Contains(filed_name_for_end):
            end_column_index = field.FieldIndex

    # Проверяем найденные индексы
    if start_column_index == -1 or end_column_index == -1 or start_column_index >= end_column_index:
        alert("Не найдены разделительные поля! " +
              "Спецификация должна содержать поля с именем " + filed_name_for_start + " и " + filed_name_for_end +
              " Плагин будет искать между ними поля спецификации для анализа на 0", exitscript=True)
    return start_column_index, end_column_index


def get_target_column_indexes(start_column_index, end_column_index):
    """
    Возвращает индексы столбцов, которые следует анализировать
    """
    target_column_indexes = []
    for i in range(start_column_index + 1, end_column_index - 1):
        target_column_indexes.append(i)
    return target_column_indexes


def get_fields_for_hide(schedule_fields, column_indexes_for_show):
    """
    Возвращает поля спецификации, которые необходимо скрыть
    """
    fields_for_hide = []
    for field in schedule_fields:
        field_index = field.FieldIndex
        if not field_index in column_indexes_for_show:
            fields_for_hide.append(field)
    return fields_for_hide


def analyze_fields_by_zero(schedule, fields_for_hide, target_column_indexes):
    """
    Дописывает в список поля спецификации, все значения в столбцах которых равны 0
    """
    # Получаем данные таблицы
    definition = schedule.Definition
    schedule_field_ids = definition.GetFieldOrder()

    table_data = schedule.GetTableData()
    section_data = table_data.GetSectionData(SectionType.Body)
    row_count = section_data.NumberOfRows

    for column_index in target_column_indexes:
        all_zeros = True
        for row in range(row_count):
            # Проверяем только на основании расчетных полей
            if not section_data.GetCellType(row, column_index).ToString().Equals("ParameterText"):
                continue
            # Проверяем значение в ячейке на нуль
            if section_data.GetCellText(row, column_index) != "0":
                all_zeros = False
                break
        if all_zeros:
            field = definition.GetField(schedule_field_ids[column_index])
            fields_for_hide.append(field)



@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    schedule = doc.ActiveView
    if schedule.ViewType != ViewType.Schedule:
        alert("Активный вид не является спецификацией!", exitscript=True)

    with revit.Transaction("КР: Скрыть колонки ВРС"):
        # Получаем все поля спецификации
        schedule_fields = get_schedule_field(schedule)

        # Делаем все поля видимыми, чтобы удалось проанализировать ячейки на нуль и сопоставить поля и видимые столбцы
        show_fields(schedule_fields)

        # Получаем индексы столбцов полей спеки, между которыми нужно будет анализировать поля на 0
        start_column_index, end_column_index = get_boundary_indexes(schedule_fields)

        # Получаем индексы целевых столбцов для анализа на 0
        target_column_indexes = get_target_column_indexes(start_column_index, end_column_index)

        # Получаем предварительный список полей спеки, которые нужно обратно скрыть.
        # Это поля, которые нужны для расчета значений, но не выводятся в графику
        fields_for_hide = get_fields_for_hide(schedule_fields, target_column_indexes)

        # Анализируем целевые столбцы, и если они содержат нули, то добавляем их к тем, что нужно скрыть
        analyze_fields_by_zero(schedule, fields_for_hide, target_column_indexes)

        # Скрываем поля
        hide_fields(fields_for_hide)


script_execute()
