# -*- coding: utf-8 -*-
import os
import clr
import datetime
from System.Collections.Generic import *

import itertools

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

param_name_for_mark = "Марка"
param_name_for_write = "ФОП_Примечание"

# Список, который будет содержать id свай,у которых нет или не корректная марка
pile_ids_without_mark = []

class RevitPileType:
    def __init__(self, pile_type_name):
        self.pile_type_name = pile_type_name
        self.piles = []
        self.marks = []
        self.all_marks = ""
        self.mark_range = ""

    def add_pile(self, pile):
        self.piles.append(pile)
        mark_as_str = pile.GetParam(BuiltInParameter.ALL_MODEL_MARK).AsString()
        if not mark_as_str:
            pile_ids_without_mark.append(pile.Id)
            return
        # Пытаемся перевести марку сваи в число
        try:
            mark = int(mark_as_str)
        except:
            pile_ids_without_mark.append(pile.Id)
            return
        self.add_mark(mark)

    def add_mark(self, mark):
        self.marks.append(mark)
        self.all_marks += '{}, '.format(mark)

    def get_range(self):
        self.marks.sort()
        for r in self.get_ranges(self.marks):
            value_0 = r[0]
            value_1 = r[1]
            if value_0 == value_1:
                self.mark_range += str(value_0) + ", "
            else:
                self.mark_range += str(value_0) + "-" + str(value_1) + ", "
        self.mark_range = self.mark_range[:-2]

    def get_ranges(self, list_of_marks):
        for a, b in itertools.groupby(enumerate(list_of_marks), lambda pair: pair[1] - pair[0]):
            b = list(b)
            yield b[0][1], b[-1][1]

    def write_ranges(self):
        if not self.piles:
            print('При работе с типом {} не было отобрано ни одной сваи!'.format(self.pile_type_name))
        for pile in self.piles:
            pile.GetParam(param_name_for_write).Set(self.mark_range)


def get_piles():
    selected_elem_ids = uidoc.Selection.GetElementIds()
    if selected_elem_ids.Count == 0:
        alert("Выберите сваи перед тем, как использовать плагин", exitscript=True)
    piles = []
    for id in selected_elem_ids:
        elem = doc.GetElement(id)
        if "Свая" in elem.Symbol.FamilyName:
            piles.append(elem)

    if len(piles) == 0:
        alert("Выберите сваи", exitscript=True)
    return piles


def get_pile_types(piles):
    dictionary = {}
    for pile in piles:
        pile_name = pile.Name
        if pile_name not in dictionary.keys():
            dictionary[pile_name] = RevitPileType(pile_name)
        dictionary[pile_name].add_pile(pile)

    return dictionary.values()


def write_values_of_pile_ranges(pile_types):
    report = []
    with revit.Transaction("BIM: Диапазоны свай"):
        for pile_type in pile_types:
            pile_type.get_range()
            pile_type.write_ranges()
            # Формируем отчет по колонкам:
            #   - имя типоразмера свай;
            #   - марки свай этого типа через запятую;
            #   - марки свай этого типа диапазоном;
            report.append(
                [
                    pile_type.pile_type_name,
                    pile_type.all_marks,
                    pile_type.mark_range
                ]
            )
    return report

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    print("Здравствуйте! Данный плагин предназначен для записи у выбранных свай в параметр ФОП_Примечание "
          + "информации о диапазоне марок с распределением по типам свай.")

    print("Собираю выбранные элементы свай...")
    piles = get_piles()
    print("Собрано свай: " + str(len(piles)))

    # Список типоразмеров свай, обернутых в класс-оболочку RevitPileType
    pile_types = get_pile_types(piles)
    print("Из них разных типов: " + str(len(pile_types)))

    # Выполняем расчет и запись диапазонов свай в экземпляры свай
    report = write_values_of_pile_ranges(pile_types)

    print("В таблице ниже обработанные сваи:")
    output = script.output.get_output()
    output.print_table(table_data=report,
                       title="Отчет работы плагина",
                       columns=["Тип сваи", "Марки свай", "Диапазон марок"])

    if pile_ids_without_mark:
        print("\nНайдены не замаркированные сваи:")
        for pile_id_without_mark in pile_ids_without_mark:
            print('{}: {}'.format("- свая с id: ", output.linkify(pile_id_without_mark)))


script_execute()
