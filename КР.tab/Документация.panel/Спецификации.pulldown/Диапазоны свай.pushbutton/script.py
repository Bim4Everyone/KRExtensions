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

class RevitPileType:
    def __init__(self):
        self.piles = []
        self.marks = []
        self.mark_range = ""

    def get_range(self):
        self.marks.sort()
        for r in self.get_ranges(self.marks):
            value_0 = r[0]
            value_1 = r[1]
            if value_0 == value_1:
                self.mark_range += str(value_0) + ", "
            else:
                self.mark_range += str(value_0) + "-" + str(value_1) + ", "
        self.range = self.mark_range[:-2]

    def get_ranges(self, list_of_marks):
        for a, b in itertools.groupby(enumerate(list_of_marks), lambda pair: pair[1] - pair[0]):
            b = list(b)
            yield b[0][1], b[-1][1]

    def write_ranges(self):
        for pile in self.piles:
            pile.GetParam(param_name_for_write).Set(self.range)


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

def get_piles_by_types(piles):
    dictionary = {}
    pile_ids_without_mark = []
    for pile in piles:
        pile_name = pile.Name
        if pile_name not in dictionary.keys():
            dictionary[pile_name] = RevitPileType()
        dictionary[pile_name].piles.append(pile)

        mark_as_str = pile.GetParam(param_name_for_mark).AsString()
        if not mark_as_str:
            pile_ids_without_mark.append(pile.Id)
            continue
        mark = int(mark_as_str)
        dictionary[pile_name].marks.append(mark)
    return dictionary, pile_ids_without_mark

def write_values_of_pile_ranges(dictionary):
    with revit.Transaction("BIM: Диапазоны свай"):
        for key in dictionary.keys():
            revit_piles = dictionary[key]
            revit_piles.get_range()
            revit_piles.write_ranges()


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    print("Здравствуйте! Данный плагин предназначен для записи у выбранных свай в параметр ФОП_Примечание "
          + "информации о диапазоне марок с распределением по типам свай.")

    print("Собираю выбранные элементы свай...")
    piles = get_piles()
    print("Собрано свай: " + str(len(piles)))

    # В словаре элементы помещены в класс-оболочку RevitPileRange и сгруппированы по типам
    dictionary, pile_ids_without_mark = get_piles_by_types(piles)
    print("Из них разных типов: " + str(len(dictionary.keys())))

    # Выполняем расчет и запись диапазонов свай в экземпляры свай
    write_values_of_pile_ranges(dictionary)

    output = script.output.get_output()
    if pile_ids_without_mark:
        print("Найдены не замаркированные сваи:")
        for pile_id_without_mark in pile_ids_without_mark:
            print('{}: {}'.format("- свая с id: ", output.linkify(pile_id_without_mark)))


script_execute()
