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

output = script.output.get_output()

param_name_for_mark = "Марка"
"""Имя параметра, где плагин будет искать марку сваи, записанную пользователем"""

param_name_for_write = "ФОП_Примечание"
"""Имя параметра, в который плагин будет записывать диапазон марок свай (напр., "1-3, 5")"""

pile_ids_without_mark = []
"""Список, который будет содержать id свай, у которых нет или не корректная марка"""

mark_separator = ", "
"""Разделитель между марками, например, в '1, 2, 3, 4' разделитель - ', ' """


class RevitPileType:
    """Класс-оболочка для типоразмера элемента сваи."""

    def __init__(self, pile_type_name, pile_elevation_after_driving, pile_elevation_after_cutting):
        self.pile_type_name = pile_type_name
        """Имя типоразмера сваи"""

        self.elevation_after_driving = pile_elevation_after_driving
        """Высотная отметка сваи после забивки"""

        self.elevation_after_cutting = pile_elevation_after_cutting
        """Высотная отметка сваи после срубки"""

        self.piles = []
        """Перечень свай этого типоразмера семейства"""

        self.marks = []
        """
        Перечень марок свай этого типоразмера списком.
        Используется для получения диапазона свай.
        Заполняется одновременно с добавлением сваи через add_pile().
        """

        self.all_marks = ""
        """
        Перечень марок свай этого типоразмера в строку.
        Используется для вывода в отчет.
        Заполняется одновременно с добавлением сваи через add_pile().
        Пример: "1, 2, 3, 5"...
        """

        self.mark_range = ""
        """
        Диапазон марок свай этого типоразмера.
        \nПример: "1-3, 5"
        """

    def add_pile(self, pile):
        """ Добавляет:
            - элемент сваи в список свай RevitPileType (piles);
            - марку сваи из параметра "Марка" в список марок RevitPileType (marks);
            - марку сваи в общий список марок свай (all_marks).
        """

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
        self.piles.append(pile)

    def add_mark(self, mark):
        self.marks.append(mark)
        self.all_marks += '{}{}'.format(mark, mark_separator)

    def get_range(self):
        """
        Получает значение диапазона марок свай и записывает в mark_range.
        Напр., "1,2,3,4" -> "1-3,4"
        """
        self.marks.sort()
        for r in self.get_ranges(self.marks):
            value_0 = r[0]
            value_1 = r[1]
            if value_0 == value_1:
                self.mark_range += str(value_0) + mark_separator
            else:
                self.mark_range += str(value_0) + "-" + str(value_1) + mark_separator

        # "1, 2, 3, 4, " -> "1, 2, 3, 4"
        self.mark_range = remove_suffix(self.mark_range, mark_separator)

    def get_ranges(self, list_of_marks):
        """
        Возвращает значение диапазона марок свай на основе списка марок.
        """
        for a, b in itertools.groupby(enumerate(list_of_marks), lambda pair: pair[1] - pair[0]):
            b = list(b)
            yield b[0][1], b[-1][1]

    def write_ranges(self):
        """
        Записывает диапазон марок mark_range в параметр param_name_for_write у каждой сваи.
        Обычно param_name_for_write = "ФОП_Примечание"
        """
        if not self.piles:
            print('При работе с типом {} не было отобрано ни одной сваи!'.format(self.pile_type_name))
        for pile in self.piles:
            pile.GetParam(param_name_for_write).Set(self.mark_range)

    def get_all_marks(self):
        return remove_suffix(self.all_marks, mark_separator)


def remove_suffix(string, suffix):
    """
    :param string: строка, в которой нужно удалить суффикс
    :param suffix: подстрока, которую нужно найти в строке и удалить
    :return: Возвращает строку без указанного окончания или полученную строку в исходном состоянии,
    если указанного окончания нет.
    Напр., "1, 2, 3, 4, " -> "1, 2, 3, 4", когда подстрока ", "
    """
    if string.endswith(suffix):
        last_sep_index = string.rfind(suffix)
        return string[:last_sep_index]
    else:
        return string


def get_piles():
    """
    Возвращает список выбранных свай - экземпляров Revit, которые имеют в имени типа "Свая".
    """
    selected_elem_ids = uidoc.Selection.GetElementIds()
    if selected_elem_ids.Count == 0:
        output.close()
        alert("Выберите сваи перед тем, как использовать плагин", exitscript=True)
    piles = []
    for id in selected_elem_ids:
        elem = doc.GetElement(id)
        if "Свая" in elem.Symbol.FamilyName:
            piles.append(elem)

    if len(piles) == 0:
        output.close()
        alert("Выберите сваи", exitscript=True)
    return piles


def get_pile_types(piles):
    """
    Возвращает список типоразмеров обернутых классом-оболочкой RevitPileType
    """
    dictionary = {}
    for pile in piles:
        """
        Распределение по типам необходимо выполнить по трем параметрам:
        - имя типа;
        - высотная отметка головы сваи;
        - высотная отметка головы сваи после срубки.
        Для удобства сформируем ключ из этих значений.
        """
        pile_name = pile.Name
        pile_elevation_after_driving_as_str, pile_elevation_after_cutting_as_str = get_pile_elevations(pile)

        # Формируем строку "Тип 1_-3500_-3700"
        key = '{0}_{1}_{2}'.format(
            pile_name,
            pile_elevation_after_driving_as_str,
            pile_elevation_after_cutting_as_str)

        if key not in dictionary.keys():
            dictionary[key] = RevitPileType(
                pile_name,
                pile_elevation_after_driving_as_str,
                pile_elevation_after_cutting_as_str)
        dictionary[key].add_pile(pile)

    return dictionary.values()

def get_pile_elevations(pile):
    # Получаем высотную отметку головы сваи
    pile_elevation_after_driving = pile.GetParam("ФОП_Смещение от уровня").AsValueString()
    pile_elevation_after_driving_as_int = int(pile_elevation_after_driving)

    # Получаем длину срезки сваи
    pile_cutting_length = pile.GetParam("ФОП_Сваи_Срубка головы_Длина").AsValueString()
    pile_cutting_length_as_int = int(pile_cutting_length)

    # Получаем высотную отметки головы сваи после срезки
    pile_elevation_after_cutting_as_int = pile_elevation_after_driving_as_int - pile_cutting_length_as_int

    pile_elevation_after_driving_as_str = str(pile_elevation_after_driving_as_int)
    pile_elevation_after_cutting_as_str = str(pile_elevation_after_cutting_as_int)
    return pile_elevation_after_driving_as_str, pile_elevation_after_cutting_as_str

def write_values_of_pile_ranges(pile_types):
    """
    Получает диапазон марок для каждого типоразмера свай.
    Записывает в параметр param_name_for_write у каждого экземпляра сваи.
    Обычно param_name_for_write = "ФОП_Примечание".
    Возвращает отчет для табличного вывода,
    где каждый элемент отчета - [{1}, {2}, {3}, {4}, {5}], где
    - {1} - имя типоразмера;
    - {2} - марки всех свай типоразмера;
    - {3} - диапазон марок
    - {4} - высотная отметка после забивки
    - {5} - высотная отметка после срубки
    """
    report = []
    with revit.Transaction("BIM: Диапазоны свай"):
        for pile_type in pile_types:
            pile_type.get_range()
            pile_type.write_ranges()
            # Формируем отчет по колонкам:
            #   - имя типоразмера свай;
            #   - марки свай этого типа через запятую;
            #   - марки свай этого типа диапазоном;
            #   - высотная отметка после забивки;
            #   - высотная отметка после срубки
            report.append(
                [
                    pile_type.pile_type_name,
                    pile_type.get_all_marks(),
                    pile_type.mark_range,
                    pile_type.elevation_after_driving,
                    pile_type.elevation_after_cutting
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

    print("В таблице ниже успешно обработанные сваи:")
    output.print_table(table_data=report,
                       title="Отчет работы плагина",
                       columns=[
                           "Тип сваи",
                           "Марки свай",
                           "Диапазон марок",
                           "Отм. после забивки",
                           "Отм. после срубки"])

    if pile_ids_without_mark:
        print("\nНайдены некорректно замаркированные сваи:")
        for pile_id_without_mark in pile_ids_without_mark:
            print('{}: {}'.format("- свая с id: ", output.linkify(pile_id_without_mark)))


script_execute()
