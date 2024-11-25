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

struct_columns_category_id = ElementId(BuiltInCategory.OST_StructuralColumns)
'''Id категории "Несущие колонны" для фильтрации нужных элементов'''

pylon_type_name_keyword = 'Пилон'
'''Ключевое слово в имени типа элемента для фильтрации нужных элементов'''

material_with_waterproofing_name = 'Бетон с Пенетроном'
'''Фрагмент имени материала,указывающий на то, что в бетоне будут присутствовать гидрофобные добавки'''

param_name_for_length = 'ФОП_РАЗМ_Длина'
param_name_for_width = 'ФОП_РАЗМ_Ширина'
param_name_for_height = 'ФОП_РАЗМ_Высота'
param_name_for_reinforcement = 'ТЗА_Характеристики'
param_name_for_write = 'Марка'

tag_family_name = '!Марка_Несущая колонны'

tag_symbols_dict = {}
tag_symbol_name_prefix = 'Марка_Полка '
tag_symbol_name_suffix = ' мм'

tag_elbow_offset = XYZ(2.0, 3.0, 0.0)
tag_header_offset = XYZ(3.5, 0.0, 0.0)

report_about_write = []


def get_pylons():
    """
    Получает элементы пилонов среди тех элементов, что были выделены пользователем перед запуском плагина
    :return: Список пилонов
    """
    selected_ids = uidoc.Selection.GetElementIds()

    if len(selected_ids) == 0:
        global output
        output.close()
        alert("Не выбрано ни одного элемента", exitscript=True)

    elements = [doc.GetElement(selectedId) for selectedId in selected_ids]
    pylons = [elem for elem in elements
              if elem.Category.Id == struct_columns_category_id and elem.Name.Contains(pylon_type_name_keyword)]

    if len(pylons) == 0:
        global output
        output.close()
        alert("Не выбрано ни одного пилона", exitscript=True)
    return pylons


def convert_to_int_from_internal_value(value, unit_type):
    """
    Конвертирует значение из встроенного формата в целочисленное заданного типа
    :param value: значение во встроенном формате Revit
    :param unit_type: ForgeTypeId типа значения, в который нужно перевести данные
    :return: сконвертированнное в заданный тип значение
    """
    return int(round(UnitUtils.ConvertFromInternalUnits(value, unit_type)))


def get_waterproofing(pylon):
    """
    Получает значение суффикса для марки пилона в зависимости от наличия добавок в его бетоне
    :param pylon: элемент пилона
    :return: суффикс для марки пилона
    """
    waterproofing = ''
    for material_id in pylon.GetMaterialIds(False):
        material = doc.GetElement(material_id)
        if material.Name.Contains(material_with_waterproofing_name):
            waterproofing = '-д'
            break
    return waterproofing


def get_pylon_data(pylons):
    """
    Получает значение марки в соответствии со значениями его параметров
    :param pylons: список пилонов
    :return: список пар пилон - значение для записи в его параметр Марка
    """
    pylon_and_data_pairs = []
    for pylon in pylons:
        try:
            pylon_type = doc.GetElement(pylon.GetTypeId())
            # Получаем длину пилона
            length = convert_to_int_from_internal_value(
                pylon_type.GetParamValue(param_name_for_length),
                UnitTypeId.Decimeters)
            # Получаем толщину пилона
            width = convert_to_int_from_internal_value(
                pylon_type.GetParamValue(param_name_for_width),
                UnitTypeId.Centimeters)
            # Получаем высоту пилона
            height = convert_to_int_from_internal_value(
                pylon.GetParamValue(param_name_for_height),
                UnitTypeId.Decimeters)
            # Получаем выбранное пользователем армирование для пилона
            reinforcement = pylon.GetParamValue(param_name_for_reinforcement)
            if reinforcement is None:
                raise Exception("Не удалось определить армирование, заполните ТЗА")
            # Получаем суффикс, указывающий на наличие добавок в бетоне пилона
            waterproofing = get_waterproofing(pylon)
            # Формируем строку в требуемом формате
            string_for_write = ('{0}.{1}.{2}-{3}{4}'
                                .format(str(length), str(width), str(height), reinforcement, waterproofing))

            pylon_and_data_pairs.append([pylon, string_for_write])

        except Exception as e:
            print ("Ошибка у пилона с id {0}: ".format(output.linkify(pylon.Id)) + e.message)
    return pylon_and_data_pairs


def write_pylon_data(pylon_and_data_pairs):
    """
    Записывает значение в параметр Марка у пилона
    :param pylon_and_data_pairs: список пар пилон - значение для записи в его параметр Марка
    """
    with revit.Transaction("BIM: Маркировка пилонов"):
        for pair_for_write in pylon_and_data_pairs:
            pylon = pair_for_write[0]
            string_for_write = pair_for_write[1]
            try:
                pylon.SetParamValue(param_name_for_write, string_for_write)
                report_about_write.append([pylon.Name, output.linkify(pylon.Id), string_for_write, '', pylon])
            except:
                error = "Не удалось записать значение у пилона"
                report_about_write.append([pylon.Name, output.linkify(pylon.Id), error, '', pylon])


def get_pylon_tag_types():
    """
    Заполняет словарь, где ключ - длина подчеркнутой части марки, значение - типоразмера марки.
    """
    print('Выполняем поиск семейства марки несущих колонн \"{0}\".'.format(tag_family_name))
    families = FilteredElementCollector(doc).OfClass(Family)

    tag_family = None
    for family in families:
        if family.Name.Equals(tag_family_name):
            tag_family = family
            break

    if tag_family is None:
        print("Семейство марки \"{0}\" не найдено, поэтому будем размещать стандартную.".format(tag_family_name))
        return None

    print('Семейство марки найдено! Выполняем поиск нужных типоразмеров марки.')
    print('Ищем те типоразмеры, которые начинаются с \"{0}\" и заканчиваются на \"{1}\".'
          .format(tag_symbol_name_prefix, tag_symbol_name_suffix))

    for symbol_id in tag_family.GetFamilySymbolIds():
        tag_symbol = doc.GetElement(symbol_id)
        tag_symbol_name = tag_symbol.GetParamValue(BuiltInParameter.SYMBOL_NAME_PARAM)

        if tag_symbol_name.startswith(tag_symbol_name_prefix) and tag_symbol_name.endswith(tag_symbol_name_suffix):
            try:
                number = tag_symbol_name.replace(tag_symbol_name_prefix, '').replace(tag_symbol_name_suffix, '')
                number = int(number) * 100
                tag_symbols_dict[number] = tag_symbol
            except:
                continue

    if tag_symbols_dict.keys():
        print("Подходящие типоразмеры марки найдены:")
        for tag_symbol in tag_symbols_dict.values():
            tag_symbol_name = tag_symbol.GetParamValue(BuiltInParameter.SYMBOL_NAME_PARAM)
            print('- \"{0}\"'.format(tag_symbol_name))
    else:
        print("Не найдены необходимые типоразмеры в семейство марки \"{0}\", размещать будем стандартную."
              .format(tag_family_name))


def pylon_markings():
    """
    Проверяет, что активный вид - вид в плане. Если да - запускает расстановку марок
    """
    if active_view.ViewType == ViewType.FloorPlan or active_view.ViewType == ViewType.EngineeringPlan:
        place_pylon_tags()
        print('Выполняем размещение марок пилонов.')
    else:
        print('Текущий вид не является видом в плане, поэтому размещение марок производиться не будет!')


def place_pylon_tags():
    """
    Размещает марки пилонов на активном виде. Назначает типоразмер марки в зависимости от длины текста марки
    """
    tag_mode = TagMode.TM_ADDBY_CATEGORY
    tag_orientation = TagOrientation.Horizontal

    with revit.Transaction("BIM: Размещение марок пилонов"):
        for report_string in report_about_write:
            try:
                pylon = report_string[4]

                # Если пилон уже имеет марку нужного нам типа, то пропускаем, размещать повторно не нужно
                if already_has_mark(pylon):
                    report_string[3] = '<Уже размещена>'
                    continue

                pylon_ref = Reference(pylon)
                pylon_mid = pylon.Location.Point

                leader_point = pylon_mid + XYZ(5.0, 5.0, 0.0)
                pylon_tag = IndependentTag.Create(doc, active_view.Id, pylon_ref, True, tag_mode, tag_orientation, leader_point)

                pylon_tag.LeaderEndCondition = LeaderEndCondition.Free
                elbow_point = pylon_tag.LeaderEnd + tag_elbow_offset
                pylon_tag.LeaderElbow = elbow_point

                header_point = elbow_point + tag_header_offset
                pylon_tag.TagHeadPosition = header_point

                # Если типоразмеры с нужными именами найдены, то получаем нужный по длине текста типоразмер марки
                # Смысл в том, чтобы найти такую марку, которая будет умещать текст (берется из параметра пилона),
                # но при этом быть минимальной по длине, чтобы не занимать линиями место на чертеже
                if tag_symbols_dict.keys():
                    tag_type_id = get_needed_tag_type_id(pylon_tag)
                    if tag_type_id is not None:
                        pylon_tag.ChangeTypeId(tag_type_id)

                report_string[3] = output.linkify(pylon_tag.Id)
            except:
                report_string[3] = '<Не размещена>'


def already_has_mark(pylon):
    """
    Если пилон уже имеет марку на текущем виде, имя типа которой начинается с tag_symbol_name_prefix
    и заканчивается tag_symbol_name_suffix, то будем считать, что марка уже стоит и ставить ее не нужно
    """
    element_class_filter = ElementClassFilter(IndependentTag)
    pylon_tag_ids = pylon.GetDependentElements(element_class_filter)

    for pylon_tag_id in pylon_tag_ids:
        existing_pylon_tag = doc.GetElement(pylon_tag_id)
        existing_pylon_tag_symbol_name = existing_pylon_tag.Name
        if (existing_pylon_tag.OwnerViewId == active_view.Id and
                existing_pylon_tag_symbol_name.startswith(tag_symbol_name_prefix) and
                existing_pylon_tag_symbol_name.endswith(tag_symbol_name_suffix)):
            return True
    return False


def get_needed_tag_type_id(pylon_tag):
    """
    Получает id типоразмера марки пилона в зависимости от длины текста марки
    :param pylon_tag: марка пилона
    :return: id типоразмера марки пилона
    """
    # Запоминаем какое было значение
    temp_has_leader = pylon_tag.HasLeader
    # Выключаем указатель марки и регеним документ, чтобы марка перерисовалась
    pylon_tag.HasLeader = False
    '''
    Задача - получить типоразмер марки, который наилучшим образом подходит для нее (чтобы линия под текстом была самой 
    короткой, но при этом подчеркивала весь текст). 
    Для этого необходимо отключить линию выноски, получить размеры оставшейся марки по ширине, 
    подобрать из существующих типоразмеров нужный, вернуть линию выноски.
    Мы изменили свойство объекта в Revit, и далее будем при помощи BoundingBox узнавать его размеры
    Однако BoundingBox будет строиться вокруг визуального представления объекта, а оно осталось прежним
    Визуальное представление объекта можно обновить при помощи метода doc.Regenerate()
    '''
    doc.Regenerate()

    # Получаем ширину марки без выноски - только текстовое поле
    bounding_box = pylon_tag.get_BoundingBox(doc.ActiveView)
    bounding_box_max = bounding_box.Max
    bounding_box_min = bounding_box.Min
    length = float(convert_to_int_from_internal_value(bounding_box_max.X - bounding_box_min.X, UnitTypeId.Millimeters))

    # Определяем ближайшее большее значение из имеющихся в типоразмерах
    global tag_symbols_dict
    sorted_keys = sorted(tag_symbols_dict.keys())
    closest_length = min(sorted_keys, key=lambda x: abs(x-length)) if tag_symbols_dict.keys() else None

    if closest_length < length:
        index = sorted_keys.index(closest_length) + 1
        if index < len(sorted_keys):
            closest_length = sorted_keys[index]

    pylon_tag.HasLeader = temp_has_leader
    return tag_symbols_dict[closest_length].Id



@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    print("Здравствуйте! Данный плагин предназначен для маркировки пилонов на основе следующих параметров:")
    print("- \"ФОП_РАЗМ_Длина\"")
    print("- \"ФОП_РАЗМ_Ширина\"")
    print("- \"ФОП_РАЗМ_Высота\"")
    print("- \"обр_ФОП_АРМ_Пилон\"")
    print("- <Имя материала> (на наличие строки \"Бетон с Пенетроном\")")

    print("Запись будет производиться в параметр \"Марка\". Пример записи: \"18.35.40-35ф24-д\"")

    print("Из всех выбранных элементов будут отобраны только те, что имеют категорию \"Несущие колонны\" "
          "и имя типа со словом \"Пилон\".")
    print("Например, \"НН_Пилон-350х1800 (ЖБ B40 F150 W6)\".")

    pylons = get_pylons()
    print("⠀")
    print("Найдено пилонов: {0}".format(len(pylons)))
    pylon_and_data_pairs = get_pylon_data(pylons)
    print("⠀")
    write_pylon_data(pylon_and_data_pairs)
    print("Запись параметра \"Марка\" произведена!")
    print("⠀")

    get_pylon_tag_types()
    pylon_markings()

    '''
        При печати таблицы встречается ошибка, когда таблица по неустановленной причине печататься не хочет
        Решается добавлением в коллекцию еще одной строки
        Чтобы не отвлекать пользователя в доп строке содержится один невидимый символ (это не пробел)
    '''
    report_about_write.append(["⠀", "", "", "", ""])
    output.print_table(table_data=report_about_write[:],
                       title="Отчет работы плагина",
                       columns=[
                           "Имя типоразмера",
                           "ID пилона",
                           "Марка пилона",
                           "ID марки"]
                       )

    print("Скрипт завершил работу!")


script_execute()
