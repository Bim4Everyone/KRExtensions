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

param_name_for_length = 'ФОП_РАЗМ_Длина'
param_name_for_width = 'ФОП_РАЗМ_Ширина'
param_name_for_height = 'ФОП_РАЗМ_Высота'
param_name_for_reinforcement = 'обр_ФОП_АРМ_Пилон'
param_name_for_write = 'Марка'

tag_family_name = '!Марка_Несущая колонны'
tag_type_name = 'Марка_Полка 20 мм'

tag_elbow_offset = XYZ(2.0, 3.0, 0.0)
tag_header_offset = XYZ(3.5, 0.0, 0.0)

report_about_write = []

def get_pylons():
    selected_ids = uidoc.Selection.GetElementIds()

    if len(selected_ids) == 0:
        global output
        output.close()
        alert("Не выбрано ни одного элемента", exitscript=True)

    elements = [doc.GetElement(selectedId) for selectedId in selected_ids]
    pylons = [elem for elem in elements
              if elem.Category.Name.Contains('Несущие колонны') and elem.Name.Contains('Пилон')]

    if len(pylons) == 0:
        global output
        output.close()
        alert("Не выбрано ни одного пилона", exitscript=True)
    return pylons


def convert_to_int_from_internal_value(value, unit_type):
    return int(round(UnitUtils.ConvertFromInternalUnits(value, unit_type)))


def get_waterproofing(pylon):
    waterproofing = ''
    for material_id in pylon.GetMaterialIds(False):
        material = doc.GetElement(material_id)
        if material.Name.Contains('Бетон с Пенетроном'):
            waterproofing = '-д'
            break
    return waterproofing


def get_pylon_data(pylons):
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
            print ("Ошибка у пилона с id {0}: ".format(pylon.Id) + e.message)
    return pylon_and_data_pairs


def write_pylon_data(pylon_and_data_pairs):
    with revit.Transaction("КР: Маркировка пилонов"):
        for pair_for_write in pylon_and_data_pairs:
            pylon = pair_for_write[0]
            string_for_write = pair_for_write[1]
            try:
                pylon.GetParam(param_name_for_write).Set(string_for_write)
                report_about_write.append([pylon.Name, output.linkify(pylon.Id), string_for_write, '', pylon])
            except:
                error = "Не удалось записать значение у пилона с id: " + str(pylon.Id)
                report_about_write.append([pylon.Name, output.linkify(pylon.Id), error, '', pylon])


def get_pylon_tag_type_id():
    families = FilteredElementCollector(doc).OfCategory(BuiltInCategory.INVALID).OfClass(Family)

    tag_family = None
    for family in families:
        if family.Name.Contains(tag_family_name):
            tag_family = family
            break

    if tag_family is None:
        return None

    for symbol_id in tag_family.GetFamilySymbolIds():
        tag_symbol = doc.GetElement(symbol_id)
        tag_symbol_name = tag_symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        if tag_symbol_name == tag_type_name:
            return symbol_id
    return None


def pylon_markings():
    if active_view.ViewType == ViewType.FloorPlan or active_view.ViewType == ViewType.EngineeringPlan:
        print('Выполняем поиск марки несущих колонн семейства \"{0}\" типоразмера \"{1}\".'
              .format(tag_family_name, tag_type_name))
        tag_type_id = get_pylon_tag_type_id()
        if tag_type_id is None:
            print("Не найдено семейство или типоразмер марки, поэтому будем размещать стандартную.")
        else:
            print("Семейство и типоразмер марки найдено! Выполняем размещение марок на активном виде.")

        place_pylon_tags(tag_type_id)
    else:
        print('Текущий вид не является видом в плане, поэтому размещение марок производиться не будет!')


def already_has_mark(pylon, tag_type_id):
    element_class_filter = ElementClassFilter(IndependentTag)
    pylon_tag_ids = pylon.GetDependentElements(element_class_filter)

    for pylon_tag_id in pylon_tag_ids:
        existing_pylon_tag = doc.GetElement(pylon_tag_id)
        if existing_pylon_tag.GetTypeId() == tag_type_id and existing_pylon_tag.OwnerViewId == active_view.Id:
            return True
    return False


def place_pylon_tags(tag_type_id):
    tag_mode = TagMode.TM_ADDBY_CATEGORY
    tag_orientation = TagOrientation.Horizontal

    with revit.Transaction("КР: Размещение марок пилонов"):
        for report_string in report_about_write:
            pylon = report_string[4]

            # Если пилон уже имеет марку нужного нам типа, то пропускаем, размещать повторно не нужно
            if already_has_mark(pylon, tag_type_id):
                report_string[3] = '<Уже размещена>'
                continue

            try:
                pylon_ref = Reference(pylon)
                pylon_mid = pylon.Location.Point

                leader_point = pylon_mid + XYZ(5.0, 5.0, 0.0)
                pylon_tag = IndependentTag.Create(doc, active_view.Id, pylon_ref, True, tag_mode, tag_orientation, leader_point)

                pylon_tag.LeaderEndCondition = LeaderEndCondition.Free
                elbow_point = pylon_tag.LeaderEnd + tag_elbow_offset
                pylon_tag.LeaderElbow = elbow_point

                header_point = elbow_point + tag_header_offset
                pylon_tag.TagHeadPosition = header_point

                if tag_type_id is not None:
                    pylon_tag.ChangeTypeId(tag_type_id)

                report_string[3] = output.linkify(pylon_tag.Id)
            except:
                report_string[3] = '<Не размещена>'


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    print("Здравствуйте! Данный плагин предназначен для маркировки пилонов на основе следующих параметров:")
    print("- \"ФОП_РАЗМ_Длина\"")
    print("- \"ФОП_РАЗМ_Ширина\"")
    print("- \"ФОП_РАЗМ_Высота\"")
    print("- \"обр_ФОП_АРМ_Пилон\"")

    print("Запись будет производиться в параметр \"Марка\".")

    print("Из всех выбранных элементов будут отобраны только те, что имеют имя типа со словом \"Пилон\".")
    print("Например, \"НН_Пилон-350х1800 (ЖБ B40 F150 W6)\".")

    pylons = get_pylons()
    print("Найдено пилонов: {0}".format(len(pylons)))

    pylon_and_data_pairs = get_pylon_data(pylons)
    write_pylon_data(pylon_and_data_pairs)

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
