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

param_name_for_length = 'ФОП_РАЗМ_Длина'
param_name_for_width = 'ФОП_РАЗМ_Ширина'
param_name_for_height = 'ФОП_РАЗМ_Высота'
param_name_for_reinforcement = 'обр_ФОП_АРМ_Пилон'
param_name_for_write = 'Марка'


def get_pylons():
    selected_ids = uidoc.Selection.GetElementIds()

    if len(selected_ids) == 0:
        output = script.output.get_output()
        output.close()
        alert("Не выбрано ни одного элемента", exitscript=True)

    elements = [doc.GetElement(selectedId) for selectedId in selected_ids]
    pylons = [elem for elem in elements if elem.Name.Contains('Пилон')]

    if len(pylons) == 0:
        output = script.output.get_output()
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


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    pairs_for_write = []

    for pylon in get_pylons():
        try:
            pylon_type = doc.GetElement(pylon.GetTypeId())

            length = convert_to_int_from_internal_value(
                pylon_type.GetParamValue(param_name_for_length),
                UnitTypeId.Decimeters)

            width = convert_to_int_from_internal_value(
                pylon_type.GetParamValue(param_name_for_width),
                UnitTypeId.Centimeters)

            height = convert_to_int_from_internal_value(
                pylon.GetParamValue(param_name_for_height),
                UnitTypeId.Decimeters)

            reinforcement = pylon.GetParamValue(param_name_for_reinforcement)
            if reinforcement is None:
                raise Exception("Не удалось определить армирование, заполните ТЗА")

            waterproofing = get_waterproofing(pylon)

            string_for_write = ('{0}.{1}.{2}-{3}{4}'
                                .format(str(length), str(width), str(height), reinforcement, waterproofing))

            pairs_for_write.append([pylon, string_for_write])

        except Exception as e:
            alert(e.message + " у пилона с id: " + str(pylon.Id), exitscript=False)

    for pair_for_write in pairs_for_write:
        with revit.Transaction("КР: Маркировка пилонов"):
            pylon = pair_for_write[0]
            string_for_write = pair_for_write[1]
            pylon.GetParam(param_name_for_write).Set(string_for_write)


script_execute()
