# -*- coding: utf-8 -*-
import clr

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

from pyrevit import EXEC_PARAMS, revit
from pyrevit import forms
from pyrevit import script

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


class SelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        if isinstance(element, ScheduleSheetInstance):
            return True
        return False

    def AllowReference(self, reference, position):
        return False


class PartsSchedule:
    def __init__(self, schedule):
        self.schedule = schedule
        self.doc = schedule.Document
        self.elements = self.__get_schedule_elements()

        self.ifc_img_param_name = "обр_ФОП_Форма_изображение IFC"
        self.num_param_name = "обр_ФОП_Форма_номер"
        self.posit_param_name = "обр_ФОП_Позиция"
        self.excl_param_name = "обр_ФОП_Исключить из ВД"
        self.excl_param_field = self.__get_field_by_name(self.excl_param_name)

    def check_schedule(self):
        params = [self.ifc_img_param_name,
                  self.excl_param_name,
                  self.num_param_name,
                  self.posit_param_name]
        errors = []
        for param in params:
            if not self.__get_field_by_name(param):
                error_info = ["", param]
                errors.append(error_info)

        if errors:
            output = script.get_output()
            output.print_table(table_data=errors,
                               title="В выбранной ведомости отсутствуют необходимые параметры:",
                               columns=["", "Имя параметра"])
            return False

        filters = self.schedule.Definition.GetFilters()
        used_filters = [x for x in filters if x.FieldId == self.excl_param_field.FieldId]
        if used_filters:
            forms.alert("В выбранной ведомости уже присутствует фильтр по параметру "
                        "'обр_ФОП_Исключить из ВД'.\nФильтр требуется удалить.")
            return False
        return True

    def update_elements(self):
        with revit.Transaction("BIM: Обновить ведомости"):
            dict_by_number = self.__group_elements()
            for number_key in dict_by_number.keys():
                for position_key in dict_by_number[number_key].keys():
                    for family_key in dict_by_number[number_key][position_key].keys():
                        elements = dict_by_number[number_key][position_key][family_key]
                        for element in elements:
                            element.SetParamValue(self.excl_param_name, "0")

                    random_key = dict_by_number[number_key][position_key].keys()[0]
                    elements_true = dict_by_number[number_key][position_key][random_key]
                    for element in elements_true:
                        element.SetParamValue(self.excl_param_name, "1")

    def update_filters(self):
        field_id = self.excl_param_field.FieldId
        filter_type = ScheduleFilterType.Equal
        filter_value = "1"
        schedule_filter = ScheduleFilter(field_id, filter_type, filter_value)
        with revit.Transaction("BIM: Обновить ведомости"):
            self.schedule.Definition.AddFilter(schedule_filter)

    def __get_schedule_elements(self):
        elements = FilteredElementCollector(self.doc, self.schedule.Id)
        elements.ToElements()
        return elements

    def __get_field_by_name(self, name):
        fields = self.schedule.Definition.GetFieldOrder()
        for field_id in fields:
            field = self.schedule.Definition.GetField(field_id)
            if not field.IsCalculatedField:
                if field.ParameterId.IntegerValue > 0:
                    if self.doc.GetElement(field.ParameterId).Name == name:
                        return field

    def __get_param_value(self, element, param_name, is_inst):
        if not is_inst:
            type_id = element.GetTypeId()
            element = self.doc.GetElement(type_id)
        return element.GetParamValueOrDefault(param_name)

    def __group_elements(self):
        dict_by_number = self.__create_dict_by_param(self.elements, self.num_param_name, False)
        for key in dict_by_number.keys():
            dict_by_prior = self.__create_dict_by_param(dict_by_number[key], self.posit_param_name, True)
            dict_by_number[key] = dict_by_prior

        for key_num in dict_by_number.keys():
            for key_prior in dict_by_number[key_num].keys():
                elements = dict_by_number[key_num][key_prior]
                dict_by_family = self.__create_dict_by_param(elements, BuiltInParameter.ELEM_FAMILY_PARAM, True)
                dict_by_number[key_num][key_prior] = dict_by_family

        return dict_by_number

    def __create_dict_by_param(self, elements, param_name, is_inst):
        result_dict = dict()
        for element in elements:
            param_value = self.__get_param_value(element, param_name, is_inst)
            result_dict.setdefault(param_value, [])
            result_dict[param_value].append(element)
        return result_dict


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    cur_sheet = revit.active_view
    if not isinstance(cur_sheet, ViewSheet):
        forms.alert('Необходимо открыть лист.', exitscript=True)

    selection_filter = SelectionFilter()
    selected = uidoc.Selection.PickObject(ObjectType.Element, selection_filter, "Выберите ведомость")

    if selected:
        schedule_vp = doc.GetElement(selected)
        schedule_id = schedule_vp.ScheduleId
        schedule = doc.GetElement(schedule_id)

        parts_schedule = PartsSchedule(schedule)
        if parts_schedule.check_schedule():
            parts_schedule.update_elements()
            parts_schedule.update_filters()
        else:
            script.exit()


script_execute()
