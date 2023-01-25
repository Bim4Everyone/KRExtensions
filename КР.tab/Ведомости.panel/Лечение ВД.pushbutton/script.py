# -*- coding: utf-8 -*-
import clr

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

from pyrevit import EXEC_PARAMS, revit
from pyrevit import forms

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

    def check_schedule(self):
        if not self.__get_field_by_name(self.ifc_img_param_name):
            return False
        if not self.__get_field_by_name(self.excl_param_name):
            return False
        if not self.__get_field_by_name(self.num_param_name):
            return False
        if not self.__get_field_by_name(self.posit_param_name):
            return False
        return True

    def update_elements(self):
        with revit.Transaction("BIM: Обновить ведомости"):
            dict_by_number = self.__group_elements()
            for key in dict_by_number.keys():
                for key2 in dict_by_number[key].keys():
                    for key3 in dict_by_number[key][key2].keys():
                        elements = dict_by_number[key][key2][key3]
                        for element in elements:
                            element.LookupParameter(self.excl_param_name).Set(1)
                        elements[0].LookupParameter(self.excl_param_name).Set(0)

    def update_filters(self):
        field_id = self.__get_field_by_name(self.excl_param_name).FieldId
        filter_type = ScheduleFilterType.NotEqual
        filter_value = 1
        schedule_filter = ScheduleFilter(field_id, filter_type, filter_value)
        if self.__check_filters(field_id, filter_type, filter_value):
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

    def __get_param_value(self, element, name, is_inst=True):
        if not is_inst:
            type_id = element.GetTypeId()
            element = self.doc.GetElement(type_id)
        return element.LookupParameter(name).AsString()

    def __check_filters(self, field_id, filter_type, value):
        filters = self.schedule.Definition.GetFilters()
        used_filters = [x for x in filters if x.FieldId == field_id]
        if used_filters:
            if len(used_filters) == 1:
                if used_filters[0].FilterType == filter_type:
                    if used_filters[0].IsIntegerValue:
                        if used_filters[0].GetIntegerValue == value:
                            return False
                forms.alert("В выбранной ведомости есть неверный фильтр", exitscript=True)
            else:
                forms.alert("В выбранной ведомости уже есть фильтры", exitscript=True)
        return True

    def __group_elements(self):
        dict_by_number = self.__create_dict_by_param(self.elements, self.num_param_name, False)
        for key in dict_by_number.keys():
            dict_by_prior = self.__create_dict_by_param(dict_by_number[key], self.posit_param_name, True)
            dict_by_number[key] = dict_by_prior

        for key_num in dict_by_number.keys():
            for key_prior in dict_by_number[key_num].keys():
                dict_by_family = self.__create_dict_by_param(dict_by_number[key_num][key_prior], "Семейство", True)
                dict_by_number[key_num][key_prior] = dict_by_family

        for key_num in dict_by_number.keys():
            for key_prior in dict_by_number[key_num].keys():
                if len(dict_by_number[key_num][key_prior].keys()) > 1:
                    random_key = dict_by_number[key_num][key_prior].keys()[0]
                    del dict_by_number[key_num][key_prior][random_key]
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
            forms.alert("В выбранной ведомости отсутствуют необходимые параметры", exitscript=True)


script_execute()
