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

# from pyrevit.revit import Transaction

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application


class RevitRepository:
    """
    Класс для получения всего бетона и арматуры из проекта.
    """

    def __init__(self, doc):
        self.doc = doc

        self.categories = []
        self.type_key_word = []
        self.quality_indexes = []

        self.__rebar = self.__get_all_rebar()
        self.__concrete = self.__get_all_concrete()
        self.__concrete_by_table_type = []
        self.__buildings = []
        self.__construction_sections = []

    def set_table_type(self, table_type):
        self.categories = table_type.categories
        self.type_key_word = table_type.type_key_word
        self.quality_indexes = table_type.indexes_info

        self.__get_concrete_by_table_type()
        self.__buildings = self.__get_buildings()
        self.__construction_sections = self.__get_construction_sections()

    def check_exist_main_parameters(self):
        errors_dict = dict()
        common_parameters = ["обр_ФОП_Раздел проекта",
                             "ФОП_Секция СМР",
                             "обр_ФОП_Фильтрация 1",
                             "обр_ФОП_Фильтрация 2"]
        concrete_inst_parameters = ["Объем"]
        rebar_inst_type_parameters = ["обр_ФОП_Форма_номер"]
        concrete_common_parameters = common_parameters + concrete_inst_parameters

        for element in self.__rebar:
            element_type = self.doc.GetElement(element.GetTypeId())
            for parameter_name in common_parameters:
                if not element.IsExistsParam(parameter_name):
                    key = "Арматура___Отсутствует параметр у экземпляра___" + parameter_name
                    errors_dict.setdefault(key, [])
                    errors_dict[key].append(str(element.Id))

            for parameter_name in rebar_inst_type_parameters:
                if not element.IsExistsParam(parameter_name) and not element_type.IsExistsParam(parameter_name):
                    key = "Арматура___Отсутствует параметр у экземпляра или типоразмера___" + parameter_name
                    errors_dict.setdefault(key, [])
                    errors_dict[key].append(str(element.Id))

        for element in self.__concrete:
            for parameter_name in concrete_common_parameters:
                if not element.IsExistsParam(parameter_name):
                    key = "Железобетон___Отсутствует параметр у экземпляра___" + parameter_name
                    errors_dict.setdefault(key, [])
                    errors_dict[key].append(str(element.Id))

        if errors_dict:
            missing_parameters = self.__create_error_list(errors_dict)
            return missing_parameters

    def check_exist_rebar_parameters(self):
        errors_dict = dict()
        rebar_inst_parameters = ["обр_ФОП_Длина",
                                 "обр_ФОП_Группа КР",
                                 "обр_ФОП_Количество типовых на этаже",
                                 "обр_ФОП_Количество типовых этажей"]
        rebar_type_parameters = ["мод_ФОП_IFC семейство"]
        rebar_inst_type_parameters = ["мод_ФОП_Диаметр",
                                      "обр_ФОП_Форма_номер"]

        for element in self.__rebar:
            element_type = self.doc.GetElement(element.GetTypeId())
            for parameter_name in rebar_inst_parameters:
                if not element.IsExistsParam(parameter_name):
                    key = "Арматура___Отсутствует параметр у экземпляра___" + parameter_name
                    errors_dict.setdefault(key, [])
                    errors_dict[key].append(str(element.Id))

            for parameter_name in rebar_type_parameters:
                if not element_type.IsExistsParam(parameter_name):
                    key = "Арматура___Отсутствует параметр у типоразмера___" + parameter_name
                    errors_dict.setdefault(key, [])
                    errors_dict[key].append(str(element.Id))

            for parameter_name in rebar_inst_type_parameters:
                if not element.IsExistsParam(parameter_name) and not element_type.IsExistsParam(parameter_name):
                    key = "Арматура___Отсутствует параметр у экземпляра или типоразмера___" + parameter_name
                    errors_dict.setdefault(key, [])
                    errors_dict[key].append(str(element.Id))

            if not element.IsExistsParam("обр_ФОП_Количество") and not element.IsExistsParam("Количество"):
                key = "Арматура___Отсутствует параметр___Количество (для IFC - 'обр_ФОП_Количество')"
                errors_dict.setdefault(key, [])
                errors_dict[key].append(str(element.Id))

        if errors_dict:
            missing_parameters = self.__create_error_list(errors_dict)
            return missing_parameters

    def check_parameters_values(self):
        errors_dict = dict()
        concrete_inst_parameters = ["Объем"]
        # Исправить параметр обр_ФОП_Длина (мод_ФОП_)
        rebar_inst_parameters = ["обр_ФОП_Группа КР",
                                 "обр_ФОП_Количество типовых на этаже",
                                 "обр_ФОП_Количество типовых этажей"]
        rebar_inst_type_parameters = ["мод_ФОП_Диаметр"]

        for element in self.__rebar:
            element_type = self.doc.GetElement(element.GetTypeId())
            for parameter_name in rebar_inst_parameters:
                if not element.GetParam(parameter_name).HasValue:
                    key = "Арматура___Отсутствует значение у параметра___" + parameter_name
                    errors_dict.setdefault(key, [])
                    errors_dict[key].append(str(element.Id))

            for parameter_name in rebar_inst_type_parameters:
                if element.IsExistsParam(parameter_name):
                    if not element.GetParam(parameter_name).HasValue:
                        key = "Арматура___Отсутствует значение у параметра (экземпляра или типа)___" + parameter_name
                        errors_dict.setdefault(key, [])
                        errors_dict[key].append(str(element.Id))
                else:
                    if not element_type.GetParam(parameter_name).HasValue:
                        key = "Арматура___Отсутствует значение у параметра (экземпляра или типа)___" + parameter_name
                        errors_dict.setdefault(key, [])
                        errors_dict[key].append(str(element.Id))

            if element_type.GetParamValue("мод_ФОП_IFC семейство"):
                if not element.GetParam("обр_ФОП_Количество").HasValue:
                    key = "Арматура___Отсутствует значение у параметра___обр_ФОП_Количество"
                    errors_dict.setdefault(key, [])
                    errors_dict[key].append(str(element.Id))
            else:
                if not element.GetParam("Количество").HasValue:
                    key = "Арматура___Отсутствует значение у параметра___Количество"
                    errors_dict.setdefault(key, [])
                    errors_dict[key].append(str(element.Id))

        for element in self.__concrete:
            for parameter_name in concrete_inst_parameters:
                if element.GetParamValue(parameter_name) == 0:
                    key = "Железобетон___Отсутствует значение у параметра___" + parameter_name
                    errors_dict.setdefault(key, [])
                    errors_dict[key].append(str(element.Id))

        if errors_dict:
            empty_parameters = self.__create_error_list(errors_dict)
            return empty_parameters

    def filter_by_main_parameters(self):
        filter_param_1 = "обр_ФОП_Фильтрация 1"
        filter_param_2 = "обр_ФОП_Фильтрация 2"
        filter_param_3 = "обр_ФОП_Форма_номер"
        filter_value = "Исключить из показателей качества"

        self.__rebar = self.__filter_by_param(self.__rebar, filter_param_1, filter_value)
        self.__rebar = self.__filter_by_param(self.__rebar, filter_param_2, filter_value)
        self.__rebar = self.__filter_by_param(self.__rebar, filter_param_3, 1000)

        self.__concrete = self.__filter_by_param(self.__concrete, filter_param_1, filter_value)
        self.__concrete = self.__filter_by_param(self.__concrete, filter_param_2, filter_value)

    def __get_all_concrete(self):
        """
        Получение из проекта всех экземпляров семейств категорий железобетона
        """
        all_categories = []
        all_categories.append(Category.GetCategory(doc, BuiltInCategory.OST_Walls))
        all_categories.append(Category.GetCategory(doc, BuiltInCategory.OST_StructuralColumns))
        all_categories.append(Category.GetCategory(doc, BuiltInCategory.OST_StructuralFoundation))
        all_categories.append(Category.GetCategory(doc, BuiltInCategory.OST_Floors))
        all_categories.append(Category.GetCategory(doc, BuiltInCategory.OST_StructuralFraming))
        elements = self.__collect_elements_by_categories(all_categories)
        return elements

    def __get_concrete_by_table_type(self):
        """
        Получение из проекта всех экземпляров семейств
        категорий железобетона по выбранному типу таблицы
        """
        categories_id = [x.Id for x in self.categories]
        filtered_concrete = [x for x in self.__concrete if x.Category.Id in categories_id]
        self.__concrete_by_table_type = self.__filter_by_type(filtered_concrete)

    def get_filtered_concrete_by_user(self, buildings, constr_sections):
        buildings = [x.text_value for x in buildings if x.is_checked]
        constr_sections = [x.text_value for x in constr_sections if x.is_checked]
        filtered_elements = []
        for element in self.concrete:
            if element.GetParamValue("ФОП_Секция СМР") in buildings:
                if element.GetParamValue("обр_ФОП_Раздел проекта") in constr_sections:
                    filtered_elements.append(element)

        return filtered_elements

    def __get_all_rebar(self):
        elements = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_Rebar)
        elements.WhereElementIsNotElementType().ToElements()
        return elements

    def get_filtered_rebar_by_user(self, buildings, constr_sections):
        rebar_by_table_type = []
        rebar_group_values = [x.rebar_group for x in self.quality_indexes if x.index_type == "mass"]
        for value in rebar_group_values:
            rebar_by_table_type += self.__filter_by_param(self.rebar, "обр_ФОП_Группа КР", value)
        buildings = [x.text_value for x in buildings if x.is_checked]
        constr_sections = [x.text_value for x in constr_sections if x.is_checked]
        filtered_elements = []
        for element in rebar_by_table_type:
            if element.GetParamValue("ФОП_Секция СМР") in buildings:
                if element.GetParamValue("обр_ФОП_Раздел проекта") in constr_sections:
                    filtered_elements.append(element)
        return filtered_elements

    def __collect_elements_by_categories(self, categories):
        """
        Получение экземпляров семейств по списку категорий.
        """
        cat_filters = [ElementCategoryFilter(x.Id) for x in categories]
        cat_filters_typed = List[ElementFilter](cat_filters)
        logical_or_filter = LogicalOrFilter(cat_filters_typed)
        elements = FilteredElementCollector(self.doc).WherePasses(logical_or_filter)
        elements.WhereElementIsNotElementType().ToElements()
        return elements

    def __filter_by_param(self, elements, param_name, value):
        filtered_list = []
        for element in elements:
            if element.IsExistsParam(param_name):
                param_value = element.GetParamValue(param_name)
            else:
                element_type = self.doc.GetElement(element.GetTypeId())
                param_value = element_type.GetParamValueOrDefault(param_name, 0)

            if param_value != value:
                filtered_list.append(element)
        return filtered_list

    def __filter_by_type(self, elements):
        """
        Фильтрация семейств по имени типоразмера.
        В имени типоразмера должно присутствовать ключевое слово.
        """
        filtered_list = []
        for element in elements:
            # Проверить получение типоразмера по Name
            if self.type_key_word in element.Name:
                filtered_list.append(element)

        return filtered_list

    def __create_param_set(self, elements, param_name):
        set_of_values = set()
        for element in elements:
            set_of_values.add(element.GetParamValue(param_name))
        return sorted(set_of_values)

    def __create_error_list(self, errors_dict):
        missing_parameters = []
        for error in errors_dict.keys():
            error_info = []
            for word in error.split("___"):
                error_info.append(word)
            error_info.append(", ".join(errors_dict[error]))
            missing_parameters.append(error_info)
        # pyrevit не выводит таблицу из трех строк
        if len(missing_parameters) == 3:
            missing_parameters.append(["_", "_", "_", "_"])
        return missing_parameters

    def __get_buildings(self):
        buildings = self.__create_param_set(self.concrete, "ФОП_Секция СМР")
        result_buildings = []
        for building in buildings:
            result_buildings.append(ElementSection(building))
        return result_buildings

    def __get_construction_sections(self):
        sections = self.__create_param_set(self.concrete, "обр_ФОП_Раздел проекта")
        result_sections = []
        for section in sections:
            result_sections.append(ElementSection(section))
        return result_sections

    @reactive
    def concrete(self):
        return self.__concrete_by_table_type

    @reactive
    def rebar(self):
        return self.__rebar

    @reactive
    def buildings(self):
        return self.__buildings

    @reactive
    def construction_sections(self):
        return self.__construction_sections


class TableType:
    def __init__(self, name):
        self.__name = name
        self.__type_key_word = ""
        self.__categories = []
        self.__indexes_info = []

    @reactive
    def name(self):
        return self.__name

    @name.setter
    def name(self, value):
        self.__name = value

    @reactive
    def type_key_word(self):
        return self.__type_key_word

    @type_key_word.setter
    def type_key_word(self, value):
        self.__type_key_word = value

    @reactive
    def categories(self):
        return self.__categories

    @categories.setter
    def categories(self, value):
        self.__categories = value

    @reactive
    def indexes_info(self):
        return self.__indexes_info

    @indexes_info.setter
    def indexes_info(self, value):
        self.__indexes_info = value


class QualityIndex:
    def __init__(self, name, number, index_type="", rebar_group=""):
        self.__name = name
        self.__number = number
        # self.__number_in_table = number_in_table
        self.__index_type = index_type
        self.__rebar_group = rebar_group

    @reactive
    def name(self):
        return self.__name

    @name.setter
    def name(self, value):
        self.__name = value

    @reactive
    def number(self):
        return self.__number

    @number.setter
    def number(self, value):
        self.__number = value

    @reactive
    def number_in_table(self):
        return self.__number_in_table

    @number_in_table.setter
    def number_in_table(self, value):
        self.__number_in_table = value

    @reactive
    def index_type(self):
        return self.__index_type

    @index_type.setter
    def index_type(self, value):
        self.__index_type = value

    @reactive
    def rebar_group(self):
        return self.__rebar_group

    @rebar_group.setter
    def rebar_group(self, value):
        self.__rebar_group = value


class ElementSection:
    def __init__(self, text_value):
        self.__text_value = text_value
        self.__number = text_value
        self.__is_checked = False

    @reactive
    def number(self):
        if self.__number:
            return self.__number
        else:
            return "<Параметр не заполнен>"

    @number.setter
    def number(self, value):
        self.__number = value

    @reactive
    def text_value(self):
        return self.__text_value

    @text_value.setter
    def text_value(self, value):
        self.__text_value = value

    @reactive
    def is_checked(self):
        return self.__is_checked

    @is_checked.setter
    def is_checked(self, value):
        self.__is_checked = value


class Construction:
    def __init__(self, table_type, concrete_elements, rebar_elements):
        self.table_type = table_type
        self.concrete = concrete_elements
        self.rebar = rebar_elements

        self.__quality_indexes = dict()
        self.__rebar_by_function = dict()
        self.__rebar_mass_by_function = dict()

        self.__diameter_dict = {
            4: 0.098,
            5: 0.144,
            6: 0.222,
            7: 0.302,
            8: 0.395,
            9: 0.499,
            10: 0.617,
            12: 0.888,
            14: 1.208,
            16: 1.578,
            18: 1.998,
            20: 2.466,
            22: 2.984,
            25: 3.853,
            28: 4.834,
            32: 6.313,
            36: 7.990,
            40: 9.805
        }

        self.__concrete_volume = 0
        self.__calculate_concrete_volume()

        self.__group_rebar_by_function()
        self.__calculate_rebar_group_mass()

        self.__set_building_info()
        self.__set_elements_sizes()
        self.__set_concrete_class()
        self.__set_concrete_volume()
        self.__set_rebar_indexes()

    def __calculate_rebar_element_mass(self, elements):
        rebar_mass = 0
        for element in elements:
            element_type = doc.GetElement(element.GetTypeId())
            is_ifc_element = element_type.GetParamValue("мод_ФОП_IFC семейство")
            if element.IsExistsParam("мод_ФОП_Диаметр"):
                diameter_param = element.GetParam("мод_ФОП_Диаметр")
            else:
                diameter_param = element_type.GetParam("мод_ФОП_Диаметр")
            diameter = convert_value(diameter_param)
            mass_per_metr = self.__diameter_dict[diameter]

            length_param = element.GetParam("обр_ФОП_Длина")
            length = convert_value(length_param) * 0.001

            if is_ifc_element:
                amount = element.GetParamValue("обр_ФОП_Количество")
            else:
                amount = element.GetParamValue("Количество")
            amount_on_level = element.GetParamValue("обр_ФОП_Количество типовых на этаже")
            levels_amount = element.GetParamValue("обр_ФОП_Количество типовых этажей")

            element_mass = mass_per_metr * length * amount * amount_on_level * levels_amount
            rebar_mass += element_mass

        return rebar_mass

    def __calculate_rebar_group_mass(self):
        for key in self.__rebar_by_function.keys():
            elements = self.__rebar_by_function[key]
            rebar_mass = self.__calculate_rebar_element_mass(elements)
            self.__rebar_mass_by_function[key] = rebar_mass

    def __calculate_concrete_volume(self):
        for element in self.concrete:
            volume_param = element.GetParam("Объем")
            volume = convert_value(volume_param)
            self.__concrete_volume += volume

    def __group_rebar_by_function(self):
        for element in self.rebar:
            rebar_function = element.LookupParameter("обр_ФОП_Группа КР").AsString()
            self.__rebar_by_function.setdefault(rebar_function, [])
            self.__rebar_by_function[rebar_function].append(element)

    def __set_building_info(self):
        project_info = FilteredElementCollector(doc).OfClass(ProjectInfo).FirstElement()
        value = project_info.GetParamValue("Наименование проекта")
        self.__quality_indexes["Этажность здания, тип секции"] = value

    def __set_elements_sizes(self):
        elements_sizes = set()
        walls_category = Category.GetCategory(doc, BuiltInCategory.OST_Walls)
        columns_category = Category.GetCategory(doc, BuiltInCategory.OST_StructuralColumns)
        floors_category = Category.GetCategory(doc, BuiltInCategory.OST_Floors)
        if self.table_type.name == "Пилоны":
            for element in self.concrete:
                element_type = doc.GetElement(element.GetTypeId())
                if element.Category.Id == columns_category.Id:
                    if element_type.IsExistsParam("ADSK_Размер_Высота") and element_type.IsExistsParam(
                            "ADSK_Размер_Ширина"):
                        height = element_type.GetParam("ADSK_Размер_Высота").AsValueString()
                        width = element_type.GetParam("ADSK_Размер_Ширина").AsValueString()
                        size = height + "х" + width
                        elements_sizes.add(size)
                if element.Category.Id == walls_category.Id:
                    if element_type.IsExistsParam("Толщина") and element.IsExistsParam("Длина"):
                        height = element.GetParam("Длина").AsValueString()
                        width = element_type.GetParam("Толщина").AsValueString()
                        size = height + "х" + width
                        elements_sizes.add(size)
            self.__quality_indexes["Сечение пилонов, толщина х ширина, мм"] = ", ".join(elements_sizes)

            self.__quality_indexes[
                "Коэффициент суммарной площади сечений пилонов от площади перекрытия, ΣAw/Ap х 100"] = 0

        if self.table_type.name == "Стены":
            for element in self.concrete:
                element_type = doc.GetElement(element.GetTypeId())
                if element_type.IsExistsParam("Толщина"):
                    width = element_type.GetParam("Толщина").AsValueString()
                    elements_sizes.add(width)
            self.__quality_indexes["Толщина стен, мм"] = ", ".join(elements_sizes)

        if self.table_type.name == "Фундаментная плита":
            for element in self.concrete:
                element_type = doc.GetElement(element.GetTypeId())
                if element_type.IsExistsParam("Толщина"):
                    width = element_type.GetParam("Толщина").AsValueString()
                    elements_sizes.add(width)
            self.__quality_indexes["Толщина плиты, мм"] = ", ".join(elements_sizes)

        if self.table_type.name == "Плита перекрытия":
            for element in self.concrete:
                element_type = doc.GetElement(element.GetTypeId())
                if element.Category.Id == floors_category.Id:
                    if element_type.IsExistsParam("Толщина"):
                        width = element_type.GetParam("Толщина").AsValueString()
                        elements_sizes.add(width)
            self.__quality_indexes["Толщина плиты, мм"] = ", ".join(elements_sizes)

    def __set_concrete_class(self):
        concrete_classes = set()
        for element in self.concrete:
            element_type = doc.GetElement(element.GetTypeId())
            if element_type.IsExistsParam("обр_ФОП_Марка бетона B"):
                value = element_type.GetParam("обр_ФОП_Марка бетона B").AsValueString()
                concrete_class = "В" + value
                concrete_classes.add(concrete_class)
        self.__quality_indexes["Класс бетона"] = ", ".join(concrete_classes)

    def __set_concrete_volume(self):
        self.__quality_indexes["Объем бетона, м3"] = self.__concrete_volume

    def __set_rebar_indexes(self):
        full_consumption = 0
        for index_info in self.table_type.indexes_info:
            if index_info.index_type == "mass" or index_info.index_type == "consumption":
                rebar_function = index_info.rebar_group
                rebar_mass = 0
                if rebar_function in self.__rebar_mass_by_function.keys():
                    rebar_mass = self.__rebar_mass_by_function[rebar_function]
                if index_info.index_type == "mass":
                    self.__quality_indexes[index_info.name] = rebar_mass
                elif index_info.index_type == "consumption":
                    consumption = rebar_mass / self.__concrete_volume
                    full_consumption += consumption
                    self.__quality_indexes[index_info.name] = consumption
        self.__quality_indexes["Общий расход, кг/м3"] = full_consumption

    @reactive
    def quality_indexes(self):
        return self.__quality_indexes


class QualityTable:
    def __init__(self, table_type, construction):
        self.table_type = table_type
        self.table_name = table_type.name
        self.schedule_name = "РД_Показатели качества_" + table_type.name
        self.indexes_info = table_type.indexes_info
        self.indexes_values = construction.quality_indexes

        self.table_width = 186
        self.table_width_column_1 = 10
        self.table_width_column_2 = 140
        self.table_width_column_3 = 35
        self.row_height = 8

    def create_table(self):
        schedule = self.find_schedule()
        if schedule:
            self.update_schedule_name(schedule)
            schedule = self.create_new_schedule(self.table_width)
            self.set_schedule_row_values(schedule)
        else:
            schedule = self.create_new_schedule(self.table_width)
            self.set_schedule_row_values(schedule)

        # output = script.get_output()
        # data = self.set_window_row_values()
        # data.append(["_", "_", "_"])
        # output.print_table(table_data=data,
        #                    title="Показатели качества",
        #                    columns=["Номер", "Название", "Значение"])

    def find_schedule(self):
        schedules = FilteredElementCollector(doc).OfClass(ViewSchedule)

        fvp = ParameterValueProvider(ElementId(BuiltInParameter.VIEW_NAME))
        rule = FilterStringEquals()
        case_sens = False
        filter_rule = FilterStringRule(fvp, rule, self.schedule_name, case_sens)
        name_filter = ElementParameterFilter(filter_rule)
        schedules.WherePasses(name_filter)

        return schedules.FirstElement()

    def update_schedule_name(self, schedule):
        old_name = schedule.Name
        time = datetime.datetime.now().strftime("%Y-%m-%d %H.%M.%S")
        with revit.Transaction("BIM: Создание таблицы УПК"):
            schedule.Name = old_name + "_" + time

    def create_new_schedule(self, width):
        category_id = ElementId(BuiltInCategory.OST_Walls)
        with revit.Transaction("BIM: Создание таблицы УПК"):
            new_schedule = ViewSchedule.CreateSchedule(doc, category_id)
            new_schedule.Name = self.schedule_name

            s_definition = new_schedule.Definition
            s_definition.ShowHeaders = False
            table_data = new_schedule.GetTableData()
            header_data = table_data.GetSectionData(SectionType.Header)
            header_data.RemoveRow(0)

            field_type = ScheduleFieldType.Instance
            param_id = ElementId(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            s_field = s_definition.AddField(field_type, param_id)
            filter_value = "schedule_for_header"
            filter_rule = ScheduleFilterType.Equal
            my_filter = ScheduleFilter(s_field.FieldId, filter_rule, filter_value)
            s_definition.AddFilter(my_filter)

            s_field.SheetColumnWidth = convert_length(width)
        return new_schedule

    def options_cell_font(self, alignment=["left", "center", "right"]):
        tcs = TableCellStyle()
        options = TableCellStyleOverrideOptions()
        options.HorizontalAlignment = True
        options.Font = True
        options.FontSize = True
        tcs.FontName = "GOST Common"
        tcs.TextSize = 9.4482237
        tcs.SetCellStyleOverrideOptions(options)
        if alignment == "left":
            tcs.FontHorizontalAlignment = HorizontalAlignmentStyle.Left
        elif alignment == "central":
            tcs.FontHorizontalAlignment = HorizontalAlignmentStyle.Center
        elif alignment == "right":
            tcs.FontHorizontalAlignment = HorizontalAlignmentStyle.Right
        return tcs

    def set_schedule_row_values(self, schedule):
        table_data = schedule.GetTableData()
        header_data = table_data.GetSectionData(SectionType.Header)
        with revit.Transaction("BIM: Заполнение таблицы УПК"):
            header_data.InsertColumn(0)
            header_data.InsertColumn(1)
            header_data.SetColumnWidth(0, convert_length(self.table_width_column_1))
            header_data.SetColumnWidth(1, convert_length(self.table_width_column_2))
            header_data.SetColumnWidth(2, convert_length(self.table_width_column_3))

            header_data.InsertRow(0)
            height = 15
            header_data.SetRowHeight(0, convert_length(height))
            value1 = "№"
            header_data.SetCellStyle(0, 0, self.options_cell_font("central"))
            header_data.SetCellText(0, 0, value1)

            value2 = "Анализируемый параметр"
            header_data.SetCellStyle(0, 1, self.options_cell_font("central"))
            header_data.SetCellText(0, 1, value2)

            value3 = "Значения"
            header_data.SetCellStyle(0, 2, self.options_cell_font("central"))
            header_data.SetCellText(0, 2, value3)

            for i, quality_index in enumerate(self.indexes_info):
                header_data.InsertRow(i + 1)
                header_data.SetRowHeight(i + 1, convert_length(self.row_height))

                index_number = quality_index.number
                header_data.SetCellStyle(i + 1, 0, self.options_cell_font("central"))
                header_data.SetCellText(i + 1, 0, index_number)

                index_name = quality_index.name
                header_data.SetCellStyle(i + 1, 1, self.options_cell_font("left"))
                header_data.SetCellText(i + 1, 1, index_name)

                index_value = str(self.indexes_values[quality_index.name])
                header_data.SetCellStyle(i + 1, 2, self.options_cell_font("central"))
                header_data.SetCellText(i + 1, 2, index_value)

    # def set_window_row_values(self):
    #     output_rows = []
    #     for index in self.indexes_info:
    #         output_row = []
    #         output_row.append(index.number)
    #         output_row.append(index.name)
    #         output_row.append(self.indexes_values[index.name])
    #         output_rows.append(output_row)
    #     return output_rows


class CreateQualityTableCommand(ICommand):
    CanExecuteChanged, _canExecuteChanged = pyevent.make_event()

    def __init__(self, view_model, *args):
        ICommand.__init__(self, *args)
        self.__view_model = view_model
        self.__view_model.PropertyChanged += self.ViewModel_PropertyChanged

    def add_CanExecuteChanged(self, value):
        self.CanExecuteChanged += value

    def remove_CanExecuteChanged(self, value):
        self.CanExecuteChanged -= value

    def OnCanExecuteChanged(self):
        self._canExecuteChanged(self, System.EventArgs.Empty)

    def ViewModel_PropertyChanged(self, sender, e):
        self.OnCanExecuteChanged()

    def CanExecute(self, parameter):
        if not self.__view_model.revit_repository.rebar:
            self.__view_model.error_text = "Арматура не найдена в проекте"
            return False
        if not self.__view_model.buildings:
            self.__view_model.error_text = "ЖБ не найден в проекте"
            return False

        self.__view_model.error_text = None
        return True

    def Execute(self, parameter):
        buildings = self.__view_model.buildings
        construction_sections = self.__view_model.construction_sections
        concrete = self.__view_model.revit_repository.get_filtered_concrete_by_user(buildings, construction_sections)
        rebar = self.__view_model.revit_repository.get_filtered_rebar_by_user(buildings, construction_sections)
        if not concrete:
            alert("Не найден ЖБ для выбранных секций и разделов")
        elif not rebar:
            alert("Не найдена арматура для выбранных секций и разделов")
        else:
            selected_table_type = self.__view_model.selected_table_type
            construction = Construction(selected_table_type, concrete, rebar)
            quality_table = QualityTable(selected_table_type, construction)
            quality_table.create_table()


class MainWindow(WPFWindow):
    def __init__(self):
        self._context = None
        self.xaml_source = op.join(op.dirname(__file__), 'MainWindow.xaml')
        super(MainWindow, self).__init__(self.xaml_source)


class MainWindowViewModel(Reactive):
    def __init__(self, revit_repository, table_types):
        Reactive.__init__(self)
        self.__table_types = table_types
        self.__revit_repository = revit_repository
        self.__buildings = []
        self.__construction_sections = []
        self.__error_text = ""

        self.__create_tables_command = CreateQualityTableCommand(self)

        self.selected_table_type = table_types[0]

    @reactive
    def revit_repository(self):
        return self.__revit_repository

    @reactive
    def table_types(self):
        return self.__table_types

    @reactive
    def selected_table_type(self):
        return self.__selected_table_type

    @selected_table_type.setter
    def selected_table_type(self, value):
        self.revit_repository.set_table_type(value)
        self.buildings = self.__revit_repository.buildings
        self.construction_sections = self.__revit_repository.construction_sections
        self.__selected_table_type = value

    @reactive
    def buildings(self):
        return self.__buildings

    @buildings.setter
    def buildings(self, value):
        self.__buildings = value

    @reactive
    def construction_sections(self):
        return self.__construction_sections

    @construction_sections.setter
    def construction_sections(self, value):
        self.__construction_sections = value

    @reactive
    def error_text(self):
        return self.__error_text

    @error_text.setter
    def error_text(self, value):
        self.__error_text = value

    @property
    def create_tables_command(self):
        return self.__create_tables_command


def convert_value(parameter):
    if parameter.StorageType == StorageType.Double:
        value = parameter.AsDouble()
    if parameter.StorageType == StorageType.Integer:
        value = parameter.AsInteger()
    if parameter.StorageType == StorageType.String:
        value = parameter.AsString()
    if parameter.StorageType == StorageType.ElementId:
        value = parameter.AsValueString()

    if int(app.VersionNumber) > 2021:
        try:
            d_type = parameter.GetUnitTypeId()
            result = UnitUtils.ConvertFromInternalUnits(value, d_type)
        except:
            result = value
    else:
        try:
            d_type = parameter.DisplayUnitType
            result = UnitUtils.ConvertFromInternalUnits(value, d_type)
        except:
            result = value
    return result


def convert_length(value):
    if int(app.VersionNumber) < 2021:
        unit_type = DisplayUnitType.DUT_MILLIMETERS
    else:
        unit_type = UnitTypeId.Millimeters
    converted_value = UnitUtils.ConvertToInternalUnits(value, unit_type)
    return converted_value


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    table_types = []

    walls_cat = Category.GetCategory(doc, BuiltInCategory.OST_Walls)
    columns_cat = Category.GetCategory(doc, BuiltInCategory.OST_StructuralColumns)
    foundation_cat = Category.GetCategory(doc, BuiltInCategory.OST_Doors)
    floor_cat = Category.GetCategory(doc, BuiltInCategory.OST_Floors)
    framing_cat = Category.GetCategory(doc, BuiltInCategory.OST_StructuralFraming)

    walls_table_type = TableType("Стены")
    walls_table_type.categories = [walls_cat]
    walls_table_type.type_key_word = "Стена"
    walls_table_type.indexes_info = [
        QualityIndex("Этажность здания, тип секции", "1"),
        QualityIndex("Толщина стен, мм", "2"),
        QualityIndex("Класс бетона", "3"),
        QualityIndex("Объем бетона, м3", "4"),
        QualityIndex("Масса вертикальной арматуры, кг", "5.1", "mass", "Стены_Вертикальная"),
        QualityIndex("Расход вертикальной арматуры, кг/м3", "5.2", "consumption", "Стены_Вертикальная"),
        QualityIndex("Масса горизонтальной арматуры, кг", "6.1", "mass", "Стены_Горизонтальная"),
        QualityIndex("Расход горизонтальной арматуры, кг/м3", "6.2", "consumption", "Стены_Горизонтальная"),
        QualityIndex("Масса конструктивной арматуры, кг", "7.1", "mass", "Стены_Конструктивная"),
        QualityIndex("Расход конструктивной арматуры, кг/м3", "7.2", "consumption", "Стены_Конструктивная"),
        QualityIndex("Общий расход, кг/м3", "8")]

    columns_table_type = TableType("Пилоны")
    columns_table_type.categories = [walls_cat, columns_cat]
    columns_table_type.type_key_word = "Пилон"
    columns_table_type.indexes_info = [
        QualityIndex("Этажность здания, тип секции", "1"),
        QualityIndex("Сечение пилонов, толщина х ширина, мм", "2"),
        QualityIndex("Класс бетона", "3"),
        QualityIndex("Объем бетона, м3", "4"),
        QualityIndex("Коэффициент суммарной площади сечений пилонов от площади перекрытия, ΣAw/Ap х 100", "5"),
        QualityIndex("Масса продольной арматуры, кг", "6.1", "mass", "Пилоны_Продольная"),
        QualityIndex("Расход продольной арматуры, кг/м3", "6.2", "consumption", "Пилоны_Продольная"),
        QualityIndex("Масса поперечной арматуры, кг", "7.1", "mass", "Пилоны_Поперечная"),
        QualityIndex("Расход поперечной арматуры, кг/м3", "7.2", "consumption", "Пилоны_Поперечная"),
        QualityIndex("Общий расход, кг/м3", "8")]

    foundation_table_type = TableType("Фундаментная плита")
    foundation_table_type.categories = [foundation_cat]
    foundation_table_type.type_key_word = "ФПлита"
    foundation_table_type.indexes_info = [
        QualityIndex("Этажность здания, тип секции", "1"),
        QualityIndex("Толщина плиты, мм", "2"),
        QualityIndex("Класс бетона", "3"),
        QualityIndex("Объем бетона, м3", "4"),
        QualityIndex("Масса нижней фоновой арматуры, кг", "5.1", "mass", "ФП_Фон_Н"),
        QualityIndex("Расход нижней фоновой арматуры, кг/м3", "5.2", "consumption", "ФП_Фон_Н"),
        QualityIndex("Масса нижней арматуры усиления, кг", "5.3", "mass", "ФП_Усиление_Н"),
        QualityIndex("Расход нижней арматуры усиления, кг/м3", "5.4", "consumption", "ФП_Усиление_Н"),
        QualityIndex("Масса верхней фоновой арматуры, кг", "6.1", "mass", "ФП_Фон_В"),
        QualityIndex("Расход верхней фоновой арматуры, кг/м3", "6.2", "consumption", "ФП_Фон_В"),
        QualityIndex("Масса верхней арматуры усиления, кг", "6.3", "mass", "ФП_Усиление_В"),
        QualityIndex("Расход верхней арматуры усиления, кг/м3", "6.4", "consumption", "ФП_Усиление_В"),
        QualityIndex("Масса поперечной арматуры в зонах продавливания, кг", "7.1", "mass", "ФП_Каркасы_Продавливание"),
        QualityIndex("Расход поперечной арматуры в зонах продавливания, кг/м3", "7.2", "consumption", "ФП_Каркасы_Продавливание"),
        QualityIndex("Масса конструктивной арматуры, кг", "7.3", "mass", "ФП_Конструктивная"),
        QualityIndex("Расход конструктивной арматуры, кг/м3", "7.4", "consumption", "ФП_Конструктивная"),
        QualityIndex("Масса выпусков, кг", "8.1", "mass", "ФП_Выпуски"),
        QualityIndex("Расход выпусков, кг/м3", "8.2", "consumption", "ФП_Выпуски"),
        QualityIndex("Масса закладных, кг", "9.1", "mass", "ФП_Закладная "),
        QualityIndex("Расход закладных, кг/м3", "9.2", "consumption", "ФП_Закладная "),
        QualityIndex("Общий расход, кг/м3", "10")]

    floor_table_type = TableType("Плита перекрытия")
    floor_table_type.categories = [floor_cat, framing_cat, walls_cat]
    floor_table_type.type_key_word = "ФПлита"
    floor_table_type.indexes_info = [
        QualityIndex("Этажность здания, тип секции", "1"),
        QualityIndex("Толщина плиты, мм", "2"),
        QualityIndex("Класс бетона", "3"),
        QualityIndex("Объем бетона, м3", "4"),
        QualityIndex("Масса нижней фоновой арматуры, кг", "5.1", "mass", "ПП_Фон_Н"),
        QualityIndex("Расход нижней фоновой арматуры, кг/м3", "5.2", "consumption", "ПП_Фон_Н"),
        QualityIndex("Масса нижней арматуры усиления, кг", "5.3", "mass", "ПП_Усиление_Н"),
        QualityIndex("Расход нижней арматуры усиления, кг/м3", "5.4", "consumption", "ПП_Усиление_Н"),
        QualityIndex("Масса верхней фоновой арматуры, кг", "6.1", "mass", "ПП_Фон_В"),
        QualityIndex("Расход верхней фоновой арматуры, кг/м3", "6.2", "consumption", "ПП_Фон_В"),
        QualityIndex("Масса верхней арматуры усиления, кг", "6.3", "mass", "ПП_Усиление_В"),
        QualityIndex("Расход верхней арматуры усиления, кг/м3", "6.4", "consumption", "ПП_Усиление_В"),
        QualityIndex("Масса поперечной арматуры в зонах продавливания, кг", "7.1", "mass", "ПП_Каркасы_Продавливание"),
        QualityIndex("Расход поперечной арматуры в зонах продавливания, кг/м3", "7.2", "consumption", "ПП_Каркасы_Продавливание"),
        QualityIndex("Масса конструктивной арматуры, кг", "8.1", "mass", "ПП_Конструктивная"),
        QualityIndex("Расход конструктивной арматуры, кг/м3", "8.2", "consumption", "ПП_Конструктивная"),
        QualityIndex("Масса арматуры балок, кг", "9.1", "mass", "ПП_Балки"),
        QualityIndex("Расход арматуры балок, кг/м3", "9.2", "consumption", "ПП_Балки"),
        QualityIndex("Общий расход, кг/м3", "10")]

    table_types.append(walls_table_type)
    table_types.append(columns_table_type)
    table_types.append(foundation_table_type)
    table_types.append(floor_table_type)

    revit_repository = RevitRepository(doc)
    # check = revit_repository.check_exist_main_parameters()
    # if check:
    #     output = script.get_output()
    #     output.print_table(table_data=check,
    #                        title="Показатели качества",
    #                        columns=["Категории", "Тип ошибки", "Название параметра", "Id"])
    #     script.exit()

    revit_repository.filter_by_main_parameters()

    # check = revit_repository.check_exist_rebar_parameters()
    # if check:
    #     output = script.get_output()
    #     output.print_table(table_data=check,
    #                        title="Показатели качества",
    #                        columns=["Категории", "Тип ошибки", "Название параметра", "Id"])
    #     script.exit()

    # check = revit_repository.check_parameters_values()
    # if check:
    #     output = script.get_output()
    #     output.print_table(table_data=check,
    #                        title="Показатели качества",
    #                        columns=["Категории", "Тип ошибки", "Название параметра", "Id"])
    #     script.exit()

    main_window = MainWindow()
    main_window.DataContext = MainWindowViewModel(revit_repository, table_types)
    main_window.show_dialog()

script_execute()
