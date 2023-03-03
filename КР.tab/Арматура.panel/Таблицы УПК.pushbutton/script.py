# -*- coding: utf-8 -*-
import clr

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
        self.quality_indexes = table_type.quality_indexes

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
        elements_new = []
        for value in self.quality_indexes.values():
            elements_new += self.__filter_by_param(self.rebar, "обр_ФОП_Группа КР", value)
        buildings = [x.text_value for x in buildings if x.is_checked]
        constr_sections = [x.text_value for x in constr_sections if x.is_checked]
        filtered_elements = []
        for element in elements_new:
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
        self.__quality_indexes = dict()

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
    def quality_indexes(self):
        return self.__quality_indexes

    @quality_indexes.setter
    def quality_indexes(self, value):
        self.__quality_indexes = value


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
    """
    Класс для расчетов всех показателей.
    """

    def __init__(self, table_type, concrete_elements, rebar_elements):
        self.table_type = table_type
        self.concrete = concrete_elements
        self.rebar = rebar_elements

        self.__quality_indexes = dict()
        self.__rebar_by_function = dict()
        self.__concrete_volume = 0

        self.__calculate_concrete_volume()
        self.__group_rebar_by_function()

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

        self.__set_building_info()
        self.__set_elements_sizes()
        self.__set_concrete_class()
        self.__set_concrete_volume()
        self.__set_rebar_indexes()

    def __calculate_rebar_mass(self, elements):
        rebar_mass = 0
        for element in elements:
            element_type = doc.GetElement(element.GetTypeId())
            is_ifc_element = element_type.GetParamValue("мод_ФОП_IFC семейство")
            if element.IsExistsParam("мод_ФОП_Диаметр"):
                diameter_param = element.GetParam("мод_ФОП_Диаметр")
            else:
                diameter_param = element_type.GetParam("мод_ФОП_Диаметр")
            diameter = convert_value(app, diameter_param)
            mass_per_metr = self.__diameter_dict[diameter]

            length_param = element.GetParam("обр_ФОП_Длина")
            length = convert_value(app, length_param)

            if is_ifc_element:
                amount = element.GetParamValue("обр_ФОП_Количество")
            else:
                amount = element.GetParamValue("Количество")
            amount_on_level = element.GetParamValue("обр_ФОП_Количество типовых на этаже")
            levels_amount = element.GetParamValue("обр_ФОП_Количество типовых этажей")

            element_mass = mass_per_metr * length * amount * amount_on_level * levels_amount
            rebar_mass += element_mass

        return rebar_mass

    def __calculate_concrete_volume(self):
        for element in self.concrete:
            volume_param = element.GetParam("Объем")
            volume = convert_value(app, volume_param)
            self.__concrete_volume += volume

    def __group_rebar_by_function(self):
        # фильтр арматуры по типу таблицы
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
        if self.table_type.name == "Пилоны":
            walls_category = Category.GetCategory(doc, BuiltInCategory.OST_Walls)
            columns_category = Category.GetCategory(doc, BuiltInCategory.OST_StructuralColumns)
            floors_category = Category.GetCategory(doc, BuiltInCategory.OST_Floors)

            for element in self.concrete:
                element_type = doc.GetElement(element.GetTypeId())
                if element.Category.Id == columns_category.Id:
                    if element_type.IsExistsParam("ADSK_Размер_Высота") and element_type.IsExistsParam("ADSK_Размер_Ширина"):
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

            self.__quality_indexes["Коэффициент суммарной площади сечений пилонов от площади перекрытия, ΣAw/Ap х 100"] = 0

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
        for key in self.table_type.quality_indexes.keys():
            rebar_function = self.table_type.quality_indexes[key]
            elements = self.__rebar_by_function[rebar_function]
            rebar_mass = self.__calculate_rebar_mass(elements)
            if "Масса" in key:
                self.__quality_indexes[key] = rebar_mass
            else:
                consumption = rebar_mass / self.__concrete_volume
                self.__quality_indexes[key] = consumption
                full_consumption += consumption
        self.__quality_indexes["Общий расход, кг/м3"] = full_consumption

    @reactive
    def quality_indexes(self):
        return self.__quality_indexes


class QualityTable:
    """
    Класс для формирования спецификации.
    """
    def __init__(self, table_type, construction):
        self.table_name = table_type.name
        self.indexes = table_type.quality_indexes
        self.quality_indexes = construction.quality_indexes

    def create_table(self):
        output = script.get_output()
        data = self.set_row_values()
        data.append(["_", "_", "_"])
        output.print_table(table_data=data,
                           title="Показатели качества",
                           columns=["Номер", "Название", "Значение"])

    def set_row_values(self):
        table = []
        if self.table_name == "Пилоны":
            table = [
                ["1", "Этажность здания, тип секции"],
                ["2", "Сечение пилонов, толщина х ширина, мм"],
                ["3", "Класс бетона"],
                ["4", "Объем бетона, м3"],
                ["5", "Коэффициент суммарной площади сечений пилонов от площади перекрытия, ΣAw/Ap х 100"],
                ["6.1", "Масса продольной арматуры, кг"],
                ["6.2", "Расход продольной арматуры, кг/м3"],
                ["7.1", "Масса поперечной арматуры, кг"],
                ["7.2", "Расход поперечной арматуры, кг/м3"],
                ["8", "Общий расход, кг/м3"],
            ]
        if self.table_name == "Стены":
            table = [
                ["1", "Этажность здания, тип секции"],
                ["2", "Толщина стен, мм"],
                ["3", "Класс бетона"],
                ["4", "Объем бетона, м3"],
                ["5.1", "Масса вертикальной арматуры, кг"],
                ["5.2", "Расход вертикальной арматуры, кг/м3"],
                ["6.1", "Масса горизонтальной арматуры, кг"],
                ["6.2", "Расход горизонтальной арматуры, кг/м3"],
                ["7.1", "Масса конструктивной арматуры, кг"],
                ["7.2", "Расход конструктивной арматуры, кг/м3"],
                ["8", "Общий расход, кг/м3"],
            ]

        output_rows = []
        for row in table:
            output_row = []
            output_row.append(row[0])
            output_row.append(row[1])
            output_row.append(self.quality_indexes[row[1]])
            output_rows.append(output_row)

        return output_rows


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
        if not concrete:
            alert("Не найден ЖБ для выбранных секций и разделов")
        else:
            rebar = self.__view_model.revit_repository.get_filtered_rebar_by_user(buildings, construction_sections)
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

        self.__selected_table_type = self.__table_types[0]
        self.__revit_repository = revit_repository
        self.__buildings = []
        self.__construction_sections = []

        self.__create_tables_command = CreateQualityTableCommand(self)

        self.__error_text = ""

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


def convert_value(app, parameter):
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


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    table_types = []

    walls_cat = Category.GetCategory(doc, BuiltInCategory.OST_Walls)
    columns_cat = Category.GetCategory(doc, BuiltInCategory.OST_StructuralColumns)
    foundation_cat = Category.GetCategory(doc, BuiltInCategory.OST_Doors)

    walls_table_type = TableType("Стены")
    walls_table_type.categories = [walls_cat]
    walls_table_type.type_key_word = "Стена"
    walls_table_type.quality_indexes = {"Масса вертикальной арматуры, кг": "Стены_Вертикальная",
                                        "Расход вертикальной арматуры, кг/м3": "Стены_Вертикальная",
                                        "Масса горизонтальной арматуры, кг": "Стены_Горизонтальная",
                                        "Расход горизонтальной арматуры, кг/м3": "Стены_Горизонтальная",
                                        "Масса конструктивной арматуры, кг": "Стены_Конструктивная",
                                        "Расход конструктивной арматуры, кг/м3": "Стены_Конструктивная"}

    columns_table_type = TableType("Пилоны")
    columns_table_type.categories = [walls_cat, columns_cat]
    columns_table_type.type_key_word = "Пилон"
    columns_table_type.quality_indexes = {"Масса продольной арматуры, кг": "Пилоны_Продольная",
                                          "Расход продольной арматуры, кг/м3": "Пилоны_Продольная",
                                          "Масса поперечной арматуры, кг": "Пилоны_Поперечная",
                                          "Расход поперечной арматуры, кг/м3": "Пилоны_Поперечная"
                                          }

    foundation_table_type = TableType("Фундаментная плита")
    foundation_table_type.categories = [foundation_cat]
    foundation_table_type.type_key_word = "ФПлита"
    foundation_table_type.quality_indexes = {"Масса продольной арматуры, кг": "Пилоны_Продольная",
                                             "Расход продольной арматуры, кг/м3": "Пилоны_Продольная",
                                             "Масса поперечной арматуры, кг": "Пилоны_Поперечная",
                                             "Расход поперечной арматуры, кг/м3": "Пилоны_Поперечная"
                                             }

    table_types.append(walls_table_type)
    table_types.append(columns_table_type)
    table_types.append(foundation_table_type)

    revit_repository = RevitRepository(doc)
    check = revit_repository.check_exist_main_parameters()
    if check:
        output = script.get_output()
        output.print_table(table_data=check,
                           title="Показатели качества",
                           columns=["Категории", "Тип ошибки", "Название параметра", "Id"])
        script.exit()

    revit_repository.filter_by_main_parameters()

    check = revit_repository.check_exist_rebar_parameters()
    if check:
        output = script.get_output()
        output.print_table(table_data=check,
                           title="Показатели качества",
                           columns=["Категории", "Тип ошибки", "Название параметра", "Id"])
        script.exit()

    check = revit_repository.check_parameters_values()
    if check:
        output = script.get_output()
        output.print_table(table_data=check,
                           title="Показатели качества",
                           columns=["Категории", "Тип ошибки", "Название параметра", "Id"])
        script.exit()

    main_window = MainWindow()
    main_window.DataContext = MainWindowViewModel(revit_repository, table_types)
    main_window.show_dialog()


script_execute()
