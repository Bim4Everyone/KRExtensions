# -*- coding: utf-8 -*-
import clr

from System.Collections.Generic import *

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

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


class RevitRepository:
    """
    Класс для получения всего бетона и арматуры из проекта.
    Фильтрует все элементы.
    """
    def __init__(self, doc, table_type):
        self.doc = doc
        self.categories = table_type.categories
        self.type_key_word = table_type.type_key_word

        self.__concrete = self.__get_concrete()
        self.__buildings = self.__get_buildings()
        self.__construction_sections = self.__get_construction_sections()
        # self.__rebar = self.__get_rebar()

    def __create_param_set(self, elements, param_name):
        set_of_values = set()
        for element in elements:
            param = element.LookupParameter(param_name)
            if param.HasValue:
                set_of_values.add(param.AsString())
            else:
                set_of_values.add("<Параметр не заполнен>")
        return sorted(set_of_values)

    def __get_buildings(self):
        buildings = self.__create_param_set(self.concrete, "ФОП_Секция СМР")
        result_buildings = []
        for building in buildings:
            result_buildings.append(Building(building))
        return result_buildings

    def __get_construction_sections(self):
        sections = self.__create_param_set(self.concrete, "обр_ФОП_Раздел проекта")
        result_sections = []
        for section in sections:
            result_sections.append(ConstructionSection(section))
        return result_sections

    def __create_param_filter(self, name):
        fvp = ParameterValueProvider(ElementId(BuiltInParameter.VIEW_NAME))
        rule = FilterStringEquals()
        value = name
        case_sens = False
        filter_rule = FilterStringRule(fvp, rule, value, case_sens)
        name_filter = ElementParameterFilter(filter_rule)
        return name_filter

    def __filter_by_param(self, elements, param_name, value):
        filtered_list = []
        for element in elements:
            param_value = str(element.LookupParameter(param_name).AsString())
            if param_value != value:
                filtered_list.append(element)
        return filtered_list

    def __filter_by_type(self, elements):
        filtered_list = []
        for element in elements:
            # Проверить получение типоразмера по Name
            if self.type_key_word in element.Name:
                filtered_list.append(element)

        return filtered_list

    def __get_concrete(self):
        cat_filters = [ElementCategoryFilter(x.Id) for x in self.categories]
        cat_filters_typed = List[ElementFilter](cat_filters)
        logical_or_filter = LogicalOrFilter(cat_filters_typed)
        elements = FilteredElementCollector(self.doc).WherePasses(logical_or_filter).WhereElementIsNotElementType().ToElements()

        elements = self.__filter_by_param(elements, "обр_ФОП_Фильтрация 1", "Исключить из показателей качества")
        elements = self.__filter_by_param(elements, "обр_ФОП_Фильтрация 2", "Исключить из показателей качества")
        elements = self.__filter_by_type(elements)

        return elements

    def __get_rebar(self):
        elements = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_Rebar).ToELements()
        elements = self.__filter_by_param(elements, "обр_ФОП_Фильтрация 1", "Исключить из показателей качества")
        elements = self.__filter_by_param(elements, "обр_ФОП_Фильтрация 2", "Исключить из показателей качества")
        # elements = self.__filter_by_param(elements, "обр_ФОП_Форма номер", "1000")
        return elements

    @reactive
    def concrete(self):
        return self.__concrete

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


class Construction:
    """
    Класс для расчетов всех показателей.
    """

    def __init__(self):
        pass

    def get_concrete_volume(self):
        pass


class Building:
    def __init__(self, number):
        self.__number = number
        self.__is_checked = False

    @reactive
    def number(self):
        return self.__number

    @number.setter
    def number(self, value):
        self.__number = value

    @reactive
    def is_checked(self):
        return self.__is_checked

    @is_checked.setter
    def is_checked(self, value):
        self.__is_checked = value


class ConstructionSection:
    def __init__(self, section_name):
        self.__section_name = section_name
        self.__is_checked = False

    @reactive
    def section_name(self):
        return self.__section_name

    @section_name.setter
    def section_name(self, value):
        self.__section_name = value

    @reactive
    def is_checked(self):
        return self.__is_checked

    @is_checked.setter
    def is_checked(self, value):
        self.__is_checked = value


class QualityTable:
    """
    Класс для формирования спецификации.
    """

    def __init__(self):
        pass

    def create_table(self):
        pass


class MainWindow(WPFWindow):
    def __init__(self):
        self._context = None
        self.xaml_source = op.join(op.dirname(__file__), 'MainWindow.xaml')
        super(MainWindow, self).__init__(self.xaml_source)


class MainWindowViewModel(Reactive):
    def __init__(self, table_types):
        Reactive.__init__(self)

        self.__table_types = []
        self.__table_types = table_types

        self.__selected_table_type = self.__table_types[0]
        self.__revit_repository = None
        self.__buildings = []
        self.__construction_sections = []
        self.__table_concrete = []

    @reactive
    def table_types(self):
        return self.__table_types

    @reactive
    def selected_table_type(self):
        return self.__selected_table_type

    @selected_table_type.setter
    def selected_table_type(self, value):
        self.__revit_repository = RevitRepository(doc, value)
        self.table_concrete = self.__revit_repository.concrete
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
    def table_concrete(self):
        return self.__table_concrete

    @table_concrete.setter
    def table_concrete(self, value):
        self.__table_concrete = value


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    table_types = []

    walls_cat = Category.GetCategory(doc, BuiltInCategory.OST_Walls)
    columns_cat = Category.GetCategory(doc, BuiltInCategory.OST_StructuralColumns)

    walls_table_type = TableType("Стены")
    walls_table_type.categories = [walls_cat]
    walls_table_type.type_key_word = "Стена"

    columns_table_type = TableType("Пилоны")
    columns_table_type.categories = [walls_cat, columns_cat]
    columns_table_type.type_key_word = "Пилон"

    table_types.append(walls_table_type)
    table_types.append(columns_table_type)

    main_window = MainWindow()
    main_window.DataContext = MainWindowViewModel(table_types)
    main_window.show_dialog()


script_execute()