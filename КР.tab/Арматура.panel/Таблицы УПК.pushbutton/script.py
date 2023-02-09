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


class RevitRepository:
    """
    Класс для получения всего бетона и арматуры из проекта.
    Фильтрует все элементы.
    """
    def __init__(self, doc, table_type):
        self.doc = doc
        self.categories = table_type.categories
        self.type_key_word = table_type.type_key_word
        self.quality_indexes = table_type.quality_indexes

        self.__concrete = self.__get_concrete()
        self.__buildings = self.__get_buildings()
        self.__construction_sections = self.__get_construction_sections()
        self.__rebar = self.__get_rebar()

    def update_elements_info(self):
        pass

    def get_filtered_concrete(self, buildings, constr_sections):
        buildings = [x.text_value for x in buildings if x.is_checked]
        constr_sections = [x.text_value for x in constr_sections if x.is_checked]
        filtered_elements = []
        for element in self.concrete:
            param1 = element.LookupParameter("ФОП_Секция СМР")
            param2 = element.LookupParameter("обр_ФОП_Раздел проекта")

            if self.get_param_value(param1) in buildings:
                if self.get_param_value(param2) in constr_sections:
                    filtered_elements.append(element)

        return filtered_elements

    def get_filtered_rebar(self, buildings, constr_sections):
        buildings = [x.text_value for x in buildings if x.is_checked]
        constr_sections = [x.text_value for x in constr_sections if x.is_checked]
        filtered_elements = []
        for element in self.rebar:
            param1 = element.LookupParameter("ФОП_Секция СМР")
            param2 = element.LookupParameter("обр_ФОП_Раздел проекта")

            if self.get_param_value(param1) in buildings:
                if self.get_param_value(param2) in constr_sections:
                    filtered_elements.append(element)

        return filtered_elements

    def __create_param_set(self, elements, param_name):
        set_of_values = set()
        for element in elements:
            param = element.LookupParameter(param_name)
            set_of_values.add(self.get_param_value(param))
        return sorted(set_of_values)

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

    def __filter_by_param(self, elements, param_name, value):
        filtered_list = []
        for element in elements:
            param = element.LookupParameter(param_name)
            if self.get_param_value(param) != value:
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
        elements = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_Rebar)
        elements.WhereElementIsNotElementType().ToElements()
        elements = self.__filter_by_param(elements, "обр_ФОП_Фильтрация 1", "Исключить из показателей качества")
        elements = self.__filter_by_param(elements, "обр_ФОП_Фильтрация 2", "Исключить из показателей качества")
        elements_new = []
        for value in self.quality_indexes.values():
            elements_new += self.__filter_by_param(elements, "обр_ФОП_Группа КР", value)

        # elements = self.__filter_by_param(elements, "обр_ФОП_Форма номер", "1000")
        return elements_new

    @staticmethod
    def get_param_value(param):
        if param.HasValue:
            return param.AsString()
        else:
            return ""

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
        if text_value:
            self.__number = text_value
        else:
            self.__number = "<Параметр не заполнен>"
        self.__is_checked = False

    @reactive
    def number(self):
        return self.__number

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

    def __init__(self, concrete_elements, rebar_elements):
        self.__concrete = concrete_elements
        self.__rebar = rebar_elements

        self.rebar_by_function = dict()
        self.rebar_mass_by_function = dict()

        self.group_rebar_by_function()
        self.calculate_quality_indexes()
        self.calculate_concrete_volume()

    def calculate_concrete_volume(self):
        volume = 0
        for element in self.__concrete:
            volume += element.LookupParameter("Объем").AsDouble()

    def calculate_rebar_mass(self, elements):
        return 100

    def group_rebar_by_function(self):
        for element in self.__rebar:
            rebar_function = element.LookupParameter("обр_ФОП_Группа КР").AsString()
            self.rebar_by_function.setdefault(rebar_function, [])
            self.rebar_by_function[rebar_function].append(element)

    def calculate_quality_indexes(self):
        for key in self.rebar_by_function.keys():
            elements = self.rebar_by_function[key]
            rebar_mass = self.calculate_rebar_mass(elements)
            self.rebar_mass_by_function[key] = rebar_mass

    def get_quality_index(self, name):
        return self.rebar_mass_by_function[name]


class QualityTable:
    """
    Класс для формирования спецификации.
    """
    def __init__(self, table_type, construction):
        self.indexes = table_type.quality_indexes
        self.construction = construction

    def create_table(self):
        output = script.get_output()
        data = self.set_row_values()
        output.print_table(table_data=data,
                           title="Показатели качества",
                           columns=["Название", "Значение"])

    def set_row_values(self):
        rows = []
        for index_name in self.indexes.keys():
            # alert(index_name)
            row = []
            row.append(index_name)
            index = self.indexes[index_name]
            value = self.construction.get_quality_index(index)
            # alert(str(value))
            row.append(str(value))
            rows.append(row)
        return rows


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
        return True

    def Execute(self, parameter):
        self.__view_model.quality_table.create_table()
        # alert(str(len(self.__view_model.quality_table)))


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
        self.__quality_table = []

        self.__create_tables_command = CreateQualityTableCommand(self)

    @reactive
    def table_types(self):
        return self.__table_types

    @reactive
    def selected_table_type(self):
        return self.__selected_table_type

    @selected_table_type.setter
    def selected_table_type(self, value):
        self.__revit_repository = RevitRepository(doc, value)
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
    def quality_table(self):
        concrete = self.__revit_repository.get_filtered_concrete(self.buildings, self.construction_sections)
        rebar = self.__revit_repository.get_filtered_rebar(self.buildings, self.construction_sections)
        construction = Construction(concrete, rebar)

        return QualityTable(self.selected_table_type, construction)

    # @quality_table.setter
    # def quality_table(self, value):
    #     self.__quality_table = value

    @property
    def create_tables_command(self):
        return self.__create_tables_command


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    table_types = []

    walls_cat = Category.GetCategory(doc, BuiltInCategory.OST_Walls)
    columns_cat = Category.GetCategory(doc, BuiltInCategory.OST_StructuralColumns)

    walls_table_type = TableType("Стены")
    walls_table_type.categories = [walls_cat]
    walls_table_type.type_key_word = "Стена"
    walls_table_type.quality_indexes = {"Масса вертикальной арматуры, кг": "Продольная арматура",
                                  "Масса горизонтальной арматуры, кг": "Горизонтальная арматура",
                                  "Масса конструктивной арматуры, кг": "Конструктивная арматура"}

    columns_table_type = TableType("Пилоны")
    columns_table_type.categories = [walls_cat, columns_cat]
    columns_table_type.type_key_word = "Пилон"
    columns_table_type.quality_indexes = {"Масса продольной арматуры, кг": "Пилоны_Продольная",
                                "Масса поперечной арматуры, кг": "Пилоны_Поперечная"}

    table_types.append(walls_table_type)
    table_types.append(columns_table_type)

    main_window = MainWindow()
    main_window.DataContext = MainWindowViewModel(table_types)
    main_window.show_dialog()


script_execute()
