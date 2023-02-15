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
    def __init__(self, doc):
        self.doc = doc

        self.categories = []
        self.type_key_word = []
        self.quality_indexes = []
        self.__rebar = []
        self.__concrete = []
        self.__buildings = []
        self.__construction_sections = []

    def set_table_type(self, table_type):
        self.categories = table_type.categories
        self.type_key_word = table_type.type_key_word
        self.quality_indexes = table_type.quality_indexes

        # self.__rebar = self.__get_filtered_rebar_by_rules()

        self.__concrete = self.__get_concrete_table_type()
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

        all_rebar = self.__get_all_rebar()
        all_concrete = self.__get_all_concrete()

        for element in all_rebar:
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

        for element in all_concrete:
            for parameter_name in concrete_common_parameters:
                if not element.IsExistsParam(parameter_name):
                    key = "Железобетон___Отсутствует параметр у экземпляра___" + parameter_name
                    errors_dict.setdefault(key, [])
                    errors_dict[key].append(str(element.Id))

        if errors_dict:
            missing_parameters = self.__create_error_list(errors_dict)
            return missing_parameters
        else:
            self.__rebar = self.__get_filtered_rebar_by_rules(all_rebar)

    def check_exist_rebar_parameters(self):
        errors_dict = dict()
        rebar_inst_parameters = ["обр_ФОП_Длина",
                                 "обр_ФОП_Группа КР",
                                 "обр_ФОП_Количество типовых на этаже",
                                 "обр_ФОП_Количество типовых этажей"]

        rebar_type_parameters = ["мод_ФОП_IFC семейство"]

        rebar_inst_type_parameters = ["мод_ФОП_Диаметр",
                                      "обр_ФОП_Форма_номер"]

        rebar_ifc_parameters = ["обр_ФОП_Количество",
                                "Количество"]

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

    def __get_all_concrete(self):
        all_categories = []
        all_categories.append(Category.GetCategory(doc, BuiltInCategory.OST_Walls))
        all_categories.append(Category.GetCategory(doc, BuiltInCategory.OST_StructuralColumns))
        all_categories.append(Category.GetCategory(doc, BuiltInCategory.OST_StructuralFoundation))
        all_categories.append(Category.GetCategory(doc, BuiltInCategory.OST_Floors))
        all_categories.append(Category.GetCategory(doc, BuiltInCategory.OST_StructuralFraming))
        elements = self.__collect_elements_by_categories(all_categories)
        return elements

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

    def __get_filtered_rebar_by_rules(self, elements):
        elements = self.__filter_by_param(elements, "обр_ФОП_Фильтрация 1", "Исключить из показателей качества")
        elements = self.__filter_by_param(elements, "обр_ФОП_Фильтрация 2", "Исключить из показателей качества")
        elements = self.__filter_by_param(elements, "обр_ФОП_Форма_номер", 1000)
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
        cat_filters = [ElementCategoryFilter(x.Id) for x in categories]
        cat_filters_typed = List[ElementFilter](cat_filters)
        logical_or_filter = LogicalOrFilter(cat_filters_typed)
        elements = FilteredElementCollector(self.doc).WherePasses(logical_or_filter)
        elements.WhereElementIsNotElementType().ToElements()
        return elements

    def __get_concrete_table_type(self):
        elements = self.__collect_elements_by_categories(self.categories)

        elements = self.__filter_by_param(elements, "обр_ФОП_Фильтрация 1", "Исключить из показателей качества")
        elements = self.__filter_by_param(elements, "обр_ФОП_Фильтрация 2", "Исключить из показателей качества")
        elements = self.__filter_by_type(elements)

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
    def __init__(self, concrete_elements, rebar_elements):
        self.concrete = concrete_elements
        self.rebar = rebar_elements

        self.rebar_by_function = dict()
        self.rebar_mass_by_function = dict()
        self.concrete_volume = 0

        self.group_rebar_by_function()
        self.calculate_quality_indexes()
        self.calculate_concrete_volume()

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

    def calculate_concrete_volume(self):
        for element in self.concrete:
            self.concrete_volume += element.GetParamValue("Объем")

    def calculate_rebar_mass(self, elements):
        rebar_mass = 0
        for element in elements:
            element_type = doc.GetElement(element.GetTypeId())
            is_ifc_element = element_type.GetParamValue("мод_ФОП_IFC семейство")
            try:
                diameter = element.GetParamValue("мод_ФОП_Диаметр")
                length = element.GetParamValue("обр_ФОП_Длина")
                # mass_per_metr = self.__diameter_dict[diameter]
                mass_per_metr = 1
                if is_ifc_element:
                    amount = element.GetParamValue("обр_ФОП_Количество")
                else:
                    amount = element.GetParamValue("Количество")
                amount_on_level = element.GetParamValue("обр_ФОП_Количество типовых на этаже")
                levels_amount = element.GetParamValue("обр_ФОП_Количество типовых этажей")

                element_mass = mass_per_metr * length * amount * amount_on_level * levels_amount
                rebar_mass += element_mass
            except:
                pass

        return rebar_mass

    def group_rebar_by_function(self):
        for element in self.rebar:
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
        # alert(str(len(self.construction.concrete)))
        # alert(str(len(self.construction.rebar)))

        output = script.get_output()
        data = self.set_row_values()
        output.print_table(table_data=data,
                           title="Показатели качества",
                           columns=["Название", "Значение"])

    def set_row_values(self):
        rows = []
        # 1 Этажность здания, тип секции
        first_row = []
        first_row.append("1 Этажность здания, тип секции")
        first_row.append("0")
        rows.append(first_row)

        # 2 Толщина стен, мм / Сечение пилонов, толщина х ширина, мм
        second_row = []
        second_row.append("2 Толщина")
        second_row.append("0")
        rows.append(second_row)

        # 3 Класс бетона
        third_row = []
        third_row.append("3 Класс бетона")
        third_row.append("B0")
        rows.append(third_row)

        # 4 Объем бетона, м3
        fourth_row = []
        fourth_row.append("4 Объем бетона, м3")
        fourth_row.append(str(self.construction.concrete_volume))
        rows.append(fourth_row)

        # 5 Показатели качества
        for index_name in self.indexes.keys():
            row = []
            row.append(index_name)
            index = self.indexes[index_name]
            value = self.construction.get_quality_index(index)
            if "Масса" in index_name:
                row.append(str(value))
            else:
                concrete_volume = self.construction.concrete_volume
                row.append(str(value/concrete_volume))
            rows.append(row)

        # Общий расход, кг / м3
        last_row = []
        last_row.append("10 Общий расход, кг / м3")
        last_row.append(str(self.construction.concrete_volume))
        rows.append(last_row)
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
        # if not self.__view_model.self.filtered_concrete:
        #     self.__view_model.error_text = "Заполните все поля."
        #     return False
        #
        # self.__view_model.error_text = None
        return True

    def Execute(self, parameter):
        if self.__view_model.quality_table:
            self.__view_model.quality_table.create_table()


class MainWindow(WPFWindow):
    def __init__(self):
        self._context = None
        self.xaml_source = op.join(op.dirname(__file__), 'MainWindow.xaml')
        super(MainWindow, self).__init__(self.xaml_source)


class MainWindowViewModel(Reactive):
    def __init__(self, revit_repository, table_types):
        Reactive.__init__(self)

        self.__table_types = []
        self.__table_types = table_types

        self.__selected_table_type = self.__table_types[0]
        self.__revit_repository = revit_repository
        self.__buildings = []
        self.__construction_sections = []
        self.__quality_table = []

        self.__create_tables_command = CreateQualityTableCommand(self)

        self.__error_text = ""

    @reactive
    def table_types(self):
        return self.__table_types

    @reactive
    def selected_table_type(self):
        return self.__selected_table_type

    @selected_table_type.setter
    def selected_table_type(self, value):
        self.__revit_repository.set_table_type(value)
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
        concrete = self.__revit_repository.get_filtered_concrete_by_user(self.buildings, self.construction_sections)
        rebar = self.__revit_repository.get_filtered_rebar_by_user(self.buildings, self.construction_sections)
        construction = Construction(concrete, rebar)

        return QualityTable(self.selected_table_type, construction)

    @property
    def create_tables_command(self):
        return self.__create_tables_command


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    table_types = []

    walls_cat = Category.GetCategory(doc, BuiltInCategory.OST_Walls)
    columns_cat = Category.GetCategory(doc, BuiltInCategory.OST_StructuralColumns)

    empty_table_type = TableType("<Тип таблицы>")
    empty_table_type.categories = []
    empty_table_type.type_key_word = ""
    empty_table_type.quality_indexes = dict()

    walls_table_type = TableType("Стены")
    walls_table_type.categories = [walls_cat]
    walls_table_type.type_key_word = "Стена"
    walls_table_type.quality_indexes = {"Масса вертикальной арматуры, кг": "Продольная арматура",
                                        "Расход вертикальной арматуры, кг/м3": "Продольная арматура",
                                        "Масса горизонтальной арматуры, кг": "Горизонтальная арматура",
                                        "Расход горизонтальной арматуры, кг/м3": "Горизонтальная арматура",
                                        "Масса конструктивной арматуры, кг": "Конструктивная арматураа",
                                        "Расход конструктивной арматуры, кг/м3": "Конструктивная арматура"}

    columns_table_type = TableType("Пилоны")
    columns_table_type.categories = [walls_cat, columns_cat]
    columns_table_type.type_key_word = "Пилон"
    columns_table_type.quality_indexes = {"6.1_Масса продольной арматуры, кг": "Пилоны_Продольная",
                                          "6.2_Расход продольной арматуры, кг/м3": "Пилоны_Продольная",
                                          "7.1_Масса поперечной арматуры, кг": "Пилоны_Поперечная",
                                          "7.2_Расход поперечной арматуры, кг/м3": "Пилоны_Поперечная"
                                          }

    table_types.append(empty_table_type)
    table_types.append(walls_table_type)
    table_types.append(columns_table_type)

    revit_repository = RevitRepository(doc)
    check = revit_repository.check_exist_main_parameters()
    if check:
        output = script.get_output()
        output.print_table(table_data=check,
                           title="Показатели качества",
                           columns=["Категории", "Тип ошибки", "Название параметра", "Id"])
        script.exit()

    check = revit_repository.check_exist_rebar_parameters()
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
