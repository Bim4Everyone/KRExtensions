# -*- coding: utf-8 -*-
import clr
import re

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

from System.Collections.Generic import List

from pyrevit import EXEC_PARAMS, revit
from pyrevit import script

from Autodesk.Revit.DB import *

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document

# ---------------- константы ----------------
PARAM_KM_TYPE = "ADSK_Тип элемента КМ"
PARAM_PROFILE_NAME = "ADSK_Наименование профиля"
PARAM_NAME = "ADSK_Наименование"
PARAM_IMAGE = "Изображение"
PARAM_IMAGE_SMP = "Изображение для СМП"

UNKNOWN_PROFILE_NAME = "Профиль не указан"

BUILT_IN_CATEGORIES = [
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_StructConnections,
]

# Номер изображения = количество разных наименований в профиле.
# Изображения подготовлены для 1–20 наименований (полоса растёт с шагом 8 мм),
# если наименований больше, используется самое широкое изображение.
TYPE_COUNT_TO_IMAGE_NAME = {
    1:  "8 мм.png",   2:  "16 мм.png",  3:  "24 мм.png",  4:  "32 мм.png",
    5:  "40 мм.png",  6:  "48 мм.png",  7:  "56 мм.png",  8:  "64 мм.png",
    9:  "72 мм.png",  10: "80 мм.png",  11: "88 мм.png",  12: "96 мм.png",
    13: "104 мм.png", 14: "112 мм.png", 15: "120 мм.png", 16: "128 мм.png",
    17: "136 мм.png", 18: "144 мм.png", 19: "152 мм.png", 20: "160 мм.png",
}
MIN_TYPE_COUNT = min(TYPE_COUNT_TO_IMAGE_NAME)
MAX_TYPE_COUNT = max(TYPE_COUNT_TO_IMAGE_NAME)

# Сколько имён изображений из модели показывать в отчёте, если нужных не нашлось
IMAGE_NAMES_PREVIEW_LIMIT = 60


class CmpProcessor:
    def __init__(self, doc):
        self.doc = doc
        self.__image_types_by_name = {}
        self.__errors = dict()

    @staticmethod
    def _normalize_name(name):
        if name is None:
            return None
        name = name.strip().lower().replace(u" ", u" ")
        return re.sub(r"\s+", " ", name)

    # ---- поиск изображений ----

    def _get_image_name(self, image_type):
        try:
            image_name = image_type.Name
            if image_name:
                return image_name
        except Exception:
            pass

        return image_type.GetParamValueOrDefault(BuiltInParameter.ALL_MODEL_TYPE_NAME)

    def _build_image_index(self):
        image_types = FilteredElementCollector(self.doc).OfClass(ImageType).ToElements()
        for image_type in image_types:
            image_name = self._get_image_name(image_type)
            if image_name:
                self.__image_types_by_name[self._normalize_name(image_name)] = image_type

    def _find_image_type_id(self, type_count):
        target_name = TYPE_COUNT_TO_IMAGE_NAME.get(type_count)
        if not target_name:
            return None

        normalized_name = self._normalize_name(target_name)

        if normalized_name in self.__image_types_by_name:
            return self.__image_types_by_name[normalized_name].Id

        name_without_extension = normalized_name.replace(".png", "")
        if name_without_extension in self.__image_types_by_name:
            return self.__image_types_by_name[name_without_extension].Id

        for image_name, image_type in self.__image_types_by_name.items():
            if image_name and name_without_extension in image_name:
                return image_type.Id

        return None

    def _get_image_ids_by_type_count(self):
        image_ids = {}
        missing_names = []
        for type_count in sorted(TYPE_COUNT_TO_IMAGE_NAME.keys()):
            image_id = self._find_image_type_id(type_count)
            if image_id is None:
                missing_names.append(TYPE_COUNT_TO_IMAGE_NAME[type_count])
            else:
                image_ids[type_count] = image_id

        if missing_names:
            available_names = sorted(self.__image_types_by_name.keys())[:IMAGE_NAMES_PREVIEW_LIMIT]
            output = script.get_output()
            output.print_md("**Ошибка: в модели не найдены изображения:**")
            output.print_md(", ".join(missing_names))
            output.print_md("")
            output.print_md("**Доступные имена изображений (первые {}):**".format(IMAGE_NAMES_PREVIEW_LIMIT))
            output.print_md(", ".join(available_names) if available_names else "(нет)")
            script.exit()

        return image_ids

    # ---- сбор и фильтрация элементов ----

    def _get_text_value(self, element, param_name):
        """Значение параметра экземпляра, а если его нет — параметра типа."""
        value = element.GetParamValueOrDefault(param_name)
        if value is None:
            element_type = self.doc.GetElement(element.GetTypeId())
            if element_type:
                value = element_type.GetParamValueOrDefault(param_name)

        if value is None:
            return ""
        return str(value).strip()

    def _collect_filtered_elements(self):
        category_filter = ElementMulticategoryFilter(List[BuiltInCategory](BUILT_IN_CATEGORIES))
        elements = FilteredElementCollector(self.doc) \
            .WherePasses(category_filter) \
            .WhereElementIsNotElementType() \
            .ToElements()

        return [element for element in elements if self._get_text_value(element, PARAM_KM_TYPE)]

    def _group_by_profile(self, elements):
        groups = {}
        for element in elements:
            profile_name = self._get_text_value(element, PARAM_PROFILE_NAME)
            if not profile_name:
                profile_name = UNKNOWN_PROFILE_NAME

            groups.setdefault(profile_name, [])
            groups[profile_name].append(element)

        return groups

    # ---- назначение параметров ----

    def _find_parameter(self, element, param_name):
        parameter = element.LookupParameter(param_name)
        if parameter:
            return parameter

        element_type = self.doc.GetElement(element.GetTypeId())
        if element_type:
            return element_type.LookupParameter(param_name)

        return None

    def _set_image_param(self, element, param_name, image_type_id):
        parameter = self._find_parameter(element, param_name)
        if not parameter:
            return None

        if parameter.IsReadOnly:
            return "параметр только для чтения"
        if parameter.StorageType != StorageType.ElementId:
            return "неверный тип данных параметра: {}".format(parameter.StorageType)

        current_id = parameter.AsElementId()
        if current_id and current_id.GetIdValue() == image_type_id.GetIdValue():
            return None

        try:
            parameter.Set(image_type_id)
        except Exception as ex:
            return str(ex)

        return None

    # ---- обработка ошибок ----

    def _add_error(self, error_type, element, parameter_name):
        key = error_type + "___" + parameter_name
        self.__errors.setdefault(key, [])
        self.__errors[key].append(str(element.Id))

    def _get_error_list(self):
        result = []
        for error_key, element_ids in self.__errors.items():
            parts = error_key.split("___")
            parts.append(", ".join(element_ids))
            result.append(parts)
        return result

    # ---- главный метод ----

    def execute(self):
        filtered_elements = self._collect_filtered_elements()

        if not filtered_elements:
            output = script.get_output()
            output.print_md("**Не найдено элементов с заполненным параметром '{}'**".format(PARAM_KM_TYPE))
            script.exit()

        groups = self._group_by_profile(filtered_elements)

        self._build_image_index()
        image_ids_by_type_count = self._get_image_ids_by_type_count()
        smp_image_id = image_ids_by_type_count[MIN_TYPE_COUNT]

        image_set_count = 0
        smp_image_set_count = 0

        with revit.Transaction("BIM: СМП — назначение изображений"):
            for profile_elements in groups.values():
                unique_names = set()
                for element in profile_elements:
                    name_value = self._get_text_value(element, PARAM_NAME)
                    if name_value:
                        unique_names.add(name_value)

                type_count = max(MIN_TYPE_COUNT, min(len(unique_names), MAX_TYPE_COUNT))
                group_image_id = image_ids_by_type_count[type_count]

                for element in profile_elements:
                    error = self._set_image_param(element, PARAM_IMAGE, group_image_id)
                    if error:
                        self._add_error("Назначение изображения: " + error, element, PARAM_IMAGE)
                    else:
                        image_set_count += 1

                    error = self._set_image_param(element, PARAM_IMAGE_SMP, smp_image_id)
                    if error:
                        self._add_error("Назначение изображения для СМП: " + error, element, PARAM_IMAGE_SMP)
                    else:
                        smp_image_set_count += 1

        output = script.get_output()
        output.print_md("**СМП: результаты**")
        output.print_md("- Элементов обработано: {}".format(len(filtered_elements)))
        output.print_md("- Профилей: {}".format(len(groups)))
        output.print_md("- Назначено / Проверено '{}': {}".format(PARAM_IMAGE, image_set_count))
        output.print_md("- Назначено / Проверено '{}': {}".format(PARAM_IMAGE_SMP, smp_image_set_count))

        errors_list = self._get_error_list()
        if errors_list:
            output.print_table(
                table_data=errors_list,
                title="Ошибки",
                columns=["Тип ошибки", "Параметр", "ID элементов"],
            )
        else:
            output.print_md("Ошибок нет.")


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    processor = CmpProcessor(doc)
    processor.execute()


script_execute()
