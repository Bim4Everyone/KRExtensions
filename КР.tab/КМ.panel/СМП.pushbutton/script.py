# -*- coding: utf-8 -*-
import os
import sys
import clr
import re

sys.path.append(os.path.join(os.getenv("APPDATA"), r"pyRevit\Extensions\BIM4Everyone.lib"))
import dosymep_libs

dosymep_libs.load_assemblies()

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

from pyrevit import EXEC_PARAMS, revit
from pyrevit import script

from Autodesk.Revit.DB import *

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document

# ---------------- constants ----------------
PARAM_KM_TYPE = "ADSK_Тип элемента КМ"
PARAM_IMAGE = "Изображение"
PARAM_IMAGE_SMP = "Изображение для СМП"

BUILT_IN_CATS = [
    BuiltInCategory.OST_StructuralColumns,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_StructConnections,
]

LEVEL_TO_IMAGE_NAME = {
    1:  "8 мм.png",   2:  "16 мм.png",  3:  "24 мм.png",  4:  "32 мм.png",
    5:  "40 мм.png",  6:  "48 мм.png",  7:  "56 мм.png",  8:  "64 мм.png",
    9:  "72 мм.png",  10: "80 мм.png",  11: "88 мм.png",  12: "96 мм.png",
    13: "104 мм.png", 14: "112 мм.png", 15: "120 мм.png", 16: "128 мм.png",
    17: "136 мм.png", 18: "144 мм.png", 19: "152 мм.png", 20: "160 мм.png",
}


class CmpProcessor:
    def __init__(self, doc):
        self.doc = doc
        self.__img_by_norm = {}
        self.__errors = dict()

    @staticmethod
    def _norm(s):
        if s is None:
            return None
        s = s.strip().lower().replace(u"\u00A0", u" ")
        return re.sub(r"\s+", " ", s)

    # ---- поиск изображений ----

    def _get_image_name(self, img_type):
        try:
            n = img_type.Name
            if n:
                return n
        except Exception:
            pass

        try:
            p = img_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            if p:
                n = p.AsString() or p.AsValueString()
                if n:
                    return n
        except Exception:
            pass

        return None

    def _build_image_index(self):
        image_types = list(FilteredElementCollector(self.doc).OfClass(ImageType).ToElements())
        for it in image_types:
            nm = self._get_image_name(it)
            if nm:
                self.__img_by_norm[self._norm(nm)] = it

    def _find_image_type_id(self, level):
        target = LEVEL_TO_IMAGE_NAME.get(level)
        if not target:
            return None

        tnorm = self._norm(target)

        if tnorm in self.__img_by_norm:
            return self.__img_by_norm[tnorm].Id

        if tnorm.endswith(".png"):
            t2 = tnorm[:-4]
            if t2 in self.__img_by_norm:
                return self.__img_by_norm[t2].Id

        key = tnorm.replace(".png", "")
        for k, it in self.__img_by_norm.items():
            if k and key in k:
                return it.Id

        return None

    def _validate_images(self, plugin_logger):
        missing = []
        for lvl in sorted(LEVEL_TO_IMAGE_NAME.keys()):
            img_id = self._find_image_type_id(lvl)
            if img_id is None:
                missing.append(LEVEL_TO_IMAGE_NAME[lvl])

        if missing:
            preview = sorted(self.__img_by_norm.keys())[:60]
            output = script.get_output()
            output.print_md("**Ошибка: не найдены изображения для уровней:**")
            output.print_md(", ".join(missing))
            output.print_md("")
            output.print_md("**Доступные имена изображений (первые 60):**")
            output.print_md(", ".join(preview) if preview else "(нет)")
            plugin_logger.error("Не найдено изображений для уровней: {}".format(missing))
            script.exit()

    # ---- сбор и фильтрация элементов ----

    def _lookup_param(self, elem, param_name):
        p = elem.LookupParameter(param_name)
        if p:
            return p

        try:
            type_id = elem.GetTypeId()
            if type_id and type_id != ElementId.InvalidElementId:
                elem_type = self.doc.GetElement(type_id)
                if elem_type:
                    return elem_type.LookupParameter(param_name)
        except Exception:
            pass

        return None

    def _get_param_value(self, elem, param_name):
        p = self._lookup_param(elem, param_name)
        if not p:
            return None

        st = p.StorageType
        if st == StorageType.String:
            return p.AsString() or p.AsValueString()
        if st == StorageType.Integer:
            return p.AsInteger()
        if st == StorageType.Double:
            return p.AsDouble()
        if st == StorageType.ElementId:
            eid = p.AsElementId()
            return eid.IntegerValue if eid else None
        return p.AsValueString()

    def _collect_filtered_elements(self):
        seen = set()
        all_elems = []

        for bic in BUILT_IN_CATS:
            col = FilteredElementCollector(self.doc).OfCategory(bic).WhereElementIsNotElementType()
            for e in col:
                iid = e.Id.IntegerValue
                if iid in seen:
                    continue
                seen.add(iid)
                all_elems.append(e)

        filtered = []
        for e in all_elems:
            km = self._get_param_value(e, PARAM_KM_TYPE)
            if km is None or str(km).strip() == "":
                continue
            filtered.append(e)

        return filtered

    def _group_by_family(self, elements):
        groups = {}
        for e in elements:
            type_id = e.GetTypeId()
            if type_id == ElementId.InvalidElementId:
                continue

            elem_type = self.doc.GetElement(type_id)
            if not elem_type:
                continue

            if hasattr(elem_type, "FamilyName") and elem_type.FamilyName:
                fam_name = elem_type.FamilyName
            else:
                fam_name = elem_type.Name

            if not fam_name:
                fam_name = "Unknown Family"

            groups.setdefault(fam_name, [])
            groups[fam_name].append(e)

        return groups

    # ---- назначение параметров ----

    def _set_image_param(self, elem, param_name, image_type_id):
        p = self._lookup_param(elem, param_name)
        if not p:
            return None

        if p.IsReadOnly:
            return "read-only"
        if p.StorageType != StorageType.ElementId:
            return "wrong StorageType: {}".format(p.StorageType)

        current_id = p.AsElementId()
        if current_id and current_id.IntegerValue == image_type_id.IntegerValue:
            return None

        try:
            p.Set(image_type_id)
        except Exception as e:
            return str(e)

        return None

    # ---- обработка ошибок ----

    def _add_error(self, error_type, element, parameter_name):
        key = error_type + "___" + parameter_name
        self.__errors.setdefault(key, [])
        self.__errors[key].append(str(element.Id))

    def _get_error_list(self):
        result = []
        for error_key, ids in self.__errors.items():
            parts = error_key.split("___")
            parts.append(", ".join(ids))
            result.append(parts)
        return result

    # ---- главный метод ----

    def execute(self, plugin_logger):
        plugin_logger.info("CMP: Сбор элементов...")

        filtered = self._collect_filtered_elements()
        plugin_logger.info("CMP: Найдено {} отфильтрованных элементов".format(len(filtered)))

        if not filtered:
            output = script.get_output()
            output.print_md("**Не найдено элементов с заполненным параметром '{}'**".format(PARAM_KM_TYPE))
            plugin_logger.warning("CMP: нет элементов для обработки")
            script.exit()

        groups = self._group_by_family(filtered)
        plugin_logger.info("CMP: {} семейств".format(len(groups)))

        self._build_image_index()
        self._validate_images(plugin_logger)

        level_to_img_id = {}
        for lvl in sorted(LEVEL_TO_IMAGE_NAME.keys()):
            img_id = self._find_image_type_id(lvl)
            if img_id is not None:
                level_to_img_id[lvl] = img_id

        img_id_smp = level_to_img_id[1]

        set_img = 0
        set_smp = 0

        with revit.Transaction("BIM: CMP — назначение изображений"):
            for fam_name, elems in groups.items():
                unique_type_ids = set()
                for e in elems:
                    unique_type_ids.add(e.GetTypeId().IntegerValue)

                type_count = len(unique_type_ids)
                lvl = max(1, min(type_count, 20))
                img_id_for_group = level_to_img_id[lvl]

                for e in elems:
                    err = self._set_image_param(e, PARAM_IMAGE, img_id_for_group)
                    if err:
                        self._add_error("Назначение изображения", e, PARAM_IMAGE)
                    else:
                        set_img += 1

                    err2 = self._set_image_param(e, PARAM_IMAGE_SMP, img_id_smp)
                    if err2:
                        self._add_error("Назначение изображения для СМП", e, PARAM_IMAGE_SMP)
                    else:
                        set_smp += 1

        output = script.get_output()
        output.print_md("**CMP: результаты**")
        output.print_md("- Элементов обработано: {}".format(len(filtered)))
        output.print_md("- Семейств: {}".format(len(groups)))
        output.print_md("- Назначено / Проверено '{}': {}".format(PARAM_IMAGE, set_img))
        output.print_md("- Назначено / Проверено '{}': {}".format(PARAM_IMAGE_SMP, set_smp))

        errors_list = self._get_error_list()
        if errors_list:
            output.print_table(
                table_data=errors_list,
                title="Ошибки",
                columns=["Тип ошибки", "Параметр", "ID элементов"],
            )
        else:
            output.print_md("Ошибок нет.")

        plugin_logger.info("CMP: Завершено. Изображение — {}, СМП — {}".format(set_img, set_smp))


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    processor = CmpProcessor(doc)
    processor.execute(plugin_logger)


script_execute()