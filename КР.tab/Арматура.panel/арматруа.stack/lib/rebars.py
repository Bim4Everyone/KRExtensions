# -*- coding: utf-8 -*-

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from System import *
from System.Collections.Generic import *

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *

from pyrevit import forms
from pyrevit.script import output


def set_solid_in_view(application, document, is_solid):
    if not isinstance(document.ActiveView, View3D):
        forms.alert("Активный вид должен быть 3D", exitscript=True)


    category_list = List[Type]()
    category_list.Add(Rebar)
    category_list.Add(RebarInSystem)

    category_filter = ElementMulticlassFilter(category_list)
    elements = FilteredElementCollector(document, document.ActiveView.Id)\
        .WherePasses(category_filter)\
        .WhereElementIsNotElementType()\
        .ToElements()

    report_rows = set()
    with Transaction(document) as transaction:
        transaction.Start("Включение 3D арматуры")

        for element in elements:
            if isinstance(element, (Rebar, RebarInSystem)):

                edited_by = element.GetParamValueOrDefault(BuiltInParameter.EDITED_BY)
                if edited_by and edited_by != application.Username:
                    report_rows.add(edited_by)
                    continue

                element.SetSolidInView(document.ActiveView, is_solid)

        transaction.Commit()

        if report_rows:
            output1 = output.get_output()
            output1.set_title("Включение 3D арматуры")

            print "Некоторые элементы не были обработаны, так как были заняты пользователями:"
            print "\r\n".join(report_rows)