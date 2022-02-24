# -*- coding: utf-8 -*-

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from math import sqrt

from System import *
from System.Collections.Generic import *

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *

from pyrevit import forms


application = __revit__.Application
document = __revit__.ActiveUIDocument.Document
activeView = document.ActiveView

options = Options()
options.View = activeView

if not isinstance(document.ActiveView, View3D):
    forms.alert("Активный вид должен быть 3D", exitscript=True)


def get_subtract(bounding_box):
    return bounding_box.Max - bounding_box.Min


def get_hypotenuse(bounding_box):
    bb_min = bounding_box.Min
    bb_max = bounding_box.Max

    return sqrt((bb_max.X - bb_min.X) ** 2 + (bb_max.Y - bb_min.Y) ** 2)


def ZIsLonger(subtract):
    return subtract.Z > subtract.X and subtract.Z > subtract.Y

def XIsLonger(subtract):
    return subtract.X > subtract.Z and subtract.X > subtract.Y

def YIsLonger(subtract):
    return subtract.Y > subtract.Z and subtract.Y > subtract.X


class RebarElement:
    def __init__(self, element):
        self.Element = element
        self.Host = document.GetElement(element.GetHostId())

    @property
    def BoundingBox(self):
        solid = next((g for g in self.Element.get_Geometry(options) if isinstance(g, Solid)), None)
        if solid:
            return solid.GetBoundingBox()

        return self.Element.get_BoundingBox(activeView)

    @property
    def HostCategory(self):
        return self.Host.Category

    @property
    def HostBoundingBox(self):
        return self.Host.get_BoundingBox(activeView)

    @property
    def ZIsLonger(self):
        subtract = get_subtract(self.BoundingBox)
        return subtract.Z > subtract.X and subtract.Z > subtract.Y

    @property
    def XIsLonger(self):
        subtract = get_subtract(self.BoundingBox)
        return subtract.X > subtract.Z and subtract.X > subtract.Y

    @property
    def YIsLonger(self):
        subtract = get_subtract(self.BoundingBox)
        return subtract.Y > subtract.Z and subtract.Y > subtract.X

    @property
    def IsAllowProcess(self):
        if self.HostCategory.Id == ElementId(BuiltInCategory.OST_Walls):
            self_subtract = get_subtract(self.BoundingBox)
            wall_subtract = get_subtract(self.HostBoundingBox)

            if ZIsLonger(self_subtract):
                return 2 * self_subtract.Z > wall_subtract.Z

            length = get_hypotenuse(self.BoundingBox)
            wall_length = get_hypotenuse(self.HostBoundingBox)

            return 2 * length > wall_length

        return True


category_list = List[Type]()
category_list.Add(Rebar)
category_list.Add(RebarInSystem)

category_filter = ElementMulticlassFilter(category_list)
elements = FilteredElementCollector(document, document.ActiveView.Id)\
    .WherePasses(category_filter)\
    .WhereElementIsNotElementType()\
    .ToElements()

rebar_elements = [ RebarElement(element) for element in elements ]
rebar_elements = [ element for element in rebar_elements
                   if element.HostCategory.Id == ElementId(BuiltInCategory.OST_Walls)
                   or element.HostCategory.Id == ElementId(BuiltInCategory.OST_Columns)
                   or element.HostCategory.Id == ElementId(BuiltInCategory.OST_StructuralColumns)]

with Transaction(document) as transaction:
    transaction.Start("Обновление ориентации арматуры")

    for rebar in rebar_elements:
        host_mark = rebar.Element.GetParamValueOrDefault(BuiltInParameter.REBAR_ELEM_HOST_MARK)
        rebar.Element.SetParamValue("Мрк.МаркаКонструкции", host_mark)
        if rebar.IsAllowProcess:
            structure_mark = "{}{}".format(host_mark, "_Вертик" if rebar.ZIsLonger else "_Гориз")
            rebar.Element.SetParamValue("Мрк.МаркаКонструкции", structure_mark)

    transaction.Commit()