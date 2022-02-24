# -*- coding: utf-8 -*-

from rebars import set_solid_in_view


application = __revit__.Application
document = __revit__.ActiveUIDocument.Document


set_solid_in_view(application, document, True)