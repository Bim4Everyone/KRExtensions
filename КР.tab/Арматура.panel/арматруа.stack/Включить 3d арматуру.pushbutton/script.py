# -*- coding: utf-8 -*-

from pyrevit import EXEC_PARAMS
from dosymep_libs.bim4everyone import *

from rebars import set_solid_in_view

@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    application = __revit__.Application
    document = __revit__.ActiveUIDocument.Document

    set_solid_in_view(application, document, True)


script_execute()