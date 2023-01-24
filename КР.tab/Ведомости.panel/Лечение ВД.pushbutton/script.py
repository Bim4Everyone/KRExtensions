# -*- coding: utf-8 -*-
import clr

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

from System import Guid

from pyrevit import EXEC_PARAMS, revit

from Autodesk.Revit.DB import *

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep_libs.bim4everyone import *
from dosymep.Bim4Everyone.SharedParams import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    pass


script_execute()
