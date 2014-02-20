"""

Created by: Nathan Starkweather
Created on: 02/06/2014
Created in: PyCharm Community Edition


"""
__author__ = 'Nathan Starkweather'

from officelib.pbslib.recipemaker.recipemaker import Recipe, LongRecipe, RecipeVariable
from officelib.pbslib.recipemaker.makevars import vars_py_current, make_vars_py, VARS_CSV, VARS_PYFILE


# Use this to dynamically load variables at runtime from an unknown python module.
def import_py_vars(pyfile):
    """
    @param pyfile: pyfile to import
    @type pyfile: str
    @return: ModuleType
    @rtype: module
    """
    from os.path import split as path_split, splitext as path_splitext
    from sys import path as sys_path
    from importlib import import_module

    py_vars_dir, py_vars_name = path_split(pyfile)
    py_vars_name, ext = path_splitext(py_vars_name)
    sys_path.append(py_vars_dir)
    var_module = import_module(py_vars_name)
    return {k : v for k, v in var_module.__dict__.items() if not k.startswith("_")}

# Rebuild vars pyfile if not up to date.
if not vars_py_current(VARS_CSV, VARS_PYFILE):
    make_vars_py(VARS_CSV, VARS_PYFILE)


# Import vars and update global namespace
from .recipe_vars import *

if __name__ == '__main__':
    pass
