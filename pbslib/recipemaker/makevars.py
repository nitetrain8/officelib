"""

Created by: Nathan Starkweather
Created on: 02/05/2014
Created in: PyCharm Community Edition


"""
from io import StringIO
from collections import OrderedDict
from os.path import dirname
import re
from pysrc.snippets.gen_tools import make_header


__author__ = 'Nathan Starkweather'

__LOGGER_VARS_FILE__ = "C:\\Users\\PBS Biotech\\Documents\\Personal\\test files\\logger settings.cfg"
__VARS_CSV__ = "\\".join((dirname(__file__), 'recipe_vars.csv'))
__VARS_PYFILE__ = '\\'.join((dirname(__file__), 'recipe_vars.py'))

# Public copies in case caller wants to see them.
LOGGER_VARS_FILE = __LOGGER_VARS_FILE__
VARS_CSV = __VARS_CSV__
VARS_PYFILE = __VARS_PYFILE__


def vars_from_logger_settings(fpath=__LOGGER_VARS_FILE__):
    """
    @param fpath: file path to logger settings file
    @type fpath: str
    @return: dict of pynames and recipe variable objects
    @rtype: dict[str, pbslib.recipemaker.RecipeVariable]

    Load a dict of pynames <=> RecipeVariables from a logger settings file.
    """
    from .recipemaker import RecipeVariable
    logger_vars = load_logger_vars(fpath)
    pynames = raw_to_pynames(logger_vars)
    recipe_vars = {pyname : RecipeVariable(varname) for pyname, varname in pynames.items()}
    return recipe_vars


def extract_csv_vars(csvfile=__VARS_CSV__):
    """
    @param csvfile: filepath to csv
    @type csvfile: str
    @return: __generator[(str, str)] of (pyname, varname)
    @rtype: __generator[(str, str)]
    """
    with open(csvfile, 'r') as f:
        txt = (line.split(',') for line in f.read().splitlines())
    return txt


def load_vars(fpath=__VARS_CSV__):
    """
    @param fpath: filepath to csv mapping pynames to varnames
    @type fpath: str
    @return: OrderedDict
    @rtype: OrderedDict[str, pbslib.recipemaker.RecipeVariable]

    Load dict of pynames <=> RecipeVariables from database.
    """

    from .recipemaker import RecipeVariable

    txt = extract_csv_vars(fpath)

    rvars = OrderedDict()
    for pyname, varname in txt:
        rvars[pyname] = RecipeVariable(varname)

    return rvars


def vars_py_current(csvfile=__VARS_CSV__, pyfile=__VARS_PYFILE__):
    """
    @param csvfile: csvfile of varnames
    @type csvfile: str
    @param pyfile: pyfile corresponding to csvfile
    @type pyfile: str
    @return: True if vars py is up to date.
    @rtype: bool
    """
    from os.path import getmtime as mtime
    try:
        return mtime(csvfile) < mtime(pyfile)
    except FileNotFoundError:
        return False


def make_vars_py(csvfile=__VARS_CSV__, pyfile=__VARS_PYFILE__):
    """
    @param csvfile: csv file with vars to make py from
    @type csvfile: str
    @param pyfile: filepath to write to
    @type pyfile: str
    @return: None
    @rtype: None
    """

    # Solve unicode escape issues in generated text

    pyfile = pyfile.replace("\\", "/")
    csvfile = csvfile.replace("\\", "/")

    rvars = extract_csv_vars(csvfile)
    tmp = StringIO()

    srcfile = __file__
    from os.path import split as path_split
    module = path_split(srcfile)[1]

    msg_body = """Python file to map pynames to recipe variable instances.
The list of variables should (almost) never change, except if
A different protocol is used to generate pynames from var names.
Instead of needing to import and generate at runtime, just import
this python file. This file will automatically raise ValueError
if it detects that the corresponding csv file is out of date."""

    header = make_header(srcfile, msg_body)
    tmp.write(header)

    tmp.write("from os.path import getmtime as _getmtime\n")
    tmp.write("\n# Filenames as passed to makevars.make_vars_py()\n")
    tmp.write("__module = r\"%s\"\n" % pyfile)
    tmp.write("__csv = r\"%s\"\n" % csvfile)
    tmp.write("\nif _getmtime(__csv) > _getmtime(__module):\n")
    tmp.write("    raise ValueError(\"Warning! Pyfile is outdated. Regenerate from %s.\")\n\n" % module)

    tmp.write("from officelib.pbslib.recipemaker import RecipeVariable\n\n\n")

    for pyname, varname in rvars:
        tmp.write(pyname)
        tmp.write(' = RecipeVariable("')
        tmp.write(varname)
        tmp.write('")\n')
    with open(pyfile, 'w') as f:
        f.write(tmp.getvalue())


def load_logger_vars(log_file=__LOGGER_VARS_FILE__):
    """
    @param log_file: log_file containing an example logger log_file,
                        from which to extract all settings.
    @tlog_filefile: str
    @return: list of all settings, as would be found in a recipe.
    @rtype: list[str]
    """
    with open(log_file, 'r') as f:
        f.readline()  # discard header
        body = f.read()

    # tab delimited columns
    line_start_to_tab = r"^([^\t]*)"
    vars = [v for v in re.findall(line_start_to_tab, body, re.MULTILINE) if v.strip()]
    return vars


def raw_to_pynames(vars):
    """
    @param vars: list of strings representing vars
    @type vars: list[str]
    @return: dict mapping python-legal name to var
    @rtype: dict[str, str]
    Map raw variable names to a python-legal and
    relatively equivalent name.
    """

    # match whitespace, periods, "&", stuff in parenthesis
    str_to_pyname = r"[\s*\.\&]|\(.*?\)"
    pynames = {re.sub(str_to_pyname, '', var) : var for var in vars}
    return pynames


# __VARS_CACHE__ = load_vars(__VARS_CSV__)


# Rebuild vars py if it was outdated.
def update_vars_py(csvfile=__VARS_CSV__, pyfile=__VARS_CSV__):
    """
    @param csvfile: csvfile of pynames <=> variable names
    @type csvfile: str
    @param pyfile: python file generated from csvfile with make_vars_py
    @type pyfile: str
    @return: None
    @rtype: None
    """
    if not vars_py_current(csvfile, pyfile):
        make_vars_py(csvfile, pyfile)


# debug
if __name__ == '__main__':

    from os.path import split
    base, name = split(__file__)
    outfile = '/'.join((base, 'vartest.txt'))
    outfile = __VARS_CSV__
    make_vars_py()
    # jsonlist = [item for item in sorted(vars_from_logger_settings().items(), key=lambda x:x[1])]
    # from io import StringIO

    # buf = StringIO()
    # for pyname, rvar in jsonlist:
    #     buf.write(pyname)
    #     buf.write(',')
    #     buf.write(rvar.Name)
    #     buf.write('\n')

    # print(buf.getvalue())
    # with open(outfile, 'w') as f:
    #     f.write(buf.getvalue())
    #
    # startfile(outfile)
