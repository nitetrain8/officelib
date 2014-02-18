"""

Created by: Nathan Starkweather
Created on: 02/17/2014
Created in: PyCharm Community Edition


Because special characters in the filepath to wincom's gen_py
cache are not valid python identifiers, the classes and modules
contained in that directory cannot be accessed directly through
python code, and thus cannot be accessed by PyCharm for
type hinting.

This module will handle functionality related to maintaining an
updated working directory containing exact copies of the contents
of the gen_py cache, except with valid (and simplified) python
identifiers.
"""

from os import listdir, mkdir, remove as os_remove
from os.path import isdir, join as path_join, getmtime, dirname, split as path_split
from re import compile as re_compile
from io import StringIO
from shutil import copytree, copy2 as _copy2


gen_py_basepath = "C:\\Python33\\Lib\\site-packages\\win32com\\gen_py"

typehint_dir = dirname(__file__) + '\\_typehint'

dir_whitelist = (
                 '\\'.join((gen_py_basepath, "00020813-0000-0000-C000-000000000046x0x1x6")),
                 '\\'.join((gen_py_basepath, "00020905-0000-0000-C000-000000000046x0x8x4"))
                )

time_file = typehint_dir + '\\lastupdate.txt'


def gen_py_flatten(filepath, isdir=isdir, path_join=path_join, getmtime=getmtime):
    """
    @param filepath: filepath to flatten
    @type filepath: str
    @return: filenames
    @rtype: __generator[(str, float)]

    Recursively return list of all files in dir, flattened.
    Internal helper.

    Iterates depth-first.
    """
    names = listdir(filepath)
    for name in names:
        path = path_join(filepath, name)
        if isdir(path):
            if path in dir_whitelist:
                yield from gen_py_flatten(path)
        else:
            yield path, getmtime(path)


def type_hint_flatten(filepath, isdir=isdir, path_join=path_join, getmtime=getmtime):
    """
    @param filepath: filepath to flatten
    @type filepath: str
    @return: filenames
    @rtype: __generator[(str, float)]

    Recursively return list of all files in dir, flattened.
    Internal helper.

    Iterates depth-first.

    Because I lack foresight, this is identical to the above function,
    but only operates on type hint folder,
    and exists because I need to flatten the type hint folder
    without checking if path is in the dir_whitelist.
    """
    names = listdir(filepath)
    for name in names:
        path = path_join(filepath, name)
        if isdir(path):
            yield from gen_py_flatten(path)
        else:
            yield path, getmtime(path)


def find_old_files(wincom, reference):
    """
    @param wincom: the dict of wincom files and timestamps
    @type wincom: dict[str, float]
    @param reference: the dict of typehint files and timestamps
    @type reference: dict[str, float]
    @return: list[(src, dst)]
    @rtype: list[(str, str)]
    """
    old = []
    for w_path, w_mtime in wincom.items():
        base, head = path_split(w_path)
        ref_name = '\\'.join((wincom_to_typehint(base), head))
        ref_mtime = reference.get(ref_name, 0)
        if w_mtime > ref_mtime:
            if ref_name.startswith("\\"):
                ref_name = ref_name.lstrip('\\')
            ref_target = '\\'.join((typehint_dir, ref_name))
            old.append((w_path, ref_target))
    return old


def update_files(old):
    rm = os_remove
    copy2 = _copy2
    for src, dst in old:
        try:
            rm(dst)
        except FileNotFoundError:
            pass
        copy2(src, dst)


def update_typehints():
    """
    @return: None
    @rtype: None

    Update the typehints folder
    This is the function that should be called at
    startup by xllib.
    """

    wincom = dict(gen_py_flatten(gen_py_basepath))
    reference = load_hint_timestamps()
    old = find_old_files(wincom, reference)

    if old:
        update_files(old)

    def py_path(fpath):
        base, head = path_split(fpath)
        ref_name = '\\'.join((wincom_to_typehint(base), head))
        return ref_name

    file_output = ((py_path(fpath), time) for fpath, time in wincom.items())
    write_hint_timestamps(file_output)


def wincom_to_typehint(fpath, base=gen_py_basepath):
    """
    @param fpath: filepath to pythonify. Directories only!
    @type fpath: str
    @return: str
    @rtype: str

    This function has gone through a lot of iterations.
    It kind of sucks. Basically, it strips off the gen py path,
    takes up to the last 5 characters of the resulting path,
    and replaces any invalid python identifiers with underscores.

    """
    name = fpath.lstrip(base)

    pyname_re = re_compile(r"[^_a-zA-Z0-9]")
    new_name = pyname_re.sub("_", name[-5:])
    return new_name


def load_hint_timestamps(time_file=time_file):
    with open(time_file, 'r') as f:
        timestamps = tuple(line.split(',') for line in f)

    stamps = {k : float(v) for k, v in timestamps}
    return stamps


def write_hint_timestamps(timestamps, foutput=time_file):
    """
    @param timestamps: mapping of names to timestamps
    @type timestamps: collections.Iterable([str, float])
    @return: None
    @rtype: None
    """
    buf = StringIO()
    for file, time in timestamps:
        buf.write(file)
        buf.write(',')
        buf.write(str(time))
        buf.write('\n')

    with open(foutput, 'w') as f:
        f.write(buf.getvalue())


def init_hint_timestamps():
    """
    @return: None
    @rtype: None

    Build the list of timestamps from the wincom folder
    for the first time. Do not use!
    """

    files = type_hint_flatten(typehint_dir)
    timestamps = ((fpath.lstrip(typehint_dir).replace('/', '\\'), time) for fpath, time in files)
    write_hint_timestamps(timestamps)


def init_hint_folder():
    """
    @return: None
    @rtype: None

    Build the hint folder for the first time. Do not use.
    """

    hint_dir = typehint_dir

    try:
        mkdir(hint_dir)
    except FileExistsError:
        pass

    for fldr in dir_whitelist:
        new_name = wincom_to_typehint(fldr)
        new = '\\'.join((typehint_dir, new_name))
        copytree(fldr, new)

    init_hint_timestamps()


if __name__ == '__main__':
    update_typehints()

