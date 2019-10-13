"""

Created by: Nathan Starkweather
Created on: 03/11/2014
Created in: PyCharm Community Edition

Module: test_module
Functions: test_functions

"""
import unittest
from os import makedirs
from os.path import dirname, join, exists, normpath, split, splitext
from shutil import rmtree
# noinspection PyUnresolvedReferences
from officelib.olutils import getFullFilename, _get_lib_path_no_extension, _get_lib_path_no_ctxt,\
    _get_lib_path_no_basename, _lib_path_search_dir_list_builder, _get_lib_path_parital_qualname

__author__ = 'PBS Biotech'

curdir = dirname(__file__)
test_dir = dirname(curdir)
test_temp_dir = join(test_dir, "temp")
temp_dir = join(test_temp_dir, "temp_dir_path")
test_input = join(curdir, "test_input")


def setUpModule():
    try:
        makedirs(temp_dir)
    except FileExistsError:
        pass


def tearDownModule():
    try:
        rmtree(temp_dir)
    except FileNotFoundError:
        pass


class TestGetFullFilename(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        """
        @return: None
        @rtype: None
        """

        temp_dirs = [join(temp_dir, f) for f in ("temp1", "temp2")]
        temp_dirs2 = ("temptemp1", "temptemp2")
        temp_temp = ("tmp1", "tmp2", "tmp3")

        files = []

        def make(file):
            with open(file, 'w') as _:
                pass

        for n, L1 in enumerate(temp_dirs):
            for i, L2 in enumerate(temp_dirs2):
                tmp_dir = join(L1, L2)

                try:
                    makedirs(tmp_dir)
                except FileExistsError:
                    pass

                for j, file in enumerate(temp_temp):
                    tmp_name = ''.join((file, str(n), str(i), str(j), ".tmp"))
                    tmp_file = join(tmp_dir, tmp_name)
                    make(tmp_file)
                    files.append(normpath(tmp_file))

        cls.files = files

    def test_get_full_filename_exact(self):
        """
        When inputing an exact file, get the exact filename
        back.

        @return: None
        @rtype: None
        """

        for file in self.files:
            if not exists(file):
                raise self.failureException("Temp directory unexpectedly absent")
            expected = normpath(file)
            result = getFullFilename(file, temp_dir)

            self.assertEqual(expected, result)

    def test_get_full_filename_no_directory_with_drive(self):
        """
        When sending in the filename with no directory,
        get the correct filepath

        @return:
        @rtype:
        """

        for file in self.files:
            if not exists(file):
                raise self.failureException("Temp directory unexpectedly absent")

            name, ext = splitext(file)
            expected = normpath(file)
            result = getFullFilename(name, temp_dir)

            self.assertEqual(expected, result)

            result2 = getFullFilename(name)
            self.assertEqual(expected, result2)

    def test_get_full_filename_name_only_ext(self):
        """
        The most common case
        @return:
        @rtype:
        """

        for file in self.files:
            base, name = split(file)

            expected = normpath(file)

            result = getFullFilename(name, temp_dir)

            self.assertEqual(expected, result)

            result2 = getFullFilename(name)
            self.assertEqual(expected, result2)

    def test_get_full_filename_dir_no_ctxt(self):
        """
        Search for a filename eg Foo/bar.txt
        with input "Foo/bar"

        @return:
        @rtype:
        """
        for file in self.files:
            if not exists(file):
                raise self.failureException("Temp directory unexpectedly absent")

            base, name = split(file)
            name, ext = splitext(name)

            name = name.rstrip('\\/')
            expected = normpath(file)

            # noinspection PyTypeChecker
            result = getFullFilename(name, temp_dir)
            self.assertEqual(expected, result)

            # noinspection PyTypeChecker
            result2 = getFullFilename(name)
            self.assertEqual(expected, result2)

    def test_get_lib_path_partial_name_ext(self):
        """
        @return:
        @rtype:
        """
        td = len(temp_dir)
        files = [f[td:] for f in self.files]
        for path, expected in zip(files, self.files):

            base, name = split(path)

            # test inner function
            direct_result = _get_lib_path_parital_qualname(name, base, (temp_dir,))
            self.assertEqual(expected, direct_result, base + path)

            # test public interface
            full_result = getFullFilename(path, temp_dir)
            self.assertEqual(expected, full_result)

            # results should be the same
            self.assertEqual(full_result, direct_result)

            # splitext returns '\' in front of name. remove that and assert again.
            noslash = _get_lib_path_parital_qualname(name, base.lstrip('\\/'), (temp_dir,))
            self.assertEqual(noslash, full_result)

    def test_get_lib_path_partial_name_noext(self):
        """
        @return:
        @rtype:
        """
        td = len(temp_dir)
        files = [splitext(f[td:])[0] for f in self.files]
        for path, expected in zip(files, self.files):

            base, name = split(path)

            # test inner function
            direct_result = _get_lib_path_parital_qualname(name, base, (temp_dir,))
            self.assertEqual(expected, direct_result)

            # test public interface
            full_result = getFullFilename(path, temp_dir)
            self.assertEqual(expected, full_result)

            # results should be the same
            self.assertEqual(full_result, direct_result)

            # splitext returns '\' in front of name. remove that and assert again.
            noslash = _get_lib_path_parital_qualname(name, base.lstrip('\\/'), (temp_dir,))
            self.assertEqual(noslash, full_result)


if __name__ == '__main__':
    unittest.main()
    # import cProfile
    # cProfile.run("unittest.main()")
