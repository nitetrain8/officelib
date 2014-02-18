"""

Created by: Nathan Starkweather
Created on: 02/14/2014
Created in: PyCharm Community Edition


"""
from collections import OrderedDict

from officelib.olutils import getFullLibraryPath
from officelib.pbslib.batchbase import BatchBase, BatchError
from officelib.pbslib.batchutil import extract_data_report, group_header_data
from officelib.pbslib.proxies import Parameter


class DataReportError(BatchError):
    pass


class DataReport(BatchBase):

    """ Class to hold data from a data report from PBS Bioreactor
    Batch File.

    Access parameters by dict-style lookup: batch[Parameter]

    Parameters are instances of class Parameter.

    Constructors:

    DataReport() -> empty data report
    DataReport(filename) -> Data report extracting data from the filename
    DataReport.fromMapping(mapping) -> Data report from mapping of parameter names, parameter instances
    DataReport.fromIterable(iterable) -> Data report from list of parameter instances

    """

    def __init__(self, filename=None):
        super().__init__()
        self.__mapping = OrderedDict()
        if filename:
            self.ProcessFile(filename)
        else:
            self._filename = None

    # Implement limited dict-like interface
    # Redirected through self.__mapping

    def __getitem__(self, key):
        return self.__mapping[key]

    def __contains__(self, key):
        return key in self.__mapping

    def __iter__(self):
        return iter(self.__mapping)

    def __setitem__(self, key, value):
        self.__mapping[key] = value

    def __delitem__(self, key):
        del self.__mapping[key]

    def get(self, key, default=None):
        return self.__mapping.get(key, default)

    def ProcessFile(self, filename):
        """
        @param filename: filename to process
        @type filename: str
        @return: None
        @rtype: None
        """
        self._filename = filename
        full_filename = getFullLibraryPath(filename)
        self._full_filename = full_filename
        self.__process_filename(full_filename)

    def __process_filename(self, filename):
        """
        @param filename: filename to process
        @type filename: str
        @return: None
        @rtype: None
        """

        headers, raw_data = extract_data_report(filename)

        for header, (times, pvs, _empty) in group_header_data(headers, raw_data):

            try:
                parameter = Parameter(header, times, pvs)
                parameter.Parent = self
            except Exception as e:
                # Catch errors during creation to reraise with filename of problem.
                raise DataReportError("Error occurred trying to make %s in batch file " % header + self.Filename) from e

            self[header] = parameter

    @classmethod
    def fromMapping(cls, mapping, filename=None):
        """
        @param mapping: mapping of parameter names to parameter instances
        @type mapping: collections.Mapping[str, Parameter]
        @param filename: optional filename.
        @type filename: str
        @return: DataReport
        @rtype: DataReport[str, Parameter]
        """
        self = cls()
        for key in mapping:
            self[key] = mapping[key]
        self._filename = filename
        return self

    @classmethod
    def fromIterable(cls, iterable, filename=None):
        """
        @param iterable: iterable of parameter instances
        @type iterable: collections.Iterable[Parameter]
        @param filename: optional filename.
        @type filename: str
        @return: DataReport
        @rtype: DataReport
        """

        self = cls()
        for parameter in iterable:
            self[parameter.Header] = parameter

        self._filename = filename
        return self

    @property
    def Filename(self):
        return self._full_filename
