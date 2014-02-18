"""

Created by: Nathan Starkweather
Created on: 02/14/2014
Created in: PyCharm Community Edition


"""
from collections import OrderedDict

from officelib.olutils import getFullLibraryPath
from officelib.pbslib.batchbase import BatchBase, BatchError
from officelib.pbslib.batchutil import ExtractDataReport, GroupHeaderData
from officelib.pbslib.proxies import Parameter


class DataReportError(BatchError):
    pass


class DataReport(BatchBase):

    """ Class to hold data from a data report from PBS Bioreactor
    Batch File.

    Access parameters by dict-style lookup: batch[Parameter]

    Parameters are instances of class Parameter.

    constructors:
    DataReport() -> empty data report
    DataReport(filename=filename) -> Data report extracting data from the filename
    DataReport.fromMapping(mapping) -> mapping of parameter names, parameter instances

    """

    def __init__(self, filename=None):
        super().__init__()
        self.__mapping = OrderedDict
        if filename:
            self.ProcessFile(filename)
        else:
            self._filename = None

    # Implement limited dict-like interface
    # Redirected through self.__mapping

    def __getitem__(self, key):
        return self.__mapping[key]

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

        headers, raw_data = ExtractDataReport(filename)

        for header, (times, pvs, _empty) in GroupHeaderData(headers, raw_data):

            try:
                param = Parameter(header, times, pvs)
                param.Parent = self
            except Exception as e:
                # Catch errors during creation to reraise with filename of problem.
                raise DataReportError("Error occurred trying to make %s in batch file " % header + self.Filename) from e

            self[header] = param

    @classmethod
    def fromMapping(cls, mapping):
        """
        @param mapping: mapping of parameter names to parameter instances
        @type mapping: collections.Mapping[str, Parameter]
        @return: DataReport
        @rtype: DataReport[str, Parameter]
        """
        self = cls()
        for key in mapping:
            self[key] = mapping[key]
        self._filename = None
        return self

    @property
    def Filename(self):
        return self._full_filename
