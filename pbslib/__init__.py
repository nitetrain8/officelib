"""

Created by: Nathan Starkweather
Created on: 02/14/2014
Created in: PyCharm Community Edition

PBSLib is mostly a collection of classes and functions
designed to work with the data reports produced by
PBS Biotech Bioreactors' Data Export functionality.


"""

from .datareport import DataReport


def open_data_report(fname):
    """
    @param fname: file name of data report to open
    @type fname: str
    @return: DataReport
    @rtype: DataReport
    """

    return DataReport(fname)
