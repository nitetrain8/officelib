"""
Created on Jan 21, 2014

@Company: PBS Biotech
@Author: Nathan Starkweather

Some helper classes/proxies for
making excel easier. Probably won't
be used much.
"""

from xllib.xladdress import cellRangeStr


class ChartSeries():
    """
    @type series_name: str
    @type start_row: int
    @type end_row: int
    @type x_column: int
    @type y_column: int
    @type sheet_name: str
    @type chart_name: str
    """

    def __init__(self,
                 series_name='',
                 start_row=1,
                 end_row=2,
                 x_column=1,
                 y_column=2,
                 sheet_name='Sheet1',
                 chart_name=''):

        """
        @type series_name: str
        @type start_row: int
        @type end_row: int
        @type x_column: int
        @type y_column: int
        @type sheet_name: str
        @type chart_name: str
        """

        self.series_name = series_name
        self.start_row = start_row
        self.end_row = end_row
        self.x_column = x_column
        self.y_column = y_column
        self.sheet_name = sheet_name
        self.chart_name = chart_name

    @property
    def xSeriesRange(self, cellRangeStr=cellRangeStr):
        return "=%s!%s" % (self.sheet_name, cellRangeStr(
                                                    (self.start_row, self.x_column, 1, 1),
                                                    (self.end_row, self.x_column, 1, 1)
                                                    ))

    @property
    def ySeriesRange(self, cellRangeStr=cellRangeStr):
        return "=%s!%s" % (self.sheet_name, cellRangeStr(
                                                    (self.start_row, self.y_column, 1, 1),
                                                    (self.end_row, self.y_column, 1, 1)
                                                    ))

    @property
    def SeriesName(self):
        """ If series name undefined, return formula
        for contents of top cell of Y column.

        Otherwise return series name.
        """
        return self.series_name
        # series_name = self.series_name
        # if not series_name:
        #     name_row = max(self.start_row - 1, 1)  # avoid negatives or 0
        #     name = "=%s!" % self.sheet_name
        #     name += cellStr(name_row, self.y_column)
        # else:
        #     name = series_name
        # return name

    @property
    def ChartName(self):
        return self.chart_name

    ChartTitle = ChartName
