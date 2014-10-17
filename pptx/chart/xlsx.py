# encoding: utf-8

"""
Chart builder and related objects.
"""

from __future__ import absolute_import, print_function, unicode_literals

from contextlib import contextmanager
from StringIO import StringIO as BytesIO

from xlsxwriter import Workbook


class WorkbookWriter(object):
    """
    Service object that knows how to write an Excel workbook for chart data.
    """
    @classmethod
    def xlsx_blob(cls, categories, series):
        """
        Return the byte stream of an Excel file formatted as chart data for
        a chart having *categories* and *series*.
        """
        xlsx_file = BytesIO()
        with cls._open_worksheet(xlsx_file) as worksheet:
            cls._populate_worksheet(worksheet, categories, series)
        return xlsx_file.getvalue()

    @staticmethod
    @contextmanager
    def _open_worksheet(xlsx_file):
        """
        Enable XlsxWriter Worksheet object to be opened, operated on, and
        then automatically closed within a `with` statement. A filename or
        stream object (such as a ``BytesIO`` instance) is expected as
        *xlsx_file*.
        """
        workbook = Workbook(xlsx_file, {'in_memory': True})
        worksheet = workbook.add_worksheet()
        yield worksheet
        workbook.close()

    @classmethod
    def _populate_worksheet(cls, worksheet, categories, series):
        """
        Write *categories* and *series* to *worksheet* in the standard
        layout, categories in first column starting in second row, and series
        as columns starting in column next to categories, series title in first
        cell. Make the whole range an Excel List.
        """
        if len(categories) != 0 and isinstance(categories[0], (list, tuple)):
            for ilvl in xrange(len(categories)):
                for idx, token in categories[ilvl]:
                    worksheet.write(1+idx, ilvl, token)
            value_start_col = len(categories)
        else:
            worksheet.write_column(1, 0, categories)
            value_start_col = 1
        for item_series in series:
            series_col = item_series.index + value_start_col
            worksheet.write(0, series_col, item_series.name)
            if len(item_series.values) != 0 and isinstance(
                item_series.values[0], tuple):
                for idx, token in item_series.values:
                    worksheet.write(1+idx, series_col, token)
            else:
                worksheet.write_column(1, series_col, item_series.values)
