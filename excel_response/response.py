import openpyxl
from django.db.models.query import QuerySet
from django.http import FileResponse


ROW_LIMIT = 1048576
COL_LIMIT = 16384


class ExcelResponse(FileResponse):
    """
    This class provides an HTTP Response in the form of an Excel spreadsheet, or CSV file.
    """

    def __init__(self, data, force_csv=False, header_font=None, data_font=None, *args, **kwargs):
        super(ExcelResponse, self).__init__(*args, **kwargs)
        self._raw_data = data
        self.force_csv = force_csv
        self.header_font = header_font
        self.data_font = data_font

    @property
    def content(self):
        pass

    @content.setter
    def content(self, value):
        self._raw_data = value
