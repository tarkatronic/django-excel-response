from __future__ import absolute_import, unicode_literals

import csv

import django
import six
from django.http.response import FileResponse
from openpyxl import Workbook
from openpyxl.writer.excel import save_virtual_workbook

if django.VERSION >= (1, 9):
    from django.db.models.query import QuerySet
else:
    from django.db.models.query import QuerySet, ValuesQuerySet


ROW_LIMIT = 1048576
COL_LIMIT = 16384


class ExcelResponse(FileResponse):
    """
    This class provides an HTTP Response in the form of an Excel spreadsheet, or CSV file.
    """

    def __init__(self, data, output_name='excel_data', force_csv=False, header_font=None, data_font=None, *args,
                 **kwargs):
        # We do not initialize this with streaming_content, as that gets generated when needed
        super(ExcelResponse, self).__init__(*args, **kwargs)
        self._raw_data = data
        self.output_name = output_name
        self.header_font = header_font
        self.data_font = data_font
        self.force_csv = force_csv

    @property
    def streaming_content(self):
        workbook = None
        if isinstance(self._raw_data, list):
            workbook = self._serialize_list(self._raw_data)
        elif isinstance(self._raw_data, QuerySet):
            workbook = self._serialize_queryset(self._raw_data)
        if django.VERSION < (1, 9):
            if isinstance(self._raw_data, ValuesQuerySet):
                workbook = self._serialize_values_queryset(self._raw_data)
        if workbook is None:
            raise ValueError('ExcelResponse accepts the following data types: list, dict, QuerySet, ValuesQuerySet')

        if self.force_csv:
            self['Content-Type'] = 'text/csv; charset=utf8'
            self['Content-Disposition'] = 'attachment;filename={}.csv'.format(self.output_name)
            workbook.seek(0)
        else:
            self['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            self['Content-Disposition'] = 'attachment; filename={}.xlsx'.format(self.output_name)
            workbook = save_virtual_workbook(workbook)
        return map(self.make_bytes, workbook)

    @streaming_content.setter
    def streaming_content(self, value):
        self._raw_data = value

    def _serialize_list(self, data):
        workbook = None
        if isinstance(data[0], dict):  # If we're dealing with a list of dicationaries, generate the headers
            headers = [key for key in data[0]]
        else:
            headers = data[0]
        if len(data) > ROW_LIMIT or len(headers) > COL_LIMIT or self.force_csv:
            self.force_csv = True
            workbook = six.StringIO()
            csvwriter = csv.writer(workbook, dialect='excel')
            append = getattr(csvwriter, 'writerow')
        else:
            workbook = Workbook()
            worksheet = workbook.active
            append = getattr(worksheet, 'append')
        if isinstance(data[0], dict):
            append(headers)
        for row in data:
            if isinstance(row, dict):
                append([row.get(col, None) for col in headers])
            else:
                append(row)
        return workbook

    def _serialize_queryset(self, data):
        return self._serialize_list(list(data.values()))

    def _serialize_values_queryset(self, data):
        return self._serialize_list(list(data))
