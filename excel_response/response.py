from __future__ import absolute_import, unicode_literals

import csv

import django
import six
from django.http.response import HttpResponse
from openpyxl import Workbook
from openpyxl.writer.excel import save_virtual_workbook
from openpyxl.writer.write_only import WriteOnlyCell


if django.VERSION >= (1, 9):
    from django.db.models.query import QuerySet
else:
    from django.db.models.query import QuerySet, ValuesQuerySet


ROW_LIMIT = 1048576
COL_LIMIT = 16384


class ExcelResponse(HttpResponse):
    """
    This class provides an HTTP Response in the form of an Excel spreadsheet, or CSV file.
    """

    def __init__(self, data, output_filename='excel_data', worksheet_name=None, force_csv=False, header_font=None,
                 data_font=None, guess_types=True, *args, **kwargs):
        # We do not initialize this with streaming_content, as that gets generated when needed
        self.output_filename = output_filename
        self.worksheet_name = worksheet_name or 'Sheet 1'
        self.header_font = header_font
        self.data_font = data_font
        self.force_csv = force_csv
        self.guess_types = guess_types
        super(ExcelResponse, self).__init__(data, *args, **kwargs)

    @property
    def content(self):
        return b''.join(self._container)

    @content.setter
    def content(self, value):
        if not bool(value) or not len(value):  # Short-circuit to protect against empty querysets/empty lists/None, etc
            self._container = []
            return

        value = self.convert_from_queryset(value)
        if not self.force_csv:
            self.force_csv = self.check_force_csv(value)
        workbook = self.get_workbook(value)

        if workbook is None:
            raise ValueError('ExcelResponse accepts the following data types: list, dict, QuerySet, ValuesQuerySet')

        if self.force_csv:
            self['Content-Type'] = 'text/csv; charset=utf8'
            self['Content-Disposition'] = 'attachment;filename={}.csv'.format(self.output_filename)
            workbook.seek(0)
            workbook = self.make_bytes(workbook.getvalue())
        else:
            self['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            self['Content-Disposition'] = 'attachment; filename={}.xlsx'.format(self.output_filename)
            workbook = save_virtual_workbook(workbook)
        self._container = [self.make_bytes(workbook)]

    def convert_from_queryset(self, value):
        if isinstance(value, dict):
            for worksheet_name in value:
                value[worksheet_name] = self.convert_queryset_to_list(value[worksheet_name])
        else:
            value = self.convert_queryset_to_list(value)
        return value

    def convert_queryset_to_list(self, value):
        if isinstance(value, QuerySet):
            if len(value) and isinstance(value[0], dict):  # .values() returns a list of dicts
                value = list(value)
            else:
                value = list(value.values())

        if django.VERSION < (1, 9):
            if isinstance(value, ValuesQuerySet):
                value = list(value)
        return value

    def check_force_csv(self, value):
        if isinstance(value, list):
            return self.check_excel_limits(value)
        elif isinstance(value, dict):
            return any([self.check_excel_limits(value[x]) for x in value])

    def check_excel_limits(self, data):
        return len(data) > ROW_LIMIT or (len(data) and len(data[0]) > COL_LIMIT)

    def get_workbook(self, value):
        workbook = None
        if self.force_csv:
            if isinstance(value, dict):
                value = self.convert_dict_to_list(value)
            workbook = six.StringIO()
            csvwriter = csv.writer(workbook, dialect='excel')
            append_func = getattr(csvwriter, 'writerow')
            write_header_func = append_func
            workbook = self._serialize_list(value, workbook, append_func, write_header_func, csvwriter, worksheet=None)
        elif isinstance(value, list) or isinstance(value, dict):
            workbook = Workbook(write_only=True)
            workbook.guess_types = self.guess_types

            # Define custom functions for appending so that we can handle any formatting
            def append_func(data):
                return self._append_excel_row(worksheet, data, header=False)

            def write_header_func(data):
                return self._append_excel_row(worksheet, data, header=True)

            if isinstance(value, list):
                worksheet = workbook.create_sheet(title=self.worksheet_name)
                workbook = self._serialize_list(value, workbook, append_func, write_header_func, None, worksheet)
            elif isinstance(value, dict):
                for worksheet_name in value:
                    # If we're dealing with a list of dictionaries, generate the headers
                    worksheet = workbook.create_sheet(title=str(worksheet_name))
                    workbook = self._serialize_list(
                        value[worksheet_name],
                        workbook,
                        append_func,
                        write_header_func,
                        None,
                        worksheet
                    )

        return workbook

    def convert_dict_to_list(self, value):
        converted_to_list = []
        for key in value:
            worksheet_list = [[], [], [key], []]
            worksheet_list.extend(value[key])
            converted_to_list.extend(worksheet_list)
        return converted_to_list

    def _serialize_list(self, data, workbook, append_func, write_header_func, csvwriter=None, worksheet=None):
        if not len(data):
            return workbook

        if isinstance(data[0], dict):  # If we're dealing with a list of dictionaries, generate the headers
            headers = [key for key in data[0]]
        else:
            headers = data[0]

        if isinstance(data[0], dict):
            append_func(headers)
        for index, row in enumerate(data):
            if isinstance(row, dict):
                write_header_func([row.get(col, None) for col in headers])
            else:
                if index > 0:
                    append_func(row)
                else:
                    write_header_func(row)
        return workbook

    def _append_excel_row(self, worksheet, data, header=False):
        if header:
            font = self.header_font
        else:
            font = self.data_font

        if not font:
            row = data
        else:
            row = []
            for cell in data:
                cell = WriteOnlyCell(worksheet, cell)
                cell.font = font
                row.append(cell)

        worksheet.append(row)
