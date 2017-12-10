from __future__ import absolute_import, unicode_literals

from collections import OrderedDict

import openpyxl
import six
from django.test import TestCase
from openpyxl.styles import Font

from excel_response import response

from .models import TestModel


class ExcelResponseCSVTest(TestCase):

    def test_force_csv(self):
        r = response.ExcelResponse([['a']], force_csv=True)
        # Call this so that all of the data gets resolved
        r.content
        self.assertEqual(r['content-type'], 'text/csv; charset=utf8')

    def test_exceeding_row_limit_with_list_creates_csv(self):
        old_limit = response.ROW_LIMIT
        response.ROW_LIMIT = 2
        r = response.ExcelResponse(
            [['a'], ['b'], ['c']]
        )
        # Call this so that all of the data gets resolved
        r.content
        self.assertEqual(r['content-type'], 'text/csv; charset=utf8')
        response.ROW_LIMIT = old_limit

    def test_exceeding_column_limit_with_list_creates_csv(self):
        old_limit = response.COL_LIMIT
        response.COL_LIMIT = 2
        r = response.ExcelResponse(
            [['a', 'b', 'c']]
        )
        # Call this so that all of the data gets resolved
        r.content
        self.assertEqual(r['content-type'], 'text/csv; charset=utf8')
        response.COL_LIMIT = old_limit

    def test_exceeding_row_limit_with_model_creates_csv(self):
        old_limit = response.ROW_LIMIT
        response.ROW_LIMIT = 2
        TestModel.objects.create(text='a', number='1')
        TestModel.objects.create(text='b', number='2')
        TestModel.objects.create(text='c', number='3')
        r = response.ExcelResponse(
            TestModel.objects.all()
        )
        # Call this so that all of the data gets resolved
        r.content
        self.assertEqual(r['content-type'], 'text/csv; charset=utf8')
        response.ROW_LIMIT = old_limit

    def test_exceeding_column_limit_with_model_creates_csv(self):
        old_limit = response.COL_LIMIT
        response.COL_LIMIT = 2
        TestModel.objects.create(text='a', number='1')
        r = response.ExcelResponse(
            TestModel.objects.all()
        )
        # Call this so that all of the data gets resolved
        r.content
        self.assertEqual(r['content-type'], 'text/csv; charset=utf8')
        response.COL_LIMIT = old_limit

    def test_csv_from_list(self):
        r = response.ExcelResponse(
            [['a', 'b', 'c'], [1, 2, 3], [4, 5, 6]],
            force_csv=True
        )
        output = r.getvalue()
        self.assertEqual(
            output,
            b'a,b,c\r\n1,2,3\r\n4,5,6\r\n'
        )

    def test_csv_from_list_of_dicts(self):
        r = response.ExcelResponse(
            [
                OrderedDict([('a', 1), ('b', 2), ('c', 3)]),  # OrderedDict ensures the order of our headers & output
                {'a': 4, 'b': 5, 'c': 6}
            ],
            force_csv=True
        )
        output = r.getvalue()
        self.assertEqual(
            output,
            b'a,b,c\r\n1,2,3\r\n4,5,6\r\n'
        )

    def test_queryset_values_limits_output_columns(self):
        TestModel.objects.create(text='a', number='1')
        TestModel.objects.create(text='b', number='2')
        TestModel.objects.create(text='c', number='3')
        r = response.ExcelResponse(
            TestModel.objects.all().values('text'),
            force_csv=True
        )
        output = r.getvalue()
        self.assertEqual(
            output,
            b'text\r\na\r\nb\r\nc\r\n'
        )

    def test_csv_output_with_empty_list(self):
        r = response.ExcelResponse([], force_csv=True)
        output = r.getvalue()
        self.assertEqual(output, b'')

    def test_csv_output_with_empty_queryset(self):
        r = response.ExcelResponse(TestModel.objects.none(), force_csv=True)
        output = r.getvalue()
        self.assertEqual(output, b'')


class ExcelResponseExcelTest(TestCase):

    def test_create_excel_from_list(self):
        r = response.ExcelResponse(
            [['a', 'b', 'c'], [1, 2, 3], [4, 5, 6]]
        )
        output = six.BytesIO(r.getvalue())
        # This should theoretically raise errors if it's not a valid spreadsheet
        wb = openpyxl.load_workbook(output, read_only=True)
        ws = wb.active
        self.assertEqual((ws['A1'].value, ws['B1'].value, ws['C1'].value), ('a', 'b', 'c'))
        self.assertEqual((ws['A2'].value, ws['B2'].value, ws['C2'].value), (1, 2, 3))

    def test_create_excel_from_list_of_dicts(self):
        r = response.ExcelResponse(
            [
                OrderedDict([('a', 1), ('b', 2), ('c', 3)]),  # OrderedDict ensures the order of our headers & output
                {'a': 4, 'b': 5, 'c': 6}
            ]
        )
        output = six.BytesIO(r.getvalue())
        openpyxl.load_workbook(output, read_only=True)

    def test_create_excel_from_queryset(self):
        TestModel.objects.create(text='a', number='1')
        TestModel.objects.create(text='b', number='2')
        TestModel.objects.create(text='c', number='3')
        r = response.ExcelResponse(
            TestModel.objects.all()
        )
        output = six.BytesIO(r.getvalue())
        openpyxl.load_workbook(output, read_only=True)

    def test_header_font_is_applied(self):
        f = Font(name='Windings')
        r = response.ExcelResponse(
            [['a', 'b', 'c'], [1, 2, 3], [4, 5, 6]],
            header_font=f
        )
        output = six.BytesIO(r.getvalue())
        book = openpyxl.load_workbook(output, read_only=True)
        sheet = book.active
        cell = sheet['A1']
        self.assertEqual(cell.font.name, 'Windings')

    def test_data_font_is_applied(self):
        f = Font(name='Windings')
        r = response.ExcelResponse(
            [['a', 'b', 'c'], [1, 2, 3], [4, 5, 6]],
            data_font=f
        )
        output = six.BytesIO(r.getvalue())
        book = openpyxl.load_workbook(output, read_only=True)
        sheet = book.active
        cell = sheet['A2']
        self.assertEqual(cell.font.name, 'Windings')


class CBVTest(TestCase):

    def test_excel_view(self):
        TestModel.objects.create(text='a', number='1')
        TestModel.objects.create(text='b', number='2')
        TestModel.objects.create(text='c', number='3')
        response = self.client.get('/test/')
        output = six.BytesIO(response.content)
        openpyxl.load_workbook(output, read_only=True)
