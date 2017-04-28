from __future__ import absolute_import, unicode_literals

from collections import OrderedDict

from django.test import TestCase

from excel_response import response

from .models import TestModel


class ExcelResponseCSVTest(TestCase):

    def test_force_csv(self):
        r = response.ExcelResponse([['a']], force_csv=True)
        # Call this so that all of the data gets resolved
        r.streaming_content
        self.assertEqual(r['content-type'], 'text/csv; charset=utf8')

    def test_exceeding_row_limit_with_list_creates_csv(self):
        old_limit = response.ROW_LIMIT
        response.ROW_LIMIT = 2
        r = response.ExcelResponse(
            [['a'], ['b'], ['c']]
        )
        # Call this so that all of the data gets resolved
        r.streaming_content
        self.assertEqual(r['content-type'], 'text/csv; charset=utf8')
        response.ROW_LIMIT = old_limit

    def test_exceeding_column_limit_with_list_creates_csv(self):
        old_limit = response.COL_LIMIT
        response.COL_LIMIT = 2
        r = response.ExcelResponse(
            [['a', 'b', 'c']]
        )
        # Call this so that all of the data gets resolved
        r.streaming_content
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
        r.streaming_content
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
        r.streaming_content
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


class ExcelResponseExcelTest(TestCase):

    def test_create_excel_from_list(self):
        pass

    def test_create_excel_from_list_of_dicts(self):
        pass

    def test_create_excel_from_queryset(self):
        pass

    def test_header_font_is_applied(self):
        pass

    def test_data_font_is_applied(self):
        pass


class CBVTest(TestCase):
    pass
