from django.test import TestCase

from excel_response import response

from .models import TestModel


class ExcelResponseTest(TestCase):

    def test_force_csv(self):
        r = response.ExcelResponse([['a']], force_csv=True)
        self.assertEqual(r['content-type'], 'text/csv; charset=utf8')

    def test_exceeding_row_limit_with_list_creates_csv(self):
        old_limit = response.ROW_LIMIT
        response.ROW_LIMIT = 2
        r = response.ExcelResponse(
            [['a'], ['b'], ['c']]
        )
        self.assertEqual(r['content-type'], 'text/csv; charset=utf8')
        response.ROW_LIMIT = old_limit

    def test_exceeding_column_limit_with_list_creates_csv(self):
        old_limit = response.COL_LIMIT
        response.COL_LIMIT = 2
        r = response.ExcelResponse(
            [['a', 'b', 'c']]
        )
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
        self.assertEqual(r['content-type'], 'text/csv; charset=utf8')
        response.ROW_LIMIT = old_limit

    def test_exceeding_column_limit_with_list_creates_csv(self):
        old_limit = response.COL_LIMIT
        response.COL_LIMIT = 2
        TestModel.objects.create(text='a', number='1')
        r = response.ExcelResponse(
            TestModel.objects.all()
        )
        self.assertEqual(r['content-type'], 'text/csv; charset=utf8')
        response.COL_LIMIT = old_limit

    def test_header_font_is_applied(self):
        pass

    def test_data_font_is_applied(self):
        pass


class CBVTest(TestCase):
    pass
