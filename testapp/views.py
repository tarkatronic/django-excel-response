from excel_response import ExcelView

from .models import TestModel


class TestView(ExcelView):
    model = TestModel
