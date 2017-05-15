from django.views.generic.base import View
from django.views.generic.list import MultipleObjectMixin

from .response import ExcelResponse


class ExcelMixin(MultipleObjectMixin):
    response_class = ExcelResponse
    header_font = None
    data_font = None
    output_filename = None
    worksheet_name = None
    force_csv = False

    def get_header_font(self):
        """
        Return the font to be applied to the header row of the spreadsheet.
        """
        return self.header_font

    def get_data_font(self):
        """
        Return the font to be applied to all data cells in the spreadsheet.
        """
        return self.data_font

    def get_output_filename(self):
        """
        Return the name of the file to be generated, minus the file extension.

        For instance, an output filename of `'excel_data'` would return either
        `'excel_data.xlsx'` or `'excel_data.csv'`
        """
        return self.output_filename or 'excel_data'

    def get_worksheet_name(self):
        """
        Return the name of the worksheet into which the data will be inserted.
        """
        return self.worksheet_name

    def get_force_csv(self):
        """
        Return a boolean for whether or not to force CSV output.
        """
        return self.force_csv

    def get_context_data(self, **kwargs):
        """
        Provide an empty stub since these responses take no context.
        """
        return {}

    def render_to_response(self, context, *args, **kwargs):
        return self.response_class(
            self.get_queryset(),
            output_filename=self.get_output_filename(),
            worksheet_name=self.get_worksheet_name(),
            force_csv=self.get_force_csv(),
            header_font=self.get_header_font(),
            data_font=self.get_data_font()
        )


class ExcelView(ExcelMixin, View):
    """
    Return the results of a queryset as an Excel spreadsheet.
    """
    def get(self, request, *args, **kwargs):
        return self.render_to_response({})
