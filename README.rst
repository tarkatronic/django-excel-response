=====================
django-excel-response
=====================

This is http://djangosnippets.org/snippets/1151/ uploaded to pypi.
Author is Tarken.

A subclass of HttpResponse which will transform a QuerySet,
or sequence of sequences, into either an Excel spreadsheet or
CSV file formatted for Excel, depending on the amount of data.
All of this is done in-memory and on-the-fly, with no disk writes,
thanks to the StringIO library.

Installation
============

::

    pip install excel_response xlwt


Usage
=====

::

    from excel_response import ExcelResponse

    def excelview(request):
        objs = SomeModel.objects.all()
        return ExcelResponse(objs)


or::

    from excel_response import ExcelResponse

    def excelview(request):
        data = [
            ['Column 1', 'Column 2'],
            [1,2]
            [23,67]
        ]
        return ExcelResponse(data, 'my_data')
