=====================
django-excel-response
=====================
.. image:: https://img.shields.io/pypi/v/django-excel-response.svg
   :target: https://pypi.python.org/pypi/django-excel-response
   :alt: Latest Version

.. image:: https://travis-ci.org/tarkatronic/django-excel-response.svg?branch=master
   :target: https://travis-ci.org/tarkatronic/django-excel-response
   :alt: Test/build status

.. image:: https://codecov.io/gh/tarkatronic/django-excel-response/branch/master/graph/badge.svg
   :target: https://codecov.io/gh/tarkatronic/django-excel-response
   :alt: Code coverage


A subclass of HttpResponse which will transform:
  - a QuerySet, or sequence of sequences, into either an Excel spreadsheet or
    CSV file formatted for Excel, depending on the amount of data;
  - a dictionary, where keys are names of spreadsheets and values are  sequences of sequences
    or QuerySets, into either an Excel file with multiply spreadsheets or CSV file formatted
    for Excel, depending on the amount of data;

Installation
============

::

    pip install django-excel-response

Provided Classes
================

* ``excel_response.response.ExcelResponse``

    Accepted arguments:

    * ``data``:
               - A queryset or list of lists from which to construct the output
               - A dictionary where keys are names of spreadsheets and values are querysets or lists of lists
                 from which to construct the spreadsheet
    * ``output_filename`` - The filename which should be suggested in the http response,
      minus the file extension (**default: excel_data**)
    * ``worksheet_name`` - The name of the worksheet inside the spreadsheet into which
      the data will be inserted (**default: None**) (ignored in case **data** is dictionary)
    * ``force_csv`` - A boolean stating whether to force CSV output (**default: False**)
    * ``header_font`` - The font to be applied to the header row of the spreadsheet;
      must be an instance of ``openpyxl.styles.Font`` (**default: None**)
    * ``data_font`` - The font to be applied to all data cells in the spreadsheet;
      must be an instance of ``openpyxl.styles.Font`` (**default: None**)

* ``excel_response.views.ExcelMixin``
* ``excel_response.views.ExcelView``

Examples
========

Function-based views
--------------------

You can construct your data from a queryset.

.. code-block:: python

    from excel_response import ExcelResponse


    def excelview(request):
        objs = SomeModel.objects.all()
        return ExcelResponse(objs)


Or you can construct your data manually.

.. code-block:: python

    from excel_response import ExcelResponse


    def excelview(request):
        data = [
            ['Column 1', 'Column 2'],
            [1,2]
            [23,67]
        ]
        return ExcelResponse(data, 'my_data')


In case you need multi spreadsheet file you need to construct dictionary:

.. code-block:: python

    from excel_response import ExcelResponse


    SOME_FIELD_NAME_LIST = [...]


    def excelview(request):
        result_dict = {}
        for some_field_name in SOME_FIELD_NAME_LIST:
            result_dict[some_field_name] = SomeModel.objects.filter(some_field=some_field_name)
        return ExcelResponse(result_dict)


Or you can construct your data manually.

.. code-block:: python

    from excel_response import ExcelResponse


    def excelview(request):
        data = {'Spreadsheet1 name': [
                ['Column 1', 'Column 2'],
                [1,2]
                [23,67]
            ],
            [
                ['Column 1', 'Column 2'],
                [3,4]
                [89,33]
            ]
        }
        return ExcelResponse(data, 'my_data')


Class-based views
-----------------

These are as simple as import and go!

.. code-block:: python

    from excel_response import ExcelView


    class ModelExportView(ExcelView):
        model = SomeModel
