#!/usr/bin/env python
from distutils.core import setup

version='1.0'

setup(
    name='django-excel-response',
    version=version,
    author='Tarken',
    author_email='?',
#    maintainer = 'Mikhail Korobov',
#    maintainer_email='kmike84@gmail.com',

    packages=['excel_response'],

    url='http://bitbucket.org/kmike/django-excel-response/',
    download_url = 'http://bitbucket.org/kmike/django-excel-response/get/tip.zip',
    description = """A subclass of HttpResponse which will transform a QuerySet,
or sequence of sequences, into either an Excel spreadsheet or
CSV file formatted for Excel, depending on the amount of data.

http://djangosnippets.org/snippets/1151/
""",

    long_description = open('README.rst').read(),

    requires = ['xlwt'],

    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Environment :: Web Environment',
        'Framework :: Django',
        'Intended Audience :: Developers',
        'Programming Language :: Python',
        'Topic :: Software Development :: Libraries :: Python Modules',
        'Topic :: Office/Business :: Financial :: Spreadsheet',

    ],
)
