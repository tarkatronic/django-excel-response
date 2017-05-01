import os

from setuptools import setup

with open('VERSION', 'r') as vfile:
    VERSION = vfile.read().strip()

# allow setup.py to be run from any path
os.chdir(os.path.normpath(os.path.join(os.path.abspath(__file__), os.pardir)))

setup(
    name='django-excel-response',
    version=VERSION,
    author='Joey Wilhelm',
    author_email='tarkatronic@gmail.com',
    license='Apache',
    description='Django package to easily render Excel spreadsheets',
    long_description=open('README.rst', 'r').read().strip(),
    packages=['excel_response'],
    include_package_data=True,
    url='https://github.com/tarkatronic/django-excel-response',
    download_url='https://github.com/tarkatronic/django-excel-response/archive/master.tar.gz',
    install_requires=[
        'Django>=1.8',
        'openpyxl'
    ],

    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Environment :: Web Environment',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: Apache Software License',
        'Natural Language :: English',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Programming Language :: Python :: 3.6',
        'Framework :: Django',
        'Framework :: Django :: 1.8',
        'Framework :: Django :: 1.9',
        'Framework :: Django :: 1.10',
        'Framework :: Django :: 1.11',
        'Topic :: Software Development :: Libraries :: Python Modules',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
    ],
    zip_safe=False,
    test_suite='runtests.runtests'
)
