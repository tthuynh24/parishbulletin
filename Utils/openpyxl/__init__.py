# Copyright (c) 2010-2015 openpyxl

# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

"""Imports for the openpyxl package."""
from .xml import LXML

from .workbook import Workbook
from .reader.excel import load_workbook


# constants
__version__ = '2.2.1'

__author__ = 'Eric Gazoni'
__license__ = 'MIT/Expat'
__author_email__ = 'eric.gazoni@gmail.com'
__maintainer_email__ = 'openpyxl-users@googlegroups.com'
__url__ = 'http://openpyxl.readthedocs.org'
