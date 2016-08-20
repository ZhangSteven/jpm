"""
Test the open_jpm.py
"""

import unittest2
import datetime
from xlrd import open_workbook
from jpm.utility import get_current_path
from jpm.open_jpm import read_jpm, read_date

class TestJPM(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestJPM, self).__init__(*args, **kwargs)

    def setUp(self):
        """
            Run before a test function
        """
        pass



    def tearDown(self):
        """
            Run after a test finishes
        """
        pass



    def test_read_date_error(self):
        """
        Read the date
        """
        filename = get_current_path() + '\\samples\\date_error.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Sheet1')
        row = 0

        # same as the above
        with self.assertRaisesRegexp(ValueError, 'invalid date format'):
            n, d = read_date(ws, row)



    def test_read_date_error2(self):
        """
        Read the date
        """
        filename = get_current_path() + '\\samples\\date_error2.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Sheet1')
        row = 0

        # same as the above
        with self.assertRaisesRegexp(ValueError, 'invalid date_string'):
            n, d = read_date(ws, row)



    def test_read_date_error3(self):
        """
        Read the date
        """
        filename = get_current_path() + '\\samples\\date_error3.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Sheet1')
        row = 0

        # same as the above
        with self.assertRaisesRegexp(ValueError, '.*is out of range for.*'):
            n, d = read_date(ws, row)