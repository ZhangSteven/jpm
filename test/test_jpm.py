"""
Test the open_jpm.py
"""

import unittest2
import datetime
from xlrd import open_workbook
from jpm.utility import get_current_path
from jpm.open_jpm import read_jpm, read_date, extract_account_info, \
                            read_holding_fields

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



    def test_syntax(self):
        """
        A trivial test function, make sure the code has no syntax error.
        """
        self.assertEqual(1, 1)



    def test_read_date(self):
        """
        Read the date
        """
        filename = get_current_path() + '\\samples\\statement.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Sheet1')
        row = 0

        n, d = read_date(ws, row)

        self.assertEqual(n, 5)  # it's in row A6
        self.assertEqual(d, datetime.datetime(2016,7,6))



    def test_extract_account_info(self):
        cell_value = \
            'Account:   48029   CLT - CLI HK BR (CLASS A-HK) TRUST FUND  '

        account_code, account_name = extract_account_info(cell_value)
        self.assertEqual(account_code, '48029')
        self.assertEqual(account_name, 'CLT - CLI HK BR (CLASS A-HK) TRUST FUND')



    def test_read_holding_fields(self):
        filename = get_current_path() + '\\samples\\holding_field_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Sheet1')
        row = 8 # holding field starts at A9

        rows_each_holding, coordinates, fields = read_holding_fields(ws, row)
        self.assertEqual(rows_each_holding, 4)
        self.assertEqual(len(coordinates), len(fields))
        self.assertEqual(len(fields), 18)

        i = 0
        for fld in fields:
            if fld == 'security_id':
                coordinates[i] == (0, 0)
            elif fld == 'security_name':
                coordinates[i] == (0, 1)
            elif fld == 'location_or_nominee':
                coordinates[i] == (0, 4)
            elif fld == 'awaiting_receipt':
                coordinates[i] == (0, 5)
            elif fld == 'settled_units':
                coordinates[i] == (0, 6)
            elif fld == 'total_units':
                coordinates[i] == (0, 7)
            elif fld == 'isin':
                coordinates[i] == (1, 0)
            elif fld == 'regional_or_sub_account':
                coordinates[i] == (1, 4)
            elif fld == 'awaiting_delivery':
                coordinates[i] == (1, 5)
            elif fld == 'current_face_settled':
                coordinates[i] == (1, 6)
            elif fld == 'current_face_total':
                coordinates[i] == (1, 7)
            elif fld == 'occ_id':
                coordinates[i] == (2, 0)
            elif fld == 'coupon_rate':
                coordinates[i] == (2, 1)
            elif fld == 'maturity_date':
                coordinates[i] == (2, 2)
            elif fld == 'pool_number':
                coordinates[i] == (2, 3)
            elif fld == 'country':
                coordinates[i] == (2, 4)
            elif fld == 'collateral_units':
                coordinates[i] == (2, 5)
            elif fld == 'borrowed_units':
                coordinates[i] == (3, 5)
            else:
                # field in not any of the above,
                # something must be wrong
                self.assertEqual(0, 1)

            i = i + 1
            # end of for loop
