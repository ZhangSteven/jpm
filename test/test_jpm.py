"""
Test the open_jpm.py
"""

import unittest2
import datetime
from xlrd import open_workbook
from jpm.utility import get_current_path
from jpm.open_jpm import read_jpm, read_date, extract_account_info, \
                            read_holding_fields, read_holding_position, \
                            read_holdings_total, validate_holdings_total

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
                coordinates[i] == (1, 1)
            elif fld == 'maturity_date':
                coordinates[i] == (1, 2)
            elif fld == 'pool_number':
                coordinates[i] == (1, 3)
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



    def test_read_holding_position(self):
        filename = get_current_path() + '\\samples\\holding_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Sheet1')
        
        row = 8 # the holding field starts at A9
        rows_each_holding, coordinates, fields = read_holding_fields(ws, row)
        
        holdings = []
        row = 13 # the equity holding field starts at A14
        read_holding_position(ws, row, coordinates, fields, holdings)

        self.assertEqual(len(holdings), 1)  # only one position there
        
        position = holdings[0]
        self.validate_equity_position(position)



    def test_read_holding_position2(self):
        filename = get_current_path() + '\\samples\\holding_sample.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Sheet1')
        
        row = 8 # the holding field starts at A9
        rows_each_holding, coordinates, fields = read_holding_fields(ws, row)
        
        holdings = []
        row = 18 # the bond holding field starts at A19
        read_holding_position(ws, row, coordinates, fields, holdings)

        self.assertEqual(len(holdings), 1)  # only one position there
        
        position = holdings[0]
        self.validate_bond_position(position)



    def test_read_holding_total(self):
        filename = get_current_path() + '\\samples\\statement.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Sheet1')
        row = 289 # look at the holdings total at E290

        n, holdings_total = read_holdings_total(ws, row)
        self.assertEqual(n, 2)
        self.assertEqual(len(holdings_total), 6)
        self.assertAlmostEqual(holdings_total['awaiting_receipt'], 0)
        self.assertAlmostEqual(holdings_total['settled_units'], 101500000)
        self.assertAlmostEqual(holdings_total['total_units'], 101500000)
        self.assertAlmostEqual(holdings_total['awaiting_delivery'], 0)
        self.assertAlmostEqual(holdings_total['current_face_settled'], 5000000)
        self.assertAlmostEqual(holdings_total['current_face_total'], 5000000)



    def test_validate_holding_position(self):
        filename = get_current_path() + '\\samples\\statement.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Sheet1')
        
        row = 210 # the holding field starts at A211
        rows_each_holding, coordinates, fields = read_holding_fields(ws, row)
        
        holdings = []
        row = 215 # the first position starts at A216
        for i in range(15): # there are 15 positions
            # print('row = {0}'.format(row))
            read_holding_position(ws, row, coordinates, fields, holdings)
            row = row + 5

        row = 289 # the holdings total at E290
        n, holdings_total = read_holdings_total(ws, row)

        try:
            validate_holdings_total(holdings, holdings_total)
        except:
            self.fail('validate_holdings_total() raises an exception')



    def test_read_holding_total2(self):
        filename = get_current_path() + '\\samples\\statement.xls'
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Sheet1')
        row = 192 # look at the holdings total at E193

        n, holdings_total = read_holdings_total(ws, row)
        self.assertEqual(n, 2)
        self.assertAlmostEqual(holdings_total['awaiting_receipt'], 1000000)
        self.assertAlmostEqual(holdings_total['settled_units'], 678902100)
        self.assertAlmostEqual(holdings_total['total_units'], 679746900)
        self.assertAlmostEqual(holdings_total['awaiting_delivery'], 155200)
        self.assertAlmostEqual(holdings_total['current_face_settled'], 0)
        self.assertAlmostEqual(holdings_total['current_face_total'], 0)



    def validate_equity_position(self, position):
        """
        Validate the equity position read from 'holding_sample.xls'
        """
        self.assertEqual(len(position), 12)

        self.assertEqual(position['security_id'], 'BY9D3L9')
        self.assertEqual(position['security_name'], '3SBIO INC COMMON STOCK HKD 0.00001')
        self.assertEqual(position['isin'], 'KYG8875G1029')
        self.assertEqual(position['regional_or_sub_account'], '002')
        self.assertEqual(position['location_or_nominee'], '0WX')
        self.assertEqual(position['country'], 'KY')
        self.assertEqual(position['awaiting_receipt'], 0)
        self.assertEqual(position['awaiting_delivery'], 0)
        self.assertEqual(position['collateral_units'], 0)
        self.assertEqual(position['borrowed_units'], 0)
        self.assertEqual(position['settled_units'], 9917000)
        self.assertEqual(position['total_units'], 9917000)



    def validate_bond_position(self, position):
        """
        Validate the equity position read from 'holding_sample.xls'
        """
        self.assertEqual(len(position), 14)

        self.assertEqual(position['security_id'], 'B8FPQB8')
        self.assertEqual(position['security_name'], '1MDB ENERGY LTD NOTES FIXED 5.99% 11/MAY/2022 USD 100000')
        self.assertEqual(position['isin'], 'XS0784926270')
        self.assertEqual(position['regional_or_sub_account'], '130')
        self.assertEqual(position['location_or_nominee'], '590')
        self.assertEqual(position['country'], 'MY')
        self.assertEqual(position['awaiting_receipt'], 0)
        self.assertEqual(position['awaiting_delivery'], 0)
        self.assertEqual(position['collateral_units'], 0)
        self.assertEqual(position['borrowed_units'], 0)
        self.assertEqual(position['settled_units'], 8000000)
        self.assertEqual(position['total_units'], 8000000)
        self.assertAlmostEqual(position['coupon_rate'], 5.99/100)
        self.assertEqual(position['maturity_date'], datetime.datetime(2022,5,11))