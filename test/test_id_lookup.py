"""
Test the open_jpm.py
"""

import unittest2
from jpm.utility import get_current_path
from jpm.id_lookup import get_investment_Ids, InvalidPortfolioId, \
                            InvestmentIdNotFound, initialize_investment_lookup



class TestLookup(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestLookup, self).__init__(*args, **kwargs)

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



    def test_get_investment_Ids(self):
        """
        Read the date
        """
        lookup_file = get_current_path() + '\\samples\\sample_lookup.xls'
        initialize_investment_lookup(lookup_file)

        geneva_investment_id_for_HTM, isin, bloomberg_figi = \
            get_investment_Ids('11490', 'ISIN', 'xyz')

        self.assertEqual(geneva_investment_id_for_HTM, '')
        self.assertEqual(isin, 'xyz')
        self.assertEqual(bloomberg_figi, '')

        geneva_investment_id_for_HTM, isin, bloomberg_figi = \
            get_investment_Ids('11490', 'CMU', 'HSBCFN13014')

        self.assertEqual(geneva_investment_id_for_HTM, '')
        self.assertEqual(isin, 'HK0000163607')
        self.assertEqual(bloomberg_figi, '')

        geneva_investment_id_for_HTM, isin, bloomberg_figi = \
            get_investment_Ids('11490', 'JPM', '4C0198S')

        self.assertEqual(geneva_investment_id_for_HTM, '')
        self.assertEqual(isin, '')
        self.assertEqual(bloomberg_figi, '<to be determined>')

        geneva_investment_id_for_HTM, isin, bloomberg_figi = \
            get_investment_Ids('12548', 'ISIN', 'xyz')

        self.assertEqual(geneva_investment_id_for_HTM, 'xyz HTM')
        self.assertEqual(isin, '')
        self.assertEqual(bloomberg_figi, '')

        geneva_investment_id_for_HTM, isin, bloomberg_figi = \
            get_investment_Ids('12548', 'CMU', 'HSBCFN13014')

        self.assertEqual(geneva_investment_id_for_HTM, 'HK0000163607 HTM')
        self.assertEqual(isin, '')
        self.assertEqual(bloomberg_figi, '')

        geneva_investment_id_for_HTM, isin, bloomberg_figi = \
            get_investment_Ids('12548', 'CMU', 'WLHKFN09007')

        self.assertEqual(geneva_investment_id_for_HTM, 'CMU_WLHKFN09007 HTM')
        self.assertEqual(isin, '')
        self.assertEqual(bloomberg_figi, '')



    def test_error1(self):
        lookup_file = get_current_path() + '\\samples\\sample_lookup.xls'
        initialize_investment_lookup(lookup_file)

        with self.assertRaises(InvalidPortfolioId):
            get_investment_Ids('88888', 'ISIN', 'xyz')



    def test_error2(self):
        lookup_file = get_current_path() + '\\samples\\sample_lookup.xls'
        initialize_investment_lookup(lookup_file)

        with self.assertRaises(InvestmentIdNotFound):
            get_investment_Ids('11490', 'CMU', '12345678')