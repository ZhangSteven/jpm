# jpm

This is to convert the JPM broker statement into a file containing investment positions and cash to reconcile with Advent Geneva system.

The positions are trade day positions, cash is settlement day cash.

++++++++++
How to use
++++++++++

Copy a JPM broker statement, similar to samples/statement.xls, to the current directory, then run
	
	python open_jpm.py <broker_statement_file>

The above will generate two files: cash.csv and holdings.csv, to be used for Geneva system reconciliation.

To run unit test, run

	nose2


++++++++++
Todo
++++++++++



++++++++++
ver 0.21
++++++++++
1. Bug fix: filename prefix now depends on the input directory instead of the output directory, so that jpm files from different input directories won't overwrite with each other.



++++++++++
ver 0.20
++++++++++
1. Add an output directory parameter to the write_csv() function, so that it can write to a different directory than the input directory. If the output directory is None (default), then it still writes to the input directory where the input files are read. This way it remains backward compatible to 0.19 if running in standalone mode. (Python open_jpm.py <input_file>)

2. The change is made to work with recon_helper package.



++++++++++
ver 0.1901
++++++++++
1. Change config file so that input directory is for office PC, previously it was for hong kong home laptop.



++++++++++
ver 0.19
++++++++++
1. The csv output filename solely depends on the input directory folder name, and will always contains 'jpm'. If the folder name is "ListCo Equity", then output listco_equity_jpm_*.csv, if the input directory ends with "CLO Equity", then output clo_equity_jpm_*.csv.



++++++++++
ver 0.18
++++++++++
1. The csv output filename now depends on the input directory, if the input directory ends with "ListCo Equity", then output listco_equity_*.csv, if the input directory ends with "CLO Equity", then output clo_equity_*.csv, else output jpm_*.csv.



++++++++++
ver 0.17
++++++++++
1. The csv output filename now depends on the input directory, if the input directory ends with "ListCo Equity", then output listco_equity_*.csv, if the input directory ends with "CLO Equity", then output clo_equity_*.csv, else output jpm_*.csv.



++++++++++
ver 0.16
++++++++++
1. Updated the portfolio code from '12306' to '12404' in function map_portfolio_id()



++++++++++
ver 0.15
++++++++++
Tested with JPM reon.

1. Change the output csv to use '|' as delimiter, to avoid potential problem due to data field such as "security name" containing commas.

2. Add date to the output csv file name.



++++++++++
ver 0.14
++++++++++
1. Move the id_lookup.py to another project, so that it becomes a centralized place for lookup from multiple projects, like jpm, bochk, etc. Easier to maintain.



++++++++++
ver 0.13
++++++++++
1. Add one more column "currency" to the output csv file, so that we can use both isin code and local currency to match, to minimize same isin code mapping to multiple instruments. 



++++++++++
ver 0.12
++++++++++
1. Add one more column "geneva_investment_id" to the output csv file, just for HTM portfolios.

2. Change the id lookup function to make it similar to project bochk's, and move them into a separate module.



++++++++++
ver 0.11
++++++++++
1. Add two entries in the config file:

	> base directory for input jpm position file and output the csv files. So those files do not mix with the code.

	> base directory for the log file. So during production deployment, the log file can be put in a different directory for easy checking.

2. logging function is handled by another package config_logging.

3. Bug fix: in utility.py, the get_datemode() function raises no exception when datemode value is invalid.



++++++++++
ver 0.1
++++++++++

date: 2016-11-24

1. Map the JPM account code to Geneva account code, now available for China Life overseas discretionary equity (11490) and non-discretionary equity (11491 etc.) portfolios, as well as for China Life ListCo discretionary equity (12307) and non-discretionary equity (12306 etc.) portfolios. To add more, see the map_portfolio_id() function in open_jpm.py module.

2. Map the JPM security id to Geneva investment id for non-public investments. Currently there is only one, the "CHINA LIFE OF CO - INVESTMENT VISTA" private equity, with JPM security id "4C0198S". To add more, update the "investmentLookup.xls" file in the current directory.

3. Now all investments are marked with isin number only, but the isin number is not a unique identifier for some equities listed in multiple exchange. E.g., for HSBC, 5 HK equity (listed in HK), and HSBC LN equity (listed in London), share the same isin code GB0005405286. To make it safer, we can consider using Bloomberg ticker (5 HK Equity) to mark the equity positions. Consider build a lookup table with JPM internal security id to ticker, JPM has a different position file format that contains the ticker, can use that to help.