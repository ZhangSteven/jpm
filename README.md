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

1. Since trade_converter established a lookup from isin to ticker, consider to make the lookup an independent module, then we can do:

	a. tell whether the holding is an equity based on: has an isin, but no coupon date and maturity date.

	b. lookup geneva investment id from there.


2. Fix the raise no exception in get_datemode()

3. Like trade_converter, use another directory as the working directory, and change to use argparse instead of just command line.


++++++++++
ver 0.1
++++++++++

date: 2016-11-24

1. Map the JPM account code to Geneva account code, now available for China Life overseas discretionary equity (11490) and non-discretionary equity (11491 etc.) portfolios, as well as for China Life ListCo discretionary equity (12307) and non-discretionary equity (12306 etc.) portfolios. To add more, see the map_portfolio_id() function in open_jpm.py module.

2. Map the JPM security id to Geneva investment id for non-public investments. Currently there is only one, the "CHINA LIFE OF CO - INVESTMENT VISTA" private equity, with JPM security id "4C0198S". To add more, update the "investmentLookup.xls" file in the current directory.

3. Now all investments are marked with isin number only, but the isin number is not a unique identifier for some equities listed in multiple exchange. E.g., for HSBC, 5 HK equity (listed in HK), and HSBC LN equity (listed in London), share the same isin code GB0005405286. To make it safer, we can consider using Bloomberg ticker (5 HK Equity) to mark the equity positions. Consider build a lookup table with JPM internal security id to ticker, JPM has a different position file format that contains the ticker, can use that to help.