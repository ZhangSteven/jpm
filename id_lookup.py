# coding=utf-8
# 
from xlrd import open_workbook
from jpm.utility import logger, get_current_path



class InvestmentCurrencyNotFound(Exception):
	pass

class InvalidPortfolioId(Exception):
	pass

class InvestmentIdNotFound(Exception):
	pass



def get_investment_Ids(portfolio_id, security_id_type, security_id):
	"""
	Determine the appropriate investment id for a security, based on:

	1. The portfolio's accounting treatment
	2. The security's id type: ISIN for isin code
	3. The security's id value

	Returns a tuple (geneva_investment_id_for_HTM, isin, bloomberg_figi)
	"""
	accounting_treatment = get_portfolio_accounting_treatment(portfolio_id)

	if security_id_type == 'ISIN':
		return get_investment_id_from_isin(accounting_treatment, security_id)
	
	else:
		isin, bbg_id, geneva_investment_id_for_HTM = lookup_investment_id(security_id_type, security_id)
		if isin != '':
			return get_investment_id_from_isin(accounting_treatment, isin)
		else:
			if accounting_treatment == 'HTM':
				return (geneva_investment_id_for_HTM, '', '')
			else:
				return ('', '', bbg_id)



def get_investment_id_from_isin(accounting_treatment, isin):
	if accounting_treatment == 'HTM':
		return (isin + ' HTM', '', '')
	else:
		return ('', isin, '')



investment_lookup = {}
currency_lookup = {}
def initialize_investment_lookup(lookup_file=get_current_path()+'\\investmentLookup.xls'):
	"""
	Initialize the lookup table from a file, for those securities that
	do have an isin code.

	To lookup,

	isin, bbg_id = investment_lookup(security_id_type, security_id)
	"""
	logger.debug('initialize_investment_lookup(): on file {0}'.format(lookup_file))

	wb = open_workbook(filename=lookup_file)
	ws = wb.sheet_by_name('Sheet1')
	row = 1
	global investment_lookup
	while (row < ws.nrows):
		security_id_type = ws.cell_value(row, 0)
		if security_id_type.strip() == '':
			break

		security_id = ws.cell_value(row, 1)
		isin = ws.cell_value(row, 3)
		bbg_id = ws.cell_value(row, 4)
		investment_id = ws.cell_value(row, 5)
		if isinstance(security_id, float):
			security_id = str(int(security_id))

		investment_lookup[(security_id_type.strip(), security_id.strip())] = \
			(isin.strip(), bbg_id.strip(), investment_id.strip())

		row = row + 1
	# end of while loop 

	ws = wb.sheet_by_name('Sheet2')
	row = 1
	global currency_lookup
	while (row < ws.nrows):
		security_id_type = ws.cell_value(row, 0)
		if security_id_type.strip() == '':
			break

		security_id = ws.cell_value(row, 1)
		currency = ws.cell_value(row, 3)
		if isinstance(security_id, float):
			security_id = str(int(security_id))

		currency_lookup[(security_id_type.strip(), security_id.strip())] = currency.strip()

		row = row + 1
	# end of while loop 



def lookup_investment_id(security_id_type, security_id):
	global investment_lookup
	if len(investment_lookup) == 0:
		initialize_investment_lookup()

	try:
		return investment_lookup[(security_id_type, security_id)]
	except KeyError:
		logger.error('lookup_investment_id(): No record found for security_id_type={0}, security_id={1}'.
						format(security_id_type, security_id))
		raise InvestmentIdNotFound()



def lookup_investment_currency(security_id_type, security_id):
	global currency_lookup
	if len(currency_lookup) == 0:
		initialize_investment_lookup()

	try:
		return currency_lookup[(security_id_type, security_id)]
	except KeyError:
		logger.error('lookup_investment_currency(): No record found for security_id_type={0}, security_id={1}'.
						format(security_id_type, security_id))
		raise InvestmentCurrencyNotFound()



def get_portfolio_accounting_treatment(portfolio_id):
	"""
	Map a portfolio id to its accounting treatment.
	"""
	a_map = {
		# China Life overseas equity, discretionary / non-discretionary
		'11490':'Trading',
		'11491':'Trading',
		'11492':'Trading',
		'11493':'Trading',
		'11494':'Trading',
		'11495':'Trading',

		# China Life ListCo equity, discretionary / non-discretionary
		'12306':'Trading',
		'12307':'Trading',
		'12308':'Trading',
		'12309':'Trading',

		# China Life overseas bond
		'12548':'HTM'
	}
	try:
		return a_map[portfolio_id]
	except KeyError:
		logger.error('get_portfolio_accounting_treatment(): {0} is not a valid portfolio id'.
						format(portfolio_id))
		raise InvalidPortfolioId()