# coding=utf-8
# 
# Read the holdings section of the excel file from trustee.
#
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
import datetime, csv, os
from jpm.utility import get_datemode, retrieve_or_create, \
						get_current_path, logger, get_input_directory
from investment_lookup.id_lookup import get_investment_Ids, \
										lookup_investment_currency



class NoCurrencyCodeInName(Exception):
	pass

class InconsistentSubtotal(Exception):
	pass

class InvalidAccountCode(Exception):
	pass



def read_jpm(ws, port_values):
	"""
	Read the worksheet with portfolio holdings. To retrieve holding, 
	we do:

	equity_holding = port_values['equity']
	for equity in equity_holding:
		equity['ticker'], equity['name']
		... retrive equity values using the following key ...

		ticker, isin, accounting_treatment, name, number_of_shares, currency, 
		listed_location, fx_on_trade_day, last_trade_date, average_cost, price, 
		book_cost, market_value, market_gain_loss, fx_gain_loss

	bond_holding = port_values['bond']
	for bond in bond_holding:
		bond['isin'], bond['name']
		... retrive bond values using the following key ...

		isin, name, accounting_treatment, par_amount, currency, is_listed, 
		listed_location, fx_on_trade_day, coupon_rate, coupon_start_date, 
		maturity_date, average_cost, amortized_cost, price, book_cost,
		interest_bought, amortized_value, market_value, accrued_interest,
		amortized_gain_loss, market_gain_loss, fx_gain_loss

	Note a bond may not have all of the above fields, depending on
	its accounting treatment. A HTM bond has amortized_cost, amortized_value,
	amortized_gain_loss, while a trading bond has price, market_value, 
	market_gain_loss set to zero.

	"""
	logger.debug('in read_jpm()')
	

	"""
	Now trying to read cash and holdings in the worksheet. The structure of
	data is as follows:

	Account1:
		holding section (there is 0 or 1 holding section)
			holding data fields
			holding1
			holding2
			...

		cash section (there is 1 cash section)
			cash account data fields
			cash account1
			cash account2
			...

	Account2:
		...


	Special case: if under an account, it specifies 'No Data for this Account',
	then ignore this account.

	"""

	row, d = read_date(ws, 0)
	port_values['date'] = d

	while (row < ws.nrows):
		rows_read = read_account(ws, row, port_values)
		row = row + rows_read

	logger.debug('out of read_jpm()')



def read_date(ws, row):
	"""
	Read the date in the broker statement
	"""
	rows_read = 0

	while (row+rows_read < ws.nrows):
		cell_value = ws.cell_value(row+rows_read, 0)
		if isinstance(cell_value, str) and cell_value.startswith('As Of:'):
			temp_list = cell_value.split(':')
			if len(temp_list) != 2:
				logger.error('read_date(): invalid date format: {0}'.format(cell_value))
				raise ValueError('invalid date format')

			date_string = str.strip(temp_list[1])
			temp_list = date_string.split('-')

			if len(temp_list) != 3:	# expect a string like '06-Jan-2015'
				logger.error('read_date(): invalid date string: {0}'.format(date_string))
				raise ValueError('invalid date_string')

			try:
				day = int(temp_list[0])
				year = int(temp_list[2])
				month = \
				{'jan':1, 'feb':2, 'mar':3, 'apr':4, 'may':5, 'jun':6,
				'jul':7, 'aug':8, 'sep':9, 'oct':10, 'nov':11, 'dec':12}[temp_list[1].lower()]

				d = datetime.datetime(year, month, day)
				break	# stop reading more rows

			except:
				logger.error('read_date(): failed to convert date string: {0}'.
								format(date_string))
				logger.exception('read_date(): ')
				raise


		rows_read = rows_read + 1
		# end of while loop

	return rows_read, d



def extract_account_info(account_info):
	"""
	Extract the account code and name from the account info string.

	The string looks like: 

	'Account:   48029   CLT - CLI HK BR (CLASS A-HK) TRUST FUND  '

	"""
	temp_list = account_info.split(':')
	if len(temp_list) != 2:
		logger.error('extract_account_info(): invalid account info: {0}'.
						format(account_info))
		raise ValueError('invalid account information')

	info_string = str.strip(temp_list[1])
	temp_list = info_string.split()
	if len(temp_list) < 2:	
		logger.error('extract_account_info(): invalid account info string: {0}'.
						format(info_string))
		raise ValueError('invalid account info string')

	account_code = temp_list[0]
	account_name = str.strip(info_string[len(account_code):])

	logger.debug('extract_account_info(): {0}, {1}'.format(account_code, account_name))
	return account_code, account_name



def account_begins(cell_value):
	if isinstance(cell_value, str) and cell_value.startswith('Account:'):
		return True



def read_account(ws, row, port_values):
	"""
	Read the information of an account into the holding object port_values
	"""
	rows_read = 0

	while (row+rows_read < ws.nrows):

		cell_value = ws.cell_value(row+rows_read, 0)
		if account_begins(cell_value):
			break

		rows_read = rows_read + 1
		# end of while loop

	if row+rows_read >= ws.nrows:	# reaches end of file
		return rows_read

	account_code, account_name = extract_account_info(cell_value)
	account = {}
	accounts = retrieve_or_create(port_values, 'accounts')
	accounts.append(account)
	account['account_code'] = account_code
	account['account_name'] = account_name
	rows_read = rows_read + 1
	cell_value = ws.cell_value(row+rows_read, 0)

	# if the following section is a holdings section (there may be
	# 0 or 1 holding section)
	if isinstance(cell_value, str) and cell_value == 'Security ID':
		holdings = []
		account['holdings'] = holdings
		n = read_holdings(ws, row+rows_read, holdings)
		rows_read = rows_read + n
		cell_value = ws.cell_value(row+rows_read, 0)

	# if the following section a cash section (there is always a cash
	# section, either following a holding section or directly following
	# the account information
	if isinstance(cell_value, str) and cell_value == 'Branch Code':
		cash = []
		account['cash'] = cash
		n = read_cash(ws, row+rows_read, cash)
		rows_read = rows_read + n

	elif isinstance(cell_value, str) and cell_value == 'No Data for this Account':
		rows_read = rows_read + 1

	elif account_begins(cell_value):	# the next account begins
		pass
	else:
		logger.error('read_account(): unexpected sub section in row {0}'.
						format(row+rows_read))

	return rows_read



def is_empty_account(account):
	"""
	Tell whether an account is an empty account, i.e., those with 'No Data for
	this Account' in the excel.
	"""
	if 'cash' in account or 'holdings' in account:
		return False
	else:
		return True



def read_holdings(ws, row, holdings):
	"""
	Read the holdings section. Each holdings section will consist of
	the following:

	holding fields subsection (1)

	holdings subsection (1..n)
		holding1
		holding2
		...

	holding total subsection(1)

	"""
	logger.debug('read_holdings(): at row {0}'.format(row))
	rows_read = 0

	rows_each_holding, coordinates, fields = read_holding_fields(ws, row+rows_read)
	rows_read = rows_read + rows_each_holding

	# read each holding position
	while (row+rows_read < ws.nrows):
		if is_holdings_subtotal(ws, row+rows_read):
			n, holdings_total = read_holdings_total(ws, row+rows_read)
			validate_holdings_total(holdings, holdings_total)
			rows_read = rows_read + n
			break

		while (is_blank_line(ws, row+rows_read)):
			rows_read = rows_read + 1

		# if it is not a blank line, not a holding sub total,
		# then it must be a holding position
		read_holding_position(ws, row+rows_read, coordinates, fields, holdings)
		rows_read = rows_read + rows_each_holding
		# end of while loop
		

	return rows_read



def read_holding_fields(ws, row):
	"""
	The holding fields subsection tells the reader which data field each
	cell contains. Because the data fields are arranged in a 2 dimensional
	way, we need to return the coordinates, the name of the fields, and
	how many rows each holding position contains.

	For example, if the return is

	rows_each_holding = 4
	coordinates = [(0,0), (0,1), (2,2)]
	fields = ['security_id', 'security_name', 'coupon_rate']

	Then it is telling the user that in the holding section subsection,
	in each holding position, relative to the position of 'security id',
	in the same row and second column, the data field is 'security name',
	in the third row and third column, the data field is 'coupon rate'.

	It also tells each holding position will take 4 rows in the excel
	spread sheet.

	"""
	logger.debug('read_holding_fields(): at row {0}'.format(row))
	rows_read = 0
	fields = []
	coordinates = []

	while (row+rows_read < ws.nrows):

		for column in range(9):
			if is_empty_cell(ws, row+rows_read, column):
				continue

			cell_value = ws.cell_value(row+rows_read, column)
			if not isinstance(cell_value, str):	# data field name needs to
												# be string
				logger.error('read_holding_fields(): invalid data field: {0}'.
								format(cell_value))
				raise ValueError('data field not a string')

			if cell_value == 'Security ID':
				fld = 'security_id'
			elif cell_value == 'Security Name':
				fld = 'security_name'
			elif cell_value == 'Location/Nominee':
				fld = 'location_or_nominee'
			elif cell_value == 'Awaiting Receipt':
				fld = 'awaiting_receipt'
			elif cell_value == 'Settled Units':
				fld = 'settled_units'
			elif cell_value == 'Total Units':
				fld = 'total_units'
			elif cell_value == 'ISIN':
				fld = 'isin'
			elif cell_value == 'Reg./Sub Acct.':
				fld = 'regional_or_sub_account'
			elif cell_value == 'Awaiting Delivery':
				fld = 'awaiting_delivery'
			elif cell_value == 'Current Face-Settled':
				fld = 'current_face_settled'
			elif cell_value == 'Current Face-Total':
				fld = 'current_face_total'
			elif cell_value == 'OCC ID':
				fld = 'occ_id'
			elif cell_value == 'Coupon Rate':
				fld = 'coupon_rate'
			elif cell_value == 'Maturity Date':
				fld = 'maturity_date'
			elif cell_value == 'Pool Number':
				fld = 'pool_number'
			elif cell_value == 'Country':
				fld = 'country'
			elif cell_value == 'Collateral Units':
				fld = 'collateral_units'
			elif cell_value == 'Borrowed Units':
				fld = 'borrowed_units'
			else:	# data field not handled
				logger.error('read_holding_fields(): unhandled data field: {0}'.
								format(cell_value))
				raise ValueError('data field not handled')

			fields.append(fld)
			if fld in ['coupon_rate', 'maturity_date', 'pool_number']:
				# in the actual holding position, the row offset for these
				# three fields are not the same as the holding fields.
				coordinates.append((rows_read-1, column))
			else:
				coordinates.append((rows_read, column))
			# end of for loop

		rows_read = rows_read + 1
		if is_blank_line(ws, row+rows_read):
			break
		# end of while loop

	return rows_read, coordinates, fields



def read_holding_position(ws, row, coordinates, fields, holdings):
	"""
	Read a holding position and save it into the holdings object.
	"""
	logger.debug('read_holding_position(): at row {0}'.format(row))
	position = {}

	i = 0
	for fld in fields:
		row_offset, col_offset = coordinates[i]
		i = i + 1
		cell_value = ws.cell_value(row+row_offset, col_offset)
		if isinstance(cell_value, str):
			cell_value = str.strip(cell_value)

		if fld in ['security_id', 'security_name', 'isin', 
					'regional_or_sub_account', 'location_or_nominee',
					'country']:	# mandatory fields whose value is string

			if isinstance(cell_value, str):
				# if cell_value == '':
				# 	# special case handling
				# 	if fld == 'isin':
				# 		position['isin'] = ''
				# 		position['geneva_investment_id'] = map_geneva_investment_id(position['security_id'])
				# 	else:
				# 		logger.error('read_holding_position(): field {0} is empty'.
				# 						format(fld))
				# 		raise ValueError('field {0} is empty'.format(fld))
				# else:
				# 	position[fld] = cell_value
				position[fld] = cell_value.strip()
			else:
				logger.error('read_holding_position(): invalid type for field {0}, value={1}'.
								format(fld, cell_value))
				raise TypeError
		
		elif fld in ['awaiting_receipt', 'settled_units', 'total_units',
						'awaiting_delivery', 'collateral_units', 
						'borrowed_units']:	# mandatory fields whose
											# value is float

			if isinstance(cell_value, float):
				position[fld] = cell_value
			else:
				logger.error('read_holding_position(): invalid type for field {0}, value={1}'.
								format(fld, cell_value))
				raise TypeError('invalid data type for field {0}'.format(fld))
		
		elif fld in ['occ_id', 'coupon_rate', 'maturity_date', 'pool_number',
						'current_face_settled', 'current_face_total']:
			
			# optional fields
			if isinstance(cell_value, str) and cell_value == '':
				pass	# if they are not there, skip it.

			elif fld in ['coupon_rate', 'maturity_date', 
							'current_face_settled', 'current_face_total']:

				if isinstance(cell_value, float):
					if fld == 'maturity_date':
						datemode = get_datemode()
						position[fld] = xldate_as_datetime(cell_value, datemode)
					elif fld == 'coupon_rate':
						position[fld] = cell_value/100
					else:
						position[fld] = cell_value
				else:
					logger.error('read_holding_position(): invalid type for field {0}, value={1}'.
								format(fld, cell_value))
					raise TypeError('invalid data type for field {0}'.format(fld))

			elif fld in ['occ_id', 'pool_number']:

				if isinstance(cell_value, str):
					position[fld] = cell_value
				else:
					logger.error('read_holding_position(): invalid type for field {0}, value={1}'.
								format(fld, cell_value))
					raise TypeError('invalid data type for field {0}'.format(fld))

		else:	# unhandled field names
			logger.error('read_holding_position(): unhandled field {0}'.format(fld))
			raise TypeError('invalid field name: {0}'.format(fld))

	# end of for loop

	holdings.append(position)



def read_holdings_total(ws, row):
	"""
	Read the sub total of all holdings in an account

	The function returns the number of rows read, then the holdings_total
	holding object. This holding object is then used to verify holding
	positions are read properly.
	"""
	logger.debug('read_holdings_total(): at row {0}'.format(row))
	holdings_total = {}
	fields = ['awaiting_receipt', 'settled_units', 'total_units',
	'awaiting_delivery', 'current_face_settled', 'current_face_total']

	i = 0
	for r in range(row, row+2):
		for column in range(5, 8):
			cell_value = ws.cell_value(r, column)
			if isinstance(cell_value, str) and str.strip(cell_value) == '':
				cell_value = 0

			try:
				holdings_total[fields[i]] = float(cell_value)
			except ValueError:	# the input could be a string in the form
								# of 1,234.88, remove the ','
				cell_value = cell_value.replace(',', '')
				holdings_total[fields[i]] = float(cell_value)

			i = i + 1

	# end of for loop
	return 2, holdings_total



def read_cash(ws, row, cash):
	"""
	Read the cash positions for each account
	"""
	logger.debug('read_cash(): at row {0}'.format(row))
	rows_read = 0

	fields = read_cash_fields(ws, row)
	rows_read = rows_read + 1

	# read each holding position
	while (row+rows_read < ws.nrows):

		while (is_blank_line(ws, row+rows_read)):
			rows_read = rows_read + 1
		
		cell_value = ws.cell_value(row+rows_read, 0)
		if account_begins(cell_value):
			break

		# try:
		# 	n = read_cash_position(ws, row+rows_read, fields, cash)
		# 	rows_read = rows_read + n
		# except (ValueError,TypeError):	# this line does not look like a valid
		# 								# cash position, stop reading
		# 	logger.info('read_cash(): row {0} is not a cash position'.
		# 					format(row+rows_read))
		# 	break

		n = read_cash_position(ws, row+rows_read, fields, cash)
		rows_read = rows_read + n

		# end of while loop

	return rows_read



def read_cash_fields(ws, row):
	"""
	Read the data fields in a cash position
	"""
	logger.debug('read_cash_fields(): at row {0}'.format(row))
	fields = []

	for column in range(10):	# read up to column J
		if is_empty_cell(ws, row, column):
			fld = 'empty_field'
			fields.append(fld)
			continue

		cell_value = ws.cell_value(row, column)
		if not isinstance(cell_value, str):	# data field name needs to
											# be string
			logger.error('read_cash_fields(): invalid cash field: {0}'.
							format(cell_value))
			raise ValueError('cash field not a string')

		if cell_value == 'Branch Code':
			fld = 'branch_code'
		elif cell_value == 'Branch Name':
			fld = 'branch_name'
		elif cell_value == 'Cash Account':
			fld = 'account_number'
		elif cell_value == 'Cash Account Name':
			fld = 'account_name'
		elif cell_value == 'Local CCY':
			fld = 'currency'
		elif cell_value == 'DGSD Eligible':
			fld = 'dgsd_eligible'
		elif cell_value == 'Opening Cash Balance':
			fld = 'opening_balance'
		elif cell_value == 'Closing Cash Balance':
			fld = 'closing_balance'
		else:	# data field not handled
			logger.error('read_cash_fields(): unhandled cash field: {0}'.
							format(cell_value))
			raise ValueError('cash field not handled')

		fields.append(fld)


	return fields



def read_cash_position(ws, row, fields, cash):
	"""
	Read a cash position
	"""
	logger.debug('read_cash_position(): at row {0}'.format(row))
	position = {}
	column = -1
	for field in fields:
		column = column + 1
		if field == 'empty_field':	# ignore
			continue

		cell_value = ws.cell_value(row, column)
		# logger.debug(cell_value)
		if field in ['branch_code', 'branch_name', 'account_number', 
						'account_name', 'currency', 'dgsd_eligible']:

			if not isinstance(field, str):
				logger.error('read_cash_position(): field {0} not a string, value = {1}'.
								format(field, cell_value))
				raise TypeError('invalid type for field {0}'.format(field))

			elif str.strip(cell_value) == '':
				logger.error('read_cash_position(): field {0} is empty'.
								format(field))
				raise ValueError('field {0} is empty'.format(field))

			position[field] = str.strip(cell_value)

		elif field in ['opening_balance', 'closing_balance']:
			try:
				position[field] = float(cell_value)
			except ValueError:
				logger.error('read_cash_position(): field {0} cannot be converted to float, value = {1}'.
								format(field, cell_value))
				raise TypeError('failed to read field {0} as float number'.format(field))

		else:	# unexpected field
			logger.error('read_cash_position(): unexpected field {0}'.
								format(field))
			raise ValueError('unexpected field {0}'.format(field))

	# end of for loop

	cash.append(position)
	return 1	# read 1 row



def validate_holdings_total(holdings, holdings_total):
        """
        Add up the six fields in each position:

        'awaiting_receipt', 'settled_units', 'total_units',
        'awaiting_delivery', 'current_face_settled', 'current_face_total'

        Then compare it to the sub total, make sure they are equal.
        """
        fields = ['awaiting_receipt', 'settled_units', 'total_units',
        'awaiting_delivery', 'current_face_settled', 'current_face_total']

        for field in fields:
            sub_total = calculate_sub_total(field, holdings)
            if abs(sub_total - holdings_total[field]) > 0.000001:
            	logger.error('validate_holdings_total(): sub total does not match for field {0}: {1} != {2}'.
            					format(field, sub_total, holdings_total[field]))
            	raise InconsistentSubtotal



def calculate_sub_total(field, holdings):
	"""
	Go through each position in the holdings, add up the number in
	'field'.
	"""
	sub_total = 0
	for position in holdings:
		try:
			n = position[field]
		except KeyError:
			n = 0

		sub_total = sub_total + n

	return sub_total



def is_holdings_subtotal(ws, row):
	"""
	Tell whether this is a holdings subtotal line, this line has the 
	first 4 cells empty, the fifth cell contains 'Totals: '
	"""
	for column in range(4):
		if not is_empty_cell(ws, row, column):
			return False

	cell_value = ws.cell_value(row, 4)
	if isinstance(cell_value, str) and cell_value.startswith('Totals:'):
		return True
	else:
		return False



def is_blank_line(ws, row):
	"""
	Tell whether it is a blank line.
	
	If the first 6 cells in this row are all empty, then it is a blank line.
	"""
	for column in range(6):
		if not is_empty_cell(ws, row, column):
			return False

	return True



def is_empty_cell(ws, row, column):
	"""
	If the cell value is all white space or an empty string, then it is
	an empty cell.
	"""
	cell_value = ws.cell_value(row, column)
	if isinstance(cell_value, str) and str.strip(cell_value) == '':
		return True
	else:
		return False



def get_portfolio_date_as_string(port_values):
	d = port_values['date']
	return convert_datetime_to_string(d)



def convert_datetime_to_string(dt):
	"""
	convert a datetime object to string in the 'yyyy-mm-dd' format.
	"""
	return '{0}-{1}-{2}'.format(dt.year, dt.month, dt.day)



def map_portfolio_id(account_code):
	"""
	Map the account code of JP Morgan to the portfolio id in Geneva.
	"""

	# China Life overseas accounts
	if account_code == '48029':
		return '11490'
	elif account_code == '48089':
		return '11491'
	elif account_code == '48090':
		return '11492'
	elif account_code == '48195':
		return '11493'
	elif account_code == '53412':
		return '11494'
	elif account_code == '53413':
		return '11495'

	# China Life ListCo accounts
	elif account_code == 'AFU34':
		return '12404'
	elif account_code == 'AFU35':
		return '12307'
	elif account_code == 'BBK32':
		return '12308'
	elif account_code == 'AFU37':
		return '12309'

	else:
		logger.error('map_portfolio_id(): invalid account code {0}'.
						format(account_code))
		raise InvalidAccountCode()



def get_currency_from_name(security_name):
	"""
	Extract the currency from the security name
	"""
	tokens = security_name.split()
	currency_code = ['HKD', 'USD', 'CNY', 'SGD', 'JPY', 'EUR']
	if tokens[-2] in currency_code:
		return tokens[-2]
	elif tokens[-1] in currency_code:
		return tokens[-1]
	else:
		raise NoCurrencyCodeInName()



def get_prefix_from_dir(input_dir):
	"""
	Work out a prefix for the filename depending on the input directory.
	"""
	token = input_dir.split('\\')[-1]
	if token.lower() == 'listco equity':
		return 'listco_equity'
	elif token.lower() == 'clo equity':
		return 'clo_equity'
	else:
		return 'jpm'



def create_csv_file_name(date_string, file_suffix):
	"""
	Create the output csv file name based on the date string, as well as
	the file suffix: cash, afs_positions, or htm_positions
	"""
	input_dir = get_input_directory()
	csv_file = input_dir + '\\' + get_prefix_from_dir(input_dir) + '_' + \
				date_string + '_' + file_suffix + '.csv'
	return csv_file



def write_csv(port_values):
	"""
	Write cash and holdings into csv files.
	"""	
	# cash_file = get_input_directory() + '\\cash.csv'
	write_cash_csv(port_values)

	# holding_file = get_input_directory() + '\\holding.csv'
	write_holding_csv(port_values)



def write_cash_csv(port_values):
	portfolio_date = get_portfolio_date_as_string(port_values)
	cash_file = create_csv_file_name(portfolio_date, 'cash')

	with open(cash_file, 'w', newline='') as csvfile:
		logger.debug('write_cash_csv(): {0}'.format(cash_file))
		file_writer = csv.writer(csvfile, delimiter='|')

		fields = ['currency', 'opening_balance', 'closing_balance']
		file_writer.writerow(['portfolio', 'date', 'custodian'] + fields)
		
		accounts = port_values['accounts']
		for account in accounts:
			if is_empty_account(account) or not 'cash' in account:
				continue

			portfolio_id = map_portfolio_id(account['account_code'])
			cash = account['cash']
			for position in cash:
				row = [portfolio_id, portfolio_date, 'JPM']

				for fld in fields:
					row.append(position[fld])

				file_writer.writerow(row)



def	write_holding_csv(port_values):
	portfolio_date = get_portfolio_date_as_string(port_values)
	holding_file = create_csv_file_name(portfolio_date, 'position')

	with open(holding_file, 'w', newline='') as csvfile:
		logger.debug('write_holding_csv(): {0}'.format(holding_file))
		file_writer = csv.writer(csvfile, delimiter='|')

		fields = ['security_name', 'country', 'awaiting_receipt', 
					'awaiting_delivery', 'collateral_units', 'borrowed_units', 
					'settled_units', 'total_units', 'coupon_rate', 'maturity_date']

		file_writer.writerow(['portfolio', 'date', 'geneva_investment_id', 'isin',
								'bloomberg_figi', 'currency'] + fields)

		accounts = port_values['accounts']
		for account in accounts:
			if is_empty_account(account) or not 'holdings' in account:
				continue

			portfolio_id = map_portfolio_id(account['account_code'])				
			holdings = account['holdings']

			for position in holdings:
				row = [portfolio_id, portfolio_date]
				if position['isin'] == '':
					security_id_type = 'JPM'
					security_id = position['security_id']
				else:
					security_id_type = 'ISIN'
					security_id = position['isin']

				investment_ids = get_investment_Ids(portfolio_id, security_id_type, security_id)
				for id in investment_ids:
					row.append(id)

				try:
					row.append(get_currency_from_name(position['security_name']))
				except NoCurrencyCodeInName:
					row.append(lookup_investment_currency('JPM', position['security_id']))

				for fld in fields:
					try:
						item = position[fld]
						if fld == 'maturity_date':
							item = convert_datetime_to_string(item)
					except KeyError:
						item = ''

					row.append(item)

				file_writer.writerow(row)




if __name__ == '__main__':
	import sys
	if len(sys.argv) < 2:
		print('use python open_jpm.py <input_file>')
		sys.exit(1)

	filename = get_input_directory() + '\\' + sys.argv[1]
	if not os.path.exists(filename):
		print('{0} does not exist'.format(filename))
		sys.exit(1)

	port_values = {}
	try:
		wb = open_workbook(filename=filename)
		ws = wb.sheet_by_name('Sheet1')
		read_jpm(ws, port_values)
		write_csv(port_values)
	except:
		logger.exception('open_jpm:main()')
		print('something goes wrong, check log file.')
	else:
		print('OK')