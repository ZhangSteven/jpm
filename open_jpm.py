# coding=utf-8
# 
# Read the holdings section of the excel file from trustee.
#
# 

from xlrd import open_workbook
from xlrd.xldate import xldate_as_datetime
import xlrd
import datetime
from jpm.utility import logger, get_datemode, retrieve_or_create



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

	return account_code, account_name



def read_account(ws, row, port_values):
	"""
	Read the information of an account into the holding object port_values
	"""
	rows_read = 0

	while (row+rows_read < ws.nrows):

		cell_value = ws.cell_value(row+rows_read, 0)

		# detect start of an account
		if isinstance(cell_value, str) and cell_value.startswith('Account:'):
			account_code, account_name = extract_account_info(cell_value)
			account = {}
			port_values[len(port_values)+1] = account
			account['account_code'] = account_code
			account['account_name'] = account_name

			# move to next row
			rows_read = rows_read + 1
			cell_value = ws.cell_value(row+rows_read, 0)

			# is the following section a holdings section (0 or 1)
			if isinstance(cell_value, str) and cell_value == 'Security ID':
				holdings = []
				account['holdings'] = holdings
				n = read_holdings(ws, row+rows_read, holdings)
				rows_read = rows_read + n

			# is the following section a cash section (always, either after
			# the holding section or directly after the account information)
			if isinstance(cell_value, str) and cell_value == 'Branch Code':
				cash = []
				account['cash'] = cash
				n = read_cash(ws, row+rows_read, cash)
				rows_read = rows_read + n
				break	# finish reading this account

			elif isinstance(cell_value, str) and cell_value == 'No Data for this Account':
				rows_read = rows_read + 1
				break	# finish reading this account, no information


		rows_read = rows_read + 1
		# end of while loop

	return rows_read



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
				if cell_value == '':
					logger.error('read_holding_position(): field {0} is empty'.
								format(fld))
					raise ValueError('field {0} is empty'.format(fld))
				else:
					position[fld] = cell_value
			else:
				logger.error('read_holding_position(): invalid type for field {0}, value={1}'.
								format(fld, cell_value))
				raise TypeError('invalid data type for field {0}'.format(fld))
		
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

	The function returns the number of rows read, the the holdings_total
	holding object. This holding object is then used to verify holding
	positions are read properly.
	"""
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
            	raise ValueError('inconsisten sub total for field {0}'.
            						format(field))



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