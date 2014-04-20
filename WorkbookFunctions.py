#####**************************************************************************#####
#####								DESCRIPTION	    						   #####
#####**************************************************************************#####


#####**************************************************************************#####
#####									CLASSES   						       #####
#####**************************************************************************#####
import datetime, itertools, os

class _InputError(Exception):
	def __init__(self, value):
		self.value = value
	def __str__(self):
		return repr(self.value)
		
class _NotFoundError(Exception):
	def __init__(self, value):
		self.value = value
	def __str__(self):
		return repr(self.value)
				
class Columns:
	def __init__(self, column_values):
		"""
		column_values : list
		
		Class for getting column values at specific points on a worksheet, and
		checking that column values at specific points on multiple worksheets are
		of equal value.\n
		
		Initialise class by passing a list of column values. Raises error if the
		column_values is not a list of integers\n
		
		Available Methods: \n
		get_values: get columns values at specific point on active sheet.
		compare_all_columns: compare column values at specific points on all sheets.
		"""
		if not all([isinstance(elem, int) for elem in column_values]):
			raise _InputError("List may only contain integers")
		self.column_values = column_values
		
	def get_values(self, row):
		"""
		row	: int
		return	: list
		method	: visible
		
		Returns list of lowered stripped string values found in each cell referenced
		by row and each column integer in column_values on active sheet.
		"""
		values_list = [str(Cell(row, col).value).lower().strip() 
					   for col in self.column_values]
		return values_list
		
	def __compare_values(self, master_list, sheet_list):
		"""
		master_list : list
		sheet_list	: list
		return 		: list or None
		method		: hidden
		
		Returns list of values where the master_list and sheet_list differ at every 
		index point.
		"""
		sheet_disparities = [i for i, x in zip(master_list, sheet_list) if i!=x]
		if not sheet_disparities:
			return None
		else:
			return sheet_disparities
	
	def __update_disparity_dict(self, sheet_disparities, disparity_dict):
		"""
		sheet_disparities : list
		disparity_dict	  : dict
		return			  : None
		method			  : hidden
		
		Updates disparity_dict if sheet_disparities are not none
		"""
		if sheet_disparities:
			disparity_dict[active_sheet()].extend(sheet_disparities)
			return
		return
		
	def compare_all_columns(self, start_row_dict):
		"""
		start_row_dict	: dict
		return		: dict
		method		: visible
		
		Returns dict that has a key for each sheet name in the workbook and values 
		that are the column header values that differ from a master list of values,
		that is drawn from the first sheet in the workbook.\n
		All values are created by calling get_values(), and the arguments to that 
		function are supplied by getting the start row from the start_row_dict and 
		the self.column_values object.
		"""
		sheets = all_sheets()
		if not all([isinstance(value, int) 
				    for key, value in start_row_dict.iteritems()]):
			raise _InputError("All values in dictionary must be integers")
			
		disparity_dict = {sheet : [] for sheet in sheets} 
		active_sheet(sheets[0])
		master_list = self.get_values(start_row_dict[sheets[0]])
		for sheet in sheets[1:]:
			active_sheet(sheet)
			sheet_list = self.get_values(start_row_dict[sheet])
			sheet_disparities = self.__compare_values(master_list, sheet_list)
			self.__update_disparity_dict(sheet_disparities, disparity_dict)
		return disparity_dict

class Dates:
	def __init__(self, date_cell_ref, strp_format = None, separator = None, index_pos = None):
		"""
		date_cell_ref	: tuple of integers
		strp_format	: string
		separator	: None or string
		index_pos	: None or int

		Class for getting dates from worksheets, converting strings to datetime
		objects, and checking date patterns in multiple worksheets.\n

		Initialise class by passing date cell reference. If dates on worksheets
		are already thought to be directly readable by DataNitro as datetime objects
		then no further arguments need be passed. If some string formatting is
		needed before passing string value to datetime.strptime() function used in
		converting strings to dates, then additional arguments needed. \n

		Available Methods \n
		get_value	: get value of date cell on active sheet
		cell_to_date	: get datetime object string in date cell on active sheet
		check_all_dates	: check dates on all sheets convertible to datetime objects
		find_duplicates	: find duplicate dates on sheets in workbook
		relative_order	: check date implied order equals order of sheets
		discontinuities	: identify when dates imply large discontinuities
		"""
		if not isinstance(date_cell_ref, tuple) or\
		not all(isinstance(elem, int) for elem in date_cell_ref) or\
		not len(date_cell_ref) == 2:
			raise _InputError("date_cell_ref must be a tuple that contains two integers")
		if strp_format: 
			if not isinstance(strp_format, str):
				raise _InputError("Argument 'strp_format' must be an string")
		if separator:
			if not isinstance(separator, str):
				raise _InputError("Argument 'separator' must be an string or None")
		if index_pos:
			if not isinstance(index_pos, int):
				raise _InputError("Argument 'index_pos' must be an integer or None")

				
		self.date_cell_ref = date_cell_ref
		self.strp_format = strp_format
		self.separator = separator
		self.index_pos = index_pos
		
	def __get_value(self):
		"""
		return	: string
		method	: visible
		
		Returns string value at cell referenced by self.date_cell_ref on active 
		sheet.
		"""
		if not self.strp_format:
			return Cell(self.date_cell_ref).value
		else:
			return str(Cell(self.date_cell_ref).value)

	def cell_to_date(self):
		"""
		return	: datetime.datetime or None
		method	: visible

		Returns datetime object of value of self.get_value() formatted according to
		strp_format on active sheet. If separator is specified self.get_value() is
		split and the value at index position index_pos is formatted according to
		strp_format and returned.\n
		Returns None if format to datetime object not possible.
		"""

		if not self.strp_format:
			date_object = self.__get_value()
			if isinstance(date_object, datetime.datetime):
				return date_object.date()
			else:
				return None
		if not self.separator:
			try:
				date_object = datetime.datetime.strptime(self.__get_value(), self.strp_format)
				return date_object.date()
			except ValueError:
				return None
		else:
			try:
				stripped_string = self.__get_value().split(self.separator)[self.index_pos]
				date_object = datetime.datetime.strptime(stripped_string, self.strp_format)
				return date_object.date()
			except ValueError:
				return None

	def __update_date_dict(self, date_object, date_dict):
		"""
		return 	: None
		method	: hidden

		Updates date_dict if date_object is not None
		"""
		if date_object:
			date_dict[active_sheet()] = date_object
			return
		else:
			date_dict[active_sheet()] = 'Date not found on this sheet'
			return

	def check_all_dates(self):
		"""
		return	: dict
		method	: visible

		Returns dict of sheet keys and values that are datetime objects created by
		calling	cell_to_date() with arguments passed in intialisation call. If no
		datetime object is created a string message to user is the key value.
		"""
		sheets = all_sheets()
		date_dict = {sheet : object() for sheet in sheets}
		for sheet in sheets:
			active_sheet(sheet)
			date_object = self.cell_to_date()
			self.__update_date_dict(date_object, date_dict)
		return date_dict

	def find_duplicates(self, date_dict = None):
		"""
		date_dict : dict or None
		return	: dict
		method	: visible

		Returns a dictionary where dates that are found more than once in the values
		of date_dict are the keys, and the values are the keys of date_dict at which
		the duplicate dates are found. Only accepts a date_dict whose values are all
		datetime.dateimte objects, otherwise an exception is raised.\n
		If no date_dict is passed one is created by calling check_all_dates().

		"""
		if not date_dict:
			date_dict = self.check_all_dates()
		date_list = [value for key, value in date_dict.iteritems()]
		if not all(isinstance(date, datetime.date) for date in date_list):
			raise _InputError("""
							  All date_cell_ref values in all sheets must be capable
							  of being datetime objects before running function.
							  Use the check_all_dates() method to perform checks.
							  """)

		duplicate_dates = set([date for date in date_list
							   if date_list.count(date) > 1])
		duplicates_dict = {duplicate : [key for key, value in date_dict.iteritems()
						   if value == duplicate] for duplicate in duplicate_dates}
		return duplicates_dict

	def relative_order(self, date_dict = None):
		"""
		date_dict	: dict or None
		return		: dict
		method		: visible

		Returns a dictionary that shows order of sheets implied by dates in the
		date_dict and the actual order of the sheets, if different.\n
		If no date_dict is passed one is created by	calling check_all_dates().\n
		Resulting order may be perverse if there are duplicate dates in date_dict.
		"""

		sheets = all_sheets()
		if not date_dict:
			date_dict = self.check_all_dates()
		date_list = [value for key, value in date_dict.iteritems()]
		if not all(isinstance(date, datetime.date) for date in date_list):
			raise _InputError("""
							  All date_cell_ref values in all sheets must be capable
							  of being datetime objects before running function.
							  Use the check_all_dates() method to perform checks
							  """)
		implied_order = [key for key, value in
						 sorted(date_dict.iteritems(), key = lambda x: x[1])]
		if implied_order == sheets:
			return {}
		else:
			relative_order_dict = {
				'actual order' : [x for x, i in zip(sheets, implied_order) if x!=i],
				'implied_order': [i for x, i in zip(sheets, implied_order) if i!=x]
								  }
			return relative_order_dict

	def discontinuities(self, discontinuity_value, date_dict = None):
		"""
		discontinuity_value : int
		date_dict	: dict or None
		return		: list of tuples
		method		: visible

		Returns list of tuples where each tuple is a pair of contiguous sheets where
		the	dates found on those sheets indicate a timedelta greater than
		discontinuity_value.\n
		If no date_dict is passed one is created by calling check_all_dates().
		"""

		sheets = all_sheets()
		if not date_dict:
			date_dict = self.check_all_dates()
		date_list = [value for key, value in sorted(date_dict.iteritems())]
		if not all(isinstance(date, datetime.date) for date in date_list):
			raise _InputError("""All date_cell_ref values in all sheets must be capable
								of being datetime objects before running this function.
								Use the check_all_dates() method to perform checks""")
		date_shift = date_list[:]
		date_shift.insert(0, date_list[0])
		date_discontinuities = [x - i for x, i in zip(date_list, date_shift[:-1])]
		unusual_discontinuities = [index for index, value in
								   enumerate(date_discontinuities)
								   if value > datetime.timedelta(discontinuity_value)]
		return [(sheets[x-1], sheets[x]) for x in unusual_discontinuities]


class FindPoints:
	def __init__(self, col, start_row, end_value, adjustments = None):
		"""
		col		: int
		start_row	: int
		end_value	: string
		adjustments	: None or int

		Class for finding specific points in  worksheet data. Useful for identifying
		'headers' as well as end points. The point of interest is identified in a
		column (col), by searching for the end_value beginning in the start_row, and
		then making any necessary adjustments as per the adjustments argument.\n

		Initialise class by passing an integer for the column to be searched, an
		integer indicating the start_row, the end_value to be found, and any
		necessary adjusments (a positive or negative integer). \n

		Available Methods\n
		find_point	: get row value of point on active worksheet
		find_all_points	: get dict of row values of points on all worksheets
		"""
		if not isinstance(col, int):
			raise _InputError("Argument 'col' must be an integer")
		if not isinstance(start_row, int):
			raise _InputError("Argument 'start_row' must be an integer")
		if not isinstance(end_value, str):
			raise _InputError("Argument 'end_value' must be an string")
		if adjustments:
			if not isinstance(adjustments, int):
				raise _InputError("Argument 'adjustments' must be an integer or None")

		self.col = col
		self.end_value = end_value.strip().lower()
		self.start_row = start_row
		self.adjustments = adjustments

	def find_point(self):
		"""
		return	: int
		method	: visible

		Returns row of cell where self.end_value is equal to cell referenced by
		self.col and a row value determined by iteration. Return value is adjusted
		by an amount as specified by self.adjustments. Raises error if end_value not
		found. \n
		Assumes end_value in first 200 rows.
		"""
		row = self.start_row
		while row < 300:
			if str(Cell(row, self.col).value).strip().lower() == self.end_value:
				if self.adjustments:
					row += self.adjustments
					return row
				return row
			else:
				row +=1
		raise _NotFoundError("Start row not found")

	def find_all_points(self):
		"""
		method : visible

		Returns dict with one key for each sheet in workbook with point row as
		value. If no point row is found, the key maps to a string value notifying
		the user of the absence of the point row.
		"""
		sheets = all_sheets()
		found_dict = {sheet : object() for sheet in sheets}
		for sheet in sheets:
			active_sheet(sheet)
			try:
				start = self.find_point()
				found_dict[sheet] = start
			except _NotFoundError:
				found_dict[sheet] = 'Point Not Found'
		return found_dict

class sheet_compiler:

	def __init__(self, filepath):
		"""
		filepath : raw string

		Class for compiling a single worksheet from multiple excel workbooks in a
		given directory into one new workbook.

		Initialise the class by passing a string of the file directoty path that
		contains the workbooks from which sheets will be compiled.
		"""
		if not isinstance(filepath, str):
			raise _InputError("Filepath must be a raw string")
		self.filepath = filepath
		os.chdir(filepath)

	def get_file_list(self):
		"""
		return 	: list
		method	: visible

		Returns list of files found in directory self.filepath
		"""
		return os.listdir(self.filepath)

	def __get_sheet(self, filename, sub_string1, sub_string2 = None):
		"""
		filename 	: string
		sub_string1	: string
		sub_string2	: string
		return		: srting
		method		: hidden

		Returns string identifying sheet in workbook identified as containing
		sub_string1	and optionally sub_string2 if provided as argument. Raises
		exception if the arguments do not uniquely identify a single sheet in the
		workbook.
		"""
		active_wkbk(filename)
		sheets = all_sheets()
		if sub_string2:
			selected_sheets = [sheet for sheet in sheets if sub_string1.lower() in
							   sheet.lower() and sub_string2.lower() in sheet.lower()]
		else:
			selected_sheets = [sheet for sheet in sheets if sub_string1.lower()
						       in sheet.lower()]

		if len(selected_sheets) == 0 or len(selected_sheets) > 1:
			raise NotFoundError("Error")

		sheet_name = selected_sheets[0]
		return sheet_name

	def __relocate_sheet(self, filename, to_workbook, sheet_name):
		"""
		filename 	: string
		to_workbook	: string
		sheet_name	: string
		return		: None
		method		: hidden

		Moves sheet with sheet_name to to_workbook using DN copy_sheet function.
		"""
		active_wkbk(filename)
		copy_sheet(to_workbook, sheet_name)
		return

	def compile_sheets(self, filelist, new_wkbk_name, sub_string1, sub_string2 = None):
		"""
		filelist		: list
		new_wkbk_name	: string
		sub_string1		: string
		sub_string2		: string or None
		return			: formatted string

		Returns formatted string that gives report to user as to success of the
		compile operation.\n
		Function opens every file in the self.filepath directory that is contained
		in the filelist passed as argument. Up to two substrings may be passed as
		arguments, and the combination of these substrings should uniquely identify
		the worksheet to be moved (as there may be more than one).\n
		Sheets that are successfully identified for copying will be copied to a new
		workbook created according to new_wkbk_name. This workbook will be in the
		same directory as the files being iterated over. \n
		Files are opened in the reverse order they are found in the filelist.
		"""

		unsuccessful = []
		new_book = new_wkbk()
		new_file_name = self.filepath + '\\' + new_wkbk_name
		save(new_file_name)
		filelist.reverse()
		for filename in filelist:
			open_wkbk(filename)
			try:
				sheet_name = self.__get_sheet(filename, sub_string1, sub_string2)
				self.__relocate_sheet(filename, new_wkbk_name, sheet_name)
			except NotFoundError:
				unsuccessful.append(filename)
			close_wkbk(filename)
		active_wkbk(new_wkbk_name)
		if not unsuccessful:
			message = "Compile successful for all files in filelist"
			return message
		else:
			message = """
					  Compile was unable to uniquely identify the sheet to be moved in the following files:

					  {}
					  """
			return message
#####**************************************************************************#####
#####									FUNCTIONS  						       #####
#####**************************************************************************#####
def rename_sheets(prefix):
	"""
	suffix : string
	return : None

	Renames sheets according to suffix + two digit serial
	"""
	sheets = all_sheets()
	codeList = list(itertools.chain(*[[prefix + '0' + str(x) for x in xrange(1, 10)],
									  [prefix + str(x) for x in xrange(10, 100)]]))
	for x in xrange(len(sheets)):
		active_sheet(sheets[x])
		rename_sheet(sheets[x], codeList[x])
	return





