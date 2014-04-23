import datetime, itertools, os, random, re, json
from dateutil import parser

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
			
	def get_types(self):
		"""
		return	: dict
		method	: visible
		
		Returns dict of sheet names and type of value found at self.date_cell_ref
		"""
		sheets = all_sheets()
		type_dict = {sheet : object() for sheet in sheets}
		for sheet in sheets:
			active_sheet(sheet)
			type_dict[sheet] = type(Cell(self.date_cell_ref).value)
		return type_dict

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

	def find_duplicates(self):
		"""
		return	: dict
		method	: visible

		Returns a dictionary where dates that are found more than once in the values
		of date_dict created upon the call are the keys, and the values are the 
		keys of date_dict at which the duplicate dates are found.\n

		"""
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

	def relative_order(self):
		"""
		return		: dict
		method		: visible

		Returns a dictionary that shows order of sheets implied by dates in the
		date_dict created upon the call and the actual order of the sheets, if 
		different.\n
		If no date_dict is passed one is created by	calling check_all_dates().\n
		Resulting order may be perverse if there are duplicate dates in date_dict.
		"""

		sheets = all_sheets()
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

	def discontinuities(self, discontinuity_value):
		"""
		discontinuity_value : int
		date_dict	: dict or None
		return		: list of tuples
		method		: visible

		Returns list of tuples where each tuple is a pair of contiguous sheets where
		the	dates found on those sheets indicate a timedelta greater than
		discontinuity_value.\n
		A date_dict is created by calling check_all_dates().
		"""

		sheets = all_sheets()
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

		
	def compare_cell_file_date(self, file_list_dict, regex, strp_format = None):
		"""
		file_list_dict 	: dict (as created when compiling)
		regex			: str
		strp_format		: None or string
		return			: dict
		method			: visible
		
		Returns a dictionary where keys are sheet names and values are tuples where
		first element of the tuple is the date as per the file name taken from 
		file_list_dict, and the second is the date as per date cell taken from a
		date_dict that is created upon the call.\n
		There is only an entry in the dictionary if the two dates in the tuple are 
		not equal.\n
		The file list dict should be that which was created when compiling the 
		sheets. If that object is still in the computer memory, then pass it 
		directly, otherwise open the 'file_list_dict.json' that was created when 
		the sheets were originally compiled.\n
		The user must supply a regular expression (regex) argument that will 
		identify the date component of the file names in the file_list_dict. The 
		user may optionally provide a strp_format argument that is passed to 
		datetime.strptime() in order to convert the date found in the file name to 
		a datetime.date() object. In most cases this will not be necessary as the
		we use dateutil.parser.parse in the code by default. If that is giving
		perverse results, then by all means pass an strp_format argument.\n
		The user is notified in the dictionary if any conversions are impossible, or
		the regular expression does not identify a date like string in the filelists.
		
		Assumes that the sheets are in same order as file list i.e. no changes have
		been made. 
		"""
		date_dict = self.check_all_dates()
		sheets = all_sheets()
		folders = file_list_dict.keys()
		folders.sort()
		re_compiler = re.compile(regex)
		count = 0
		No_file_match = []
		Bad_date_conversion = []
		Mismatches = {}
		
		for folder in folders:
			for file in file_list_dict[folder]:
				result = re_compiler.search(file)
				if not result:
					No_file_match.append(file)
					count +=1
					continue
				else:
					date_group = result.group()
					if strp_format:
						try:
							d_date = datetime.datetime.strptime(date_group, strp_format).date()
						except ValueError:
							Bad_date_conversion.append(file)
							count+=1
							continue
					else:
						try:
							d_date= parser.parse(date_group, fuzzy=True, dayfirst = True).date()
						except ValueError:
							Bad_date_conversion.append(file)
							count+=1
							continue
					if date_dict[sheets[count]] != d_date:
						Mismatches[sheets[count]] = (date_group, date_dict[sheets[count]])
				count+=1
		Mismatches.update({"No date match in file:" : No_file_match, 
						   "Date conversions not possible" : Bad_date_conversion })
		return Mismatches
			

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
	def __init__(self, top_folderpath, **kwargs):
		"""
		top_folderpath	: raw string
		**kwargs		: e.g. folder1 = path/to/file1 	

		Class for compiling a single worksheet from multiple excel workbooks in one
		or more folder of excel workbooks to a single workbook. 

		Initialise the class by passing a string of the top_folderpath where the
		compiled workbook will be stored, and a number of arguments of the format
		folder1 = r'Path/To/Folder
		"""
		if not isinstance(top_folderpath, str):
			raise _InputError("top_folderpath must be a raw string")
		self.top_folderpath = top_folderpath
		if len(kwargs) < 1:
			raise _InputError('Specify at least one kwarg (folder path)')
		self.file_dict = kwargs
		
	def get_file_dict(self):
		"""
		return 	: dict
		method	: visible
		
		Returns a dict of folder names passed during compiler constructions and 
		values that are lists of the files found in the associated folders
		"""
		
		file_list_dict = {key : list(os.listdir(self.file_dict[key])) 
						  for key in self.file_dict.keys()}
		return file_list_dict
		
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
			raise _NotFoundError("Error")
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

	def __save_to_json(self, file_list_dict):
		"""
		file_list_dict	: dict
		filename		: string (including .json extension
		returns			: none
		method:			hidden
		
		Function saves object passed as file_list_dict to json in self.top_folderpath.
		Could be any object in fact, but designed to save the file_list_dict such
		as that created by the get_file_dict() method.\n
		Function to be called in compile_sheets() method below
		"""
		os.chdir(self.top_folderpath)
		with open('final_file_list_dict_' + str(datetime.datetime.now().date()) + '.json', 'w') as out_file:
			json.dump(file_list_dict, out_file)
		return
		
	def compile_sheets(self, file_list_dict, new_wkbk_name, sub_string1, sub_string2 = None):
		"""
		file_dict		: list
		new_wkbk_name	: string
		sub_string1		: string
		sub_string2		: string or None
		return			: formatted string

		Returns formatted string that gives report to user as to success of the
		compile operation.\n
		Function opens every file in the list of files that are the keys of the 
		file_dict passed as argument. Up to two sub-strings may be passed as
		arguments, and the combination of these sub-strings should uniquely identify
		the worksheet to be moved (as there may be more than one).\n
		Sheets that are successfully identified for copying will be copied to a new
		workbook created according to new_wkbk_name. This workbook will be in the
		top_folderpath directory. \n
		Files are opened in the reverse order they are found in the file_dict.
		"""
		
		self.__save_to_json(file_list_dict) #Note __save to json call
		unsuccessful = []
		new_book = new_wkbk()
		new_file_name = self.top_folderpath + '\\' + new_wkbk_name
		save(new_file_name)
		folders = file_list_dict.keys()
		folders.sort()
		folders.reverse()
		
		for folder in folders:
			os.chdir(self.file_dict[folder])
			filelist = file_list_dict[folder][:]
			filelist.reverse()
			for filename in filelist:
				open_wkbk(filename)
				try:
					sheet_name = self.__get_sheet(filename, sub_string1, sub_string2)
					try:
						self.__relocate_sheet(filename, new_wkbk_name, sheet_name)
					except NitroException:
						new_name = sheet_name + str(random.randint(1, 10000000))
						rename_sheet(sheet_name, new_name)
						try:
							self.__relocate_sheet(filename, new_wkbk_name, new_name)
						except NitroException:
							unsuccessful.append(filename)		
				except _NotFoundError:
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
			
class workbook_structure:
	def __init__(self, Dates_class_object, start_row_dict, end_row_dict, cols_list):
		self.__date_dict = Dates_class_object.check_all_dates()
		self.__date_list = [value for key, value in self.__date_dict.iteritems()]
		self.__start_list = [value for key, value in start_row_dict.iteritems()]
		self.__end_list = [value for key, value in start_row_dict.iteritems()]
		self.__cols = cols_list
		if not all(isinstance(date, datetime.date) for date in self.__date_list):
			raise _InputError("All dates values in workbook must be datetime.date objects")
		if not all(isinstance(row, int) for row in self.__start_list):
			raise _InputError("All values in start_row_dict must be integer objects")	
		if not all(isinstance(row, int) for row in self.__end_list):
			raise _InputError("All values in end_row_dict must be integer objects")	
		if not all(isinstance(col, int) for col in self.__cols):
			raise _InputError("All values in cols_list must be integer objects")
		self.workbook_structure = {}
		self.workbook_structure['dates'] = {key : str(date) for key, date in 
											self.__date_dict.iteritems()}
		self.workbook_structure['start_rows'] = {key : row - 1 for key, row in 
												 start_row_dict.iteritems()}
		self.workbook_structure['cols'] = self.__cols
	
	def save_structure(self, top_folderpath):
		if not isinstance(top_folderpath, str):
			raise _InputError("top_folderpath must be a string value")
		os.chdir(top_folderpath)
		with open('workbook_structure_' + str(datetime.datetime.now().date()) + '.json', 'w') as out_file:
			json.dump(self.workbook_structure, out_file)
		return "Save complete"
		

def rename_sheets(prefix):
	"""
	suffix : string
	return : None
	
	Renames sheets according to prefix + two digit serial
	"""
	sheets = all_sheets()
	codeList = list(itertools.chain(*[[prefix + '0' + str(x) for x in xrange(1, 10)], 
									  [prefix + str(x) for x in xrange(10, 100)]]))
	try:
		for x in xrange(len(sheets)):
			active_sheet(sheets[x])
			rename_sheet(sheets[x], codeList[x])
	except NitroException:
		tempnames = [str(random.randint(1, 100000000)) for x in xrange(100)]
		for x in xrange(len(sheets)):
			active_sheet(sheets[x])
			rename_sheet(sheets[x], tempnames[x])
		sheets = all_sheets()
		for x in xrange(len(sheets)):
			active_sheet(sheets[x])
			rename_sheet(sheets[x], codeList[x])
	return		

