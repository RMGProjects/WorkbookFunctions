Tip: One problem that can occur is that the `get_file_list()` method returns a list that appears to be sorted approximately according to string values of integers that are used in date representation of the file names. For example:

```python
file_list
	['1 APRIL -2013.xls',
	'10 APRIL -2013.xls',
	'11 APRIL -2013.xls',
	'12 APRIL -2013.xls',
	'15 APRIL -2013.xls',
	'16 APRIL -2013.xls',
	'17 APRIL -2013.xls',
	'18 APRIL -2013.xls',
	'20 APRIL -2013.xls',
	'21 APRIL -2013.xls',
	'22 APRIL -2013.xls',
	'23 APRIL -2013.xls',
	'24 APRIL -2013.xls',
	'25 APRIL -2013.xls',
	'2 APRIL -2013.xls',
	'30 APRIL -2013.xls',
	'3 APRIL -2013.xls',
	'4 APRIL -2013.xls',
	'6 APRIL -2013.xls',
	'7 APRIL -2013.xls',
	'8 APRIL -2013.xls',
	'9 APRIL -2013.xls',]
```

The question then is 