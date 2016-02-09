# Gspread-Data-Tracker
Track changes in source data over time in Google Spreadsheets, for example the output of the importHtml command.

To use, go to Resources > Library in the Google Sheets script editor and enter the following project key: 
MlDapdgHJFiRpWKQOyZG0lIoHEfXZl5VD


JSON Spreadsheet Configuration Table

Spreadsheet Object:
Attribute	Type	Purpose	Default
name	String	Easily generate sheet names	Required
sources	Array of source objects	An array of external data sources which will be monitored (see sources below)	Required
id	String	The id of the spreadsheet to use. This can be found within the url:

https://docs.google.com/spreadsheets/d/1VJTPlM0pAMq3B3dGHcXk4rIDdrfIxWUcu8hZY67-LYw/	Output of SpreadsheetApp.getActiveSpreadsheet () method
diff_sheet_name	String	Name of changes sheet	name + '  changes'
diff_sheet_A1_header:	String	Entered in Cell A1 of the difference sheet	Equal to name
diff_sheet_B1_header	String	Entered in Cell B1 of the difference sheet as header for the current values	“Current Values”
history_name	String	Name of changes history sheet	name + ' History'
history_headers	Array of 3 strings	Header strings entered in A1, B1 and C1 of the history sheet	['Date', 'Name', 'Value']
data_sheet_name	String	Name of the data sheet	name + ' Data';
removed_text	String	Appears in a cell when a name belonging to a cell to be monitored can no longer be found in the latest data sheet	“No Longer Found”
value_format	String	Number format for the values, e.g.
 value_format: '@STRING@'	
date_format	String	The date format which will be applied to all cells containing dates, e.g. 
'ddd dd MMMM yyyy HH:mm:ss'	'dd-MM-yyyy HH:mm'
mailto	String	An e-mail address or array of e-mail addresses to send an e-mail to if anything changes.	false (no e-mail sent)
subject	String	The subject for when a mail is sent	name + ‘ has changed’
name_sources	Array of name_source objects	An array of name_sources which will be used to find the cell name to compare against. See name_sources object below. The output of each name sources will be combined to make the final name.	{method: 0} 
(for A1 notation)
Root	Number	Position in the Spreadsheet to start placing newly created sheets	0
exclude_if_any_header_empty	Bool	If any of the name sources return an empty string, the cell will not be monitored for changes	false
split_name_by	String or Regex	After a name is compiled from the list of name sources, if this parameter is set the split command will be run on the string:
name.split(split_name_by)[0]; to find the new name	No string splitting on the name
split_name_index	Number	If split_name_by is set, this will be to index used to find the new name, so the command will be run like this.
name.split(split_name_by)[split_name_index].
Ignored if split_name_by is not set.	0
split_value_by	String or Regex	If set the split command will be run on each cell value:
value.split(split_value_by)[0]; to find the new name. Can be a string or a regular expression, e.g. split_value_by: /\*/g,	No name splitting takes place
split_value_index	Number	If split_value_by is set, this will be to index used to find the new value, so the command will be run like this.
name.split(split_value_by)[split_value_index].
Ignored if split_value_by is not set.	0
replace_in_value	String or Regex	If set, the string or regex will be replaced by the value of replace_in_value_with. This is run after the split command if split_name_by is also set.	false
replace_in_value_with	String	The string or regex in replace_in_value will be replaced with this value. Default is an empty string, so the value being replaced will just be removed.	‘’
archive_time	Number	Time, in minutes, after which the difference Sheet and history Sheet will be archived and new sheets created.	
No archiving takes place.
Range	Array of 4 Numbers	The range of cells in the data sheet to monitor for changes. Format is [firstRow, firstColumn, numRows, numColumns]. If numRows and numColumns are set to 0 then getLastRow() or getLastColumn() will be run on the sheet to get the number of rows or columns to use. Examples of acceptable values are [1,1,2,1] or [1,1,0,0].	Output of sheet.getDataRange() command; the full range of cells corresponding to the dimensions in which data is present
include_rows	Array of numbers	An array of rows (relative to the range) to be monitored for changes. All other rows will not be monitored.	All rows included
include_columns	Array of numbers	An array of columns (relative to the range) to be monitored or changes. All other columns will not be monitored.	All rows included
exclude_rows	Array of numbers	An array of rows (relative to the range) which will not be monitored for changes.	Empty
exclude_columns	Array of numbers	An array of rows (relative to the range) which will not be monitored for changes.	Empty
alwaysFill	Bool	If set to false, only cells who’s values have changed will appear in new columns in the difference sheet. If true, all values will appear regardless of if they changed or not.	false
first_run	Bool	After the script has been run once and  the configuration has been verified, this parameter can be set to false to reduce the number of Spreadsheet API requests, helping the script to fun faster. Leave as true during testing and troubleshooting.	true
sort	Bool	If true; the difference sheet will be sorted by column A	false
default_formatting	Bool	If true, at the end the script will auto-resize all columns, headers will be bolded and frozen, and sheets moved to appropriate positions; for both the difference Sheet and the history Sheet	true

Source Object:
Attribute	Type	Purpose	Default
formula	String	The Google Sheets formula to enter in a cell in the data sheet to retrieve the source data. TrackData includes a queryString variable (a random string beginning with ? which can be appended to urls to ensure the data is refreshed each time TrackData is run). Example command:

‘=IMPORTHTML(“https://en.wikipedia.org/wiki/Ireland’ + TrackData.queryString + ‘”, "table", 4)’	Required
cell	String	The cell, in A1 notation, to add the formula to in the data sheet. Ignored if append is set to true.	‘A1’ (used if neither cell or append is set; so will overwrite existing sources if no cell or append is set for any source)
append	Bool	If true, the formula will automatically be appended after the last row in any data belonging to previous sources	false
append_to_column	String	If append is set to true, the letter representing the column to add the formula to. If append is set to false, this is ignored.	‘A’
space_ before	Number	If append is set to true, the number of empty rows to leave between the last data source and this one. If append is set to false, this is ignored. 	0
timeout	Number	The time in miliseconds to wait for datato be retrieved from an external source	10000
critical	Bool	If true, the script will fail with an exception if the source fails to produce any data within the timeout period. If false, an error will be added to the logfile but the script will continue with the other sources.	false
name	String	If the 4 name source method 4 is used, this will be part of the name for all cells that originated in this source.	Required for source method 4
column_header	Number	If the name source method 5 is used, the column which contains headers relative to the first column in the source’s data	1
row_header	Number	If the name source method 6 is used, the row which contains headers relative to the first row in the source’s data	1

Name Source Object:
Attribute	Type	Purpose	Default
method	Number between 0 and 3	The method by which the name of a cell will be retrieved:

0: A1 Notation. A direct cell-by-cell comparison will take place with the cell name used.
1: Column Header. The name is retrieved from the sheets column header
2: Row Header. The name is retrieved from the sheets’s row header
3: Cell Offset. The name is retrieved from another cell offset from the one being checked.
4: Source name. The name is retrieved from the name parameter which needs to be added to the data source object. This will be used for all cells that come from that source.
5: Source column header: The name is retrieved from the data source’s column header
6. Source row header: The name is retrieved from the data source’s row header.	Required
header	Number	For method 1, the column containing the sheet’s headers. For method 2, the row containing the sheet’s headers.	1
col_offset	Number	For method 3, the column offset from the cell being checked	0 (same column as the cell)
row_offset	Number	For method 3, the row offset from the cell being checked	0 (same row as the cell, if 0 is also used for col_offset, then the same cell contains both name and value so use split_name_by and split_value_by to separate them)


Understanding Cell Names
In addition to the value a cell contains, each cell will be allocated a name. This has two purposes:
1)	Cells will be compared with whatever cell had the same name the last time the script was run; therefore even if the cells containing certain values have changed position, in the data source they can still be compared correctly.
2)	The cell name will be entered in column A of the difference sheet and column B of the history sheet, so it easy to see what the rows refer to.

Cell names are generated according to the following logic:
1)	Each spreadsheet JSON object contains an array of name sources.
2)	Each name source method generates a separate name. Method 0 just gives the name of the cell in A1 notation. This means a direct cell-by-cell comparison will occur. Method 1 and Method 2 can be used to get the name from a header. Method 3 can be used to get the name from a cell offset from the one being checked.
3)	Multiple name sources can be combined. For example, if method 1 and method 2 are used together, the row header and column header will be combined to make the name with a space in between.  Method 0 should not be combined with other methods.
4)	Ideally, whatever names which are returned should be unique. If they aren’t then a number will be appended to repeat occurrences of names.
5)	The script will return an empty string for a name if the cell is to be excluded from being monitored (for example if it appears in the exclude_column list).
6)	If the exclude_if_any_header_empty is set to true, then if any one name source method returns an empty string then the script will return an empty string for the entire name, thus it will no longer be monitored. This can prove useful for easily removing pointless cells for certain data sources from being monitored.
7)	If the splitNameBy parameter is set, a javascript string split method will be run on the name:
name = name.split(splitNameBy)(splitNameIndex).
Be careful that this never returns an empty string or the cell will not be monitored.

Methods
1.	TrackData.runDiffCheck(spreadsheetsJSON, dataSheetOnly)
Run the data comparison. 
spreadsheetsJSON – Mandatory. The JSON configuration which should contain an array of Spreadsheet objects.
dataSheetOnly – Optional. If set to true, only the data sheet containing the data sources will be created and then the script will stop. This is useful for early setup to check what the data source table looks like. Default is false.
2.	TrackData.filterRow(js_spreadsheet, row)
Filter the difference sheet to only show columns for times where a cell’s value changed. 
Js_spreadsheet – Mandatory. The JSON configuration for a single spreadsheet object, e.g. spreadsheetsJSON[0].
row -  Mandatory. The number for the row containing the list of changes for the cell which is being filtered for.
3.	TrackData.showAll(js_spreadsheet)
Show all columns in a spreadsheet (undo filterRow method).
Js_spreadsheet – Mandatory. The JSON configuration for a single spreadsheet object, e.g. spreadsheetsJSON[0].

Common Issues
1.	“Range is invalid” – make sure firstrun is set to true.
