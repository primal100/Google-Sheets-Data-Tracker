# Google-Sheets-Data-Tracker
Track changes in source data over time in Google Spreadsheets, for example the output of the importHtml command.

There is no GUI interface. Configuration is done through JSON configuration. Open the script editor in Google Spreadsheets, and load the Gspread-Data-Tracker library:

To use, create a new spreadsheet and go to Resources > Library in the Google Sheets script editor. Enter the following project key: 
MlDapdgHJFiRpWKQOyZG0lIoHEfXZl5VD

Then enter the following javascript code into your script:
```javascript
spreadsheetsJSON = []

function track() {
  TrackData.runDiffCheck(spreadsheetsJSON, false)
}
```

Then create a trigger to run the track() function repeatedly, as often as required.

##How Gspread-Data-Tracker works##

Obviously the above example won't do anything as spreadsheetsJSON is an empty list.

The input into the tracker is json, which configures which data to track, how to look for patterns in that data so that the correct cells are compared even if the format changes, and how to label the cells.

The output is a "difference sheet" which tracks how data changes over time and a "history sheet" which lists all changes in chronological order.

Examples
-------------
Example script to track changes in the English Premier League table using a Wikipedia table as source data

```javascript
spreadsheetsJSON = [
  {
    'name': 'Premier League Table',                                       #Used to generate sheet names
    'id': '1bqrA2tYlvN0DO5G6lVF1T0ldAuP6ykP-qfHatbx_fmc',                 #Id of spreadsheet
    'exclude_rows': [0],                                                  #Exclude header row from being tracked for changes
    'exclude_columns': [1, 10],                                           #Exclude header column(team names) and another text column
    'sources': [      #Formula includes TrackData.queryString which makes sure data is refreshed every time script is run
      {
        'formula': '=IMPORTHTML("https://en.wikipedia.org/wiki/2015-16_Premier_League?' + TrackData.queryString + '", "table", 5)',
        'critical': true
      }
      ],
    'name_sources': [
      {
        'method': 1,              #Column Header
        'header': 2,              #Use Column 2 as header(team names)
      },
      {
        'method': 2               #Row Header (Pos, pts, etc, uses Row 1 by default)
      }
      ]
  }
]

function track() {
  TrackData.runDiffCheck(spreadsheetsJSON, false)
}
```

And here is the output with a trigger setup to run every 2 hours:

https://docs.google.com/spreadsheets/d/1bqrA2tYlvN0DO5G6lVF1T0ldAuP6ykP-qfHatbx_fmc/pubhtml

##Understanding Cell Names##
In addition to the value a cell contains, each cell will be allocated a name. This has two purposes:  
1.  Cells will be compared with whatever cell had the same name the last time the script was run; therefore even if the cells containing certain values have changed position, in the data source they can still be compared correctly.  
2.  The cell name will be entered in column A of the difference sheet and column B of the history sheet, so it easy to see what the rows refer to.

Cell names are generated according to the following logic:  
1.	Each spreadsheet JSON object contains an array of name sources.  
2.	Each name source method generates a separate name. Method 0 just gives the name of the cell in A1 notation. This means a direct cell-by-cell comparison will occur. Method 1 and Method 2 can be used to get the name from a header. Method 3 can be used to get the name from a cell offset from the one being checked. Method 4 is a name associated to the source with the name paramater. Method 5 gets a name from the column header associated to the source. Method  6 gets a name from the row header associated to the source.  
3.	An array of javascript strip and replace commands will be run on each name source if the strip_replaces parameter is configured for that name source.  
4.	Multiple name sources can be combined. For example, if method 1 and method 2 are used together, the row header and column header will be combined to make the name with a space in between.  Method 0 should not be combined with other methods.  
5.	Ideally, whatever names which are returned should be unique. If they aren’t then a number will be appended to repeat occurrences of names.  
6.	The script will return an empty string for a name if the cell is to be excluded from being monitored (for example if it appears in the exclude_column list, or is outside the selected range parameter value).  
7.	An array of javascript strip and replace functions will be run on the entire name if the strip_replaces array parameter is set.  
8.	If the exclude_if_any_header_empty is set to true, then if any one name source method returns an empty string then the script will return an empty string for the entire name, thus it will no longer be monitored. This can prove useful for easily removing pointless cells for certain data sources from being monitored.  
Be careful that this never returns an empty string or the cell will not be monitored.

The best way to understand this is via examples.

Methods
-------------

```javascript
TrackData.runDiffCheck(jsonConfig, dataSheetOnly)
```
Run the data comparison

**Arguments:**  
**jsonConfig** – Mandatory. The JSON configuration which should contain an array of Spreadsheet objects.  
**dataSheetOnly** – Optional. If set to true, only the data sheet containing the data sources will be created and then the script will stop. This is useful for early setup to check what the data source table looks like. Default is false.

```javascript
TrackData.getHistorySheet(jsonConfig)
```
Returns the spreadsheet containing a history of all changes to the data

**Arguments:**  
**jsonConfig** – Mandatory. The JSON configuration for a single spreadsheet object, e.g. spreadsheetsJSON[0].

```javascript
TrackData.getDiffSheet(jsonConfig)
```
Returns the spreadsheet containing the changes to data over time.

**Arguments:**  
**jsonConfig** – Mandatory. The JSON configuration for a single spreadsheet object, e.g. spreadsheetsJSON[0].

```javascript
TrackData.getDataSheet(jsonConfig)
```
Returns the spreadsheet containing the latest data.

**Arguments:**  
**jsonConfig** – Mandatory. The JSON configuration for a single spreadsheet object, e.g. spreadsheetsJSON[0].

```javascript
TrackData.filterRow(jsonConfig, row)
```
Filter the difference sheet to only show columns for times where a cell’s value changed. 

**Arguments:**  
**jsonConfig** – Mandatory. The JSON configuration for a single spreadsheet object, e.g. spreadsheetsJSON[0].
**row** -  Mandatory. The number for the row containing the list of changes for the cell which is being filtered for.

```javascript
TrackData.showAll(jsonConfig)
```
Show all columns in a spreadsheet (undo filterRow method).

**Arguments:**  
**jsonConfig** – Mandatory. The JSON configuration for a single spreadsheet object, e.g. spreadsheetsJSON[0].

JSON Spreadsheet Configuration Tables
-------------

###Spreadsheet Object:###


|    Attribute                      |    Type                            |    Default                                  |    Purpose                                                                                                                                                                                                                                                                                                                                     |
|-----------------------------------|------------------------------------|---------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|    name                           |    String                          |    Required                                 |    Easily generate sheet names                                                                                                                                                                                                                                                                                                                 |
|    sources                        |    Array of source objects         |    Required                                 |    An array of external data sources which will be monitored (see   sources below)                                                                                                                                                                                                                                                             |
|    strings                        |    Array of strings                |    Use all defaults                         |                                                                                                                                                                                                                                                                                                                                                |
|    id                             |    String                          |    SpreadsheetApp.getActiveSpreadsheet()    |    The id of the spreadsheet to use. This can be found within the url:       https://docs.google.com/spreadsheets/d/1VJTPlM0pAMq3B3dGHcXk4rIDdrfIxWUcu8hZY67-LYw/                                                                                                                                                                              |
|    value_format                   |    String                          |    '@STRING@'                               |    Number format for the values, e.g.    value_format: '@STRING@' –   changing from default may stop the script from working                                                                                                                                                                                                                   |
|    date_format                    |    String                          |    'dd-MM-yyyy HH:mm'                       |    The date format which will be applied to all cells containing dates,   e.g.    'ddd dd MMMM yyyy HH:mm:ss'                                                                                                                                                                                                                                  |
|    mailto                         |    String                          |    false (no e-mail sent)                   |    An e-mail address or array of e-mail addresses to send an e-mail to   if anything changes.                                                                                                                                                                                                                                                  |
|    name_sources                   |    Array of name_source objects    |    {method: 0}    (for A1 notation)         |    An array of name_sources which will be used to find the cell name to   compare against. See name_sources object below. The output of each name   sources will be combined to make the final name.                                                                                                                                           |
|    Root                           |    Number                          |    0                                        |    Position in the Spreadsheet to start placing newly created sheets                                                                                                                                                                                                                                                                           |
|    exclude_if_any_header_empty    |    Bool                            |    false                                    |    If any of the name sources return an empty string, the cell will not   be monitored for changes                                                                                                                                                                                                                                             |
|    name_strip_replaces            |    Array                           |    []                                       |    List of javascript strip and replace commands to run on each cell   name                                                                                                                                                                                                                                                                    |
|    value_strip_replaces           |    Array                           |    []                                       |    List of javascript strip and replace commands to run on each cell   value                                                                                                                                                                                                                                                                   |
|    archive_time                   |    Number                          |    No archiving takes place.                |    Time, in minutes, after which the difference Sheet and history Sheet   will be archived and new sheets created.                                                                                                                                                                                                                             |
| range                             |    Array of 4 Numbers              |    sheet.getDataRange()                     |    The range of cells in the data sheet to monitor for changes. Format   is [firstRow, firstColumn, numRows, numColumns]. If numRows and numColumns   are set to 0 then getLastRow() or getLastColumn() will be run on the sheet to   get the number of rows or columns to use. Examples of acceptable values are   [1,1,2,1] or [1,1,0,0].    |
|    include_rows                   |    Array of numbers                |    All rows included                        |    An array of rows (relative to the range) to be monitored for changes.   All other rows will not be monitored.                                                                                                                                                                                                                               |
|    include_columns                |    Array of numbers                |    All rows included                        |    An array of columns (relative to the range) to be monitored or   changes. All other columns will not be monitored.                                                                                                                                                                                                                          |
|    exclude_rows                   |    Array of numbers                |    Empty                                    |    An array of rows (relative to the range) which will not be monitored   for changes.                                                                                                                                                                                                                                                         |
|    exclude_columns                |    Array of numbers                |    Empty                                    |    An array of rows (relative to the range) which will not be monitored   for changes.                                                                                                                                                                                                                                                         |
|    always_fill                    |    Bool                            |    false                                    |    If set to false, only cells who’s values have changed will appear in   new columns in the difference sheet. If true, all values will appear   regardless of if they changed or not.                                                                                                                                                         |
|    first_run                      |    Bool                            |    true                                     |    After the script has been run once and  the configuration has been verified, this   parameter can be set to false to reduce the number of Spreadsheet API   requests, helping the script to fun faster. Leave as true during testing and   troubleshooting.                                                                                 |
|    sort_diff_sheet                |    Bool                            |    false                                    |    If true; the difference sheet will be sorted by column A                                                                                                                                                                                                                                                                                    |
|    sort_history_sheet             |    Bool                            |    false                                    |    If true; the history sheet will be sorted by column A                                                                                                                                                                                                                                                                                       |
|    default_formatting             |    Bool                            |    true                                     |    If true, at the end the script will auto-resize all columns, headers   will be bolded and frozen, and sheets moved to appropriate positions; for   both the difference Sheet and the history Sheet                                                                                                                                        | 

###Strings Object:###

|    Attribute                |    Type                  |    Purpose                      |    Default                                                                                                                |
|-----------------------------|--------------------------|---------------------------------|---------------------------------------------------------------------------------------------------------------------------|
|    data_sheet_name          |    String                |    name + '  Data'              |    Name of data sheet                                                                                                     |
|    diff_sheet_name          |    String                |    name + '  changes'           |    Name of changes sheet                                                                                                  |
|    diff_sheet_A1_header:    |    String                |    Equal to name                |    Entered in Cell A1 of the difference sheet                                                                             |
|    diff_sheet_B1_header     |    String                |    “Current Values”             |    Entered in Cell B1 of the difference sheet as header for the current   values                                          |
|    history_name             |    String                |    name + ' History'            |    Name of changes history sheet                                                                                          |
|    history_headers          |    Array of 3 strings    |    ['Date', 'Name', 'Value']    |    Header strings entered in A1, B1 and C1 of the history sheet                                                           |
|    removed_text             |    String                |    “No Longer Found”            |    Appears in a cell when a name belonging to a cell to be monitored can   no longer be found in the latest data sheet    |
|    subject                  |    String                |    name + ‘ has changed’        |    The subject for when a mail is sent                                                                                    |
###Source Object:###

|    Attribute           |    Type      |    Default                              |    Purpose                                                                                                                                                                                                                                                                                                                                                                                              |
|------------------------|--------------|-----------------------------------------|---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|    formula             |    String    |    Required                             |    The Google Sheets formula to enter in a cell in the data sheet to   retrieve the source data. TrackData includes a queryString variable (a random   string beginning with ? which can be appended to urls to ensure the data is   refreshed each time TrackData is run). Example command:       ‘=IMPORTHTML(“https://en.wikipedia.org/wiki/Ireland’   + TrackData.queryString + ‘”, "table", 4)’    |
|    cell                |    String    |    ‘A1'                                 |    The cell, in A1 notation, to add the formula to in the data sheet.   Ignored if append is set to true.                                                                                                                                                                                                                                                                                               |
|    append              |    Bool      |    false                                |    If true, the formula will automatically be appended after the last row   in any data belonging to previous sources                                                                                                                                                                                                                                                                                   |
|    append_to_column    |    String    |    ‘A’                                  |    If append is set to true, the letter representing the column to add   the formula to. If append is set to false, this is ignored.                                                                                                                                                                                                                                                                    |
|    space_ before       |    Number    |    0                                    |    If append is set to true, the number of empty rows to leave between   the last data source and this one. If append is set to false, this is   ignored.                                                                                                                                                                                                                                               |
|    timeout             |    Number    |    10000                                |    The time in miliseconds to wait for datato be retrieved from an   external source                                                                                                                                                                                                                                                                                                                    |
|    critical            |    Bool      |    false                                |    If true, the script will fail with an exception if the source fails   to produce any data within the timeout period. If false, an error will be   added to the logfile but the script will continue with the other sources.                                                                                                                                                                          |
|    name                |    String    |    Required for name source method 4    |    If the 4 name source method 4 is used, this will be part of the name   for all cells that originated in this source.                                                                                                                                                                                                                                                                                 |
|    column_header       |    Number    |    1                                    |    If the name source method 5 is used, the column which contains   headers relative to the first column in the source’s data                                                                                                                                                                                                                                                                           |
|    row_header          |    Number    |    1                                    |    If the name source method 6 is used, the row which contains headers   relative to the first row in the source’s data                                                                                                                                                                                                                                                                                 |

###Name Source Object:###

|    Attribute         |    Type                      |    Purpose                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |    Default                                                                                                                                                                     |
|----------------------|------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
|    method            |    Number between 0 and 5    |    The method by which the name of a cell will be retrieved:       0: A1 Notation. A direct cell-by-cell comparison will take place with   the cell name used.   1: Column Header. The name is retrieved from the sheets column header   2: Row Header. The name is retrieved from the sheets’s row header   3: Cell Offset. The name is retrieved from another cell offset from   the one being checked.   4: Source name. The name is retrieved from the name parameter which   needs to be added to the data source object. This will be used for all cells   that come from that source.   5: Source column header: The name is retrieved from the data source’s   column header   6. Source row header: The name is retrieved from the data source’s   row header.    |    Required                                                                                                                                                                    |
|    header            |    Number                    |    For method 1, the column containing the sheet’s headers. For method   2, the row containing the sheet’s headers.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |    1                                                                                                                                                                           |
|    strip_replaces    |    Array                     |    []                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |    Array of javascript strip and replace commands to run on the name   source                                                                                                  |
|    col_offset        |    Number                    |    For method 3, the column offset from the cell being checked                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             |    0 (same column as the cell)                                                                                                                                                 |
|    row_offset        |    Number                    |    For method 3, the row offset from the cell being checked                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |    0 (same row as the cell, if 0 is also used for col_offset, then the   same cell contains both name and value so use split_name_by and split_value_by   to separate them)    |

Common Issues  
1. “Range is invalid” – make sure firstrun is set to true.


