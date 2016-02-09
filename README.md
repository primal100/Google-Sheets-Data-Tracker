# Gspread-Data-Tracker
Track changes in source data over time in Google Spreadsheets, for example the output of the importHtml command.

There is no GUI interface. Configuration is done through JSON configuration. Open the script editor in Google Spreadsheets, and load the Gspread-Data-Tracker library:

To use, go to Resources > Library in the Google Sheets script editor and enter the following project key: 
MlDapdgHJFiRpWKQOyZG0lIoHEfXZl5VD

Then enter the following text.
```javascript
spreadsheetsJson = []

function track() {
  TrackData.runDiffCheck(spreadsheetsJSON, false)
}
```

Then create a trigger to run the track() function repeatedly, as often as required.

Examples for spreadsheetJson coming...
