var queryString = Math.random();
  
function applyStripReplacesToArray(stringArray, stripReplaces){
  var newStringArray = [];
    for ( var i in stringArray){  a
       stringArray[i] = applyStripReplaces(stringArray[i], stripReplaces)   
    }
  return stringArray;
}
  
function applyStripReplaces(string, stripReplaces){
  for ( var i in stripReplaces ){
    var stripReplace = stripReplaces[i];
    var text = stripReplace[1];
    if (stripReplace[0]){
      var index = stripReplace[2];
      string = string.split(text)[index] || '';
    }
    else{
      var replaceWith = stripReplace[2];
      string = string.replace(text, replaceWith);
    }
  } 
  string = string || '';
  return string;
}
function DataTracker(config){ 
  this.name = config.name;
  this.sources = config.sources;
  this.spreadsheet_id = config.id || false;
  config.strings = config.strings || {};
  this.dataSheetName = (config.strings.data_sheet_name || '%name% Data').replace('%name%', this.name);
  this.diffSheetName = (config.strings.diff_sheet_name || '%name% changes').replace('%name%', this.name);
  this.a1Title = (config.strings.diff_sheet_A1_header || '%name%').replace('%name%', this.name);
  this.b1Title = (config.strings.diff_sheet_B1_header || 'Current Values').replace('%name%', this.name);
  this.historySheetName = (config.strings.history_name || '%name% History ').replace('%name%', this.name);
  this.historyHeaders = config.strings.history_headers || ['Date', 'Name', 'Value'];
  this.subject = (config.strings.subject || '%name% data has changed').replace('%name%', this.name);
  this.removedTextString = (config.strings.removed_text || "No Longer Found").replace('%name%', this.name);
  this.archiveTime = (config.archive_time || 0) * 60000;
  this.mailto = config.mailto || false;
  this.firstRun = config.first_run || config.first_run==undefined || false;
  this.defaultFormatting = config.default_formatting || config.default_formatting==undefined || false;
  this.sortDiff = config.sort_diff_sheet || false;
  this.sortHist = config.sort_history_sheet || false;
  this.range = config.range || false;
  this.includeRows = config.include_rows || [];
  this.includeColumns = config.include_columns || [];
  this.excludeRows = config.exclude_rows || [];
  this.excludeColumns = config.exclude_columns || [];
  this.excludeIfHeaderEmpty = config.exclude_if_any_header_empty || false;
  this.name_sources = config.name_sources || [{method: 0}];
  this.nameStripReplaces = config.name_strip_replaces || [];
  this.valueStripReplaces = config.value_strip_replaces || [];
  this.dataSheetValueFormat = config.data_sheet_value_format || '@STRING@';
  this.valueFormat = config.value_format || '@STRING@';
  this.dateFormat = config.date_format || 'dd-MM-yyyy HH:mm';
  this.alwaysFill = config.always_fill || false;
  this.archived = false;
  this.getSpreadsheet = function(){
    if (this.spreadsheet_id){
      this.spreadsheet = SpreadsheetApp.openById(this.spreadsheet_id); 
    }
    else{
      this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    } 
  }
  this.getSpreadsheet();
  this.date = new Date(),
      this.formattedDate = [this.date.getMonth()+1,
                 this.date.getDate(),
                 this.date.getFullYear()].join('/')+' '+
                   [this.date.getHours(),
                    this.date.getMinutes(),
                    this.date.getSeconds()].join(':');
  this.archived = false;
  this.namesArr = [];
  this.formatName = function(names){
    var emptyNames = Objects1d.filterByValue(names, '');
    if ( this.excludeIfHeaderEmpty && emptyNames.length > 0 ){
      var name = '';
    }
    else{
      var name = names.join(' ');
      name = applyStripReplaces(name, this.stripReplaces);
      var indexes = Objects1d.filterByValue(this.namesArr, name);
      var length = indexes.length
      if ( length > 0 ){
        var newName = (length + 1).toString();
        name = name + ' ' + newName;
      }
    }
    this.namesArr.push(name);
    return name;
  }
  this.getCellNames = function(){
    var names = [];
    for ( var row in this.rangeValues ){
      names[row] = [];
      for ( var col in this.rangeValues[row] ){
        names[row][col] = [];
      }
    }
    for ( var i in this.name_sources ){ 
      var name_source = this.name_sources[i];
      var stripReplaces = name_source.strip_replaces || false;
      switch(name_source.method) {
        case 0:
          names = this.getNamesByA1Notation(names);
          break;
        case 1:
          var column = name_source.header -1 || 0;
          var headerSize = this.dataRange.getRow() - 1;
          var headerValues = Objects1da.getColumn(this.dataSheet, headerSize, column + 1).getValues();
          names = this.getNamesByHeader(headerValues, names, false, stripReplaces);
          break;
        case 2:
          var row = name_source.header - 1 || 0;
          var headerSize = this.dataRange.getColumn() - 1;
          var headerValues = Objects1d.getRow(this.dataSheet, headerSize, row + 1).getValues();
          names = this.getNamesByHeader(headerValues, names, true, stripReplaces);
          break;
        case 3:
          var colOffset = name_source.col_offset || 0;
          var rowOffset = name_source.row_offset || 0;
          names = this.getNamesByOffset(names, rowOffset, colOffset, stripReplaces);
          break;
        case 4:
          names = this.getNamesFromSource(names);
          break;
        case 5:
          names = this.getNamesBySourceHeader(names, false, stripReplaces);
          break;
        case 6:
          names = this.getNamesBySourceHeader(names, true, stripReplaces);
          break;
      }
    }
    if ( this.includeRows.length == 0){
      for ( var i=0; i < names.length; i++ ){
        this.includeRows.push(i);
      }
    }
    if ( this.includeColumns.length == 0){
      for ( var i=0; i < names[0].length; i++ ){
        this.includeColumns.push(i);
      }
    }
    var formattedNames = [];
    for ( var row in names ){
      formattedNames[row] = [];
      for ( var col in names[row] ){
        var rowNum = Number(row);
        var colNum = Number(col);
        if ( this.excludeRows.indexOf(rowNum) == -1 && this.includeRows.indexOf(rowNum) > -1 && this.excludeColumns.indexOf(colNum) == -1 && this.includeColumns.indexOf(colNum) > -1 ) {
          var name = names[row][col];
          var formattedName = this.formatName(name)
          formattedNames[row][col] = formattedName;
        }
        else{
          formattedNames[row][col] = '';
        }
      }
    }
    return formattedNames;
  }
  this.getCellSource = function(row){
    for ( var i in this.sources ){
      if ( row >= this.sources[i].firstRow && row <= this.sources[i].lastRow ){
        return source;
        break;
      }
    }
    return false;
  }
  this.getChanges = function(){  
    Logger.log('Compiling list of changes');
    this.cellNamesCol = Objects1d.getColumn(this.diffSheet, 0, 1);
    this.previousCol = Objects1d.getColumn(this.diffSheet, 0, 2); 
    this.changesCol = Objects1d.getColumn(this.diffSheet);
    this.previousValues = this.previousCol.getValues();
    this.changedValues = this.changesCol.getValues();
    this.rangeValues = this.dataRange.getValues();
    this.latestCellNames = this.getCellNames();
    var rangeValues = this.rangeValues;
    for ( var row in this.rangeValues ) { 
      for ( var col in this.rangeValues[row] ){
        var cellName = this.latestCellNames[row][col];
        if ( cellName ){
          var cellValue = this.rangeValues[row][col];
          var cellValue = applyStripReplaces(cellValue, this.valueStripReplaces);
          var changedRow = this.cellNamesCol.addValueIfNotExists(cellName);
          if ( cellValue != this.previousValues[changedRow] ){
            this.changedValues[changedRow] = cellValue;
            this.previousValues[changedRow] = cellValue;
          }
          else{
            this.changedValues[changedRow] ='-'; 
          }
        }
      }
    }
    Logger.log(this.changedValues);
  }
  this.getNamesByA1Notation = function(names){
    for ( var i = 0; i < names.length; i++ ){
      var row = i + 1;
      for ( var j = 0; j < names[i].length; j++ ){
        var col = j + 1; 
        var a1Notation = this.dataRange.getCell(row, col).getA1Notation();
        names[i][j].push(a1Notation);
      }
    }
    return names;
  }
  this.getNamesByHeader = function(headerValues, names, headerIsRow, stripReplaces){
    for ( var row in names ){
      for ( var col in names[row] ){
        if (headerIsRow){
          var name = headerValues[col];
        }
        else{
          var name = headerValues[row];
        }  
        name = applyStripReplaces(name, stripReplaces);
        names[row][col].push(name);
      }
    }
    return names; 
  }
  this.getNamesByOffset = function(names, rowOffset, colOffset, stripReplaces){
    var offsetValues = this.dataRange.offset(rowOffset, colOffset).getValues();
    for ( var row in names ){
      for ( var col in names[row] ){
        var name = offsetValues[row][col];
        name = applyStripReplaces(name, stripReplaces);
        names[row][col].push(name);
      }
    }
    return names;
  }
  this.getNamesFromSource = function(names){
    for (var i in this.sources){
      var source = this.sources[i];
      var name = source.name;
      for ( var i=source.firstRow - 1; i < source.lastRow; i++){
        for ( var j in names[i] ){
          names[i][j].push(name);
        }
      }
    }
    return names;
  }
  this.getNamesBySourceHeader = function(names, header, headerIsRow, stripReplaces){
    for (var i in this.sources){
      var source = this.sources[i];
      if ( headerIsRow ){
        var header = source.rowHeader + source.firstRow -1;
        var headerValues = Objects1d.getRow(this.dataSheet, 0, header).getValues();
      }
      else{
        var header = source.colHeader + source.firstColumn -1; 
        var headerValues = Objects1d.getColumn(this.dataSheet, 0, header).getValues();
      }
      for ( var i=source.firstRow - 1; i < source.lastRow; i++){
        for ( var j in names[i] ){
          var name = headerValues[j];
          name = applyStripReplaces(name, stripReplaces);
          names[i][j].push(name);        
        }
      }
    }
    return names;
  }
  this.getValuesString = function(){
    var data = this.dataSheet.getDataRange().getValues().join('').replace(/\,/g, '').replace(/#ERROR!/g, '').replace(/\n/g, '');
    //Logger.info('Values: ' + data);      
    return data
  }
 this.makeDataSheet = function(){
   Logger.log('Making data sheet for ' + this.dataSheetName);
   this.dataSheet = sheetsextra.getOrCreateSheet(this.spreadsheet, this.dataSheetName);
   oldValuesStr =this.dataSheet.getDataRange().getValues().join('');
   if (this.sources.length > 0){
      this.dataSheet.clear();
   }
   for ( var i in this.sources ) { 
     var source = this.sources[i];
     var append = source.append || false;
     if (append){
       var space_before = source.space_before || 0;
       var append_to_column = source.append_to_column || 'A';
       var row = this.dataSheet.getLastRow() + space_before + 1;
       var cell = append_to_column + row.toString();
     }
     else{
       var cell = source.cell || 'A1';
     }
     var cell = this.dataSheet.getRange(cell);
     var cellValuesBefore = this.getValuesString();
     //Logger.log('Cell values before adding formula: ' + cellValuesBefore);
     cell.setValue(source.formula);
     SpreadsheetApp.flush();
     //Logger.log('Cell values after adding formula: ' + this.getValuesString());
     source.firstColumn = cell.getColumn();
     source.firstRow = cell.getRow();
     var wait = 0;
     var waitSeconds = 0;
     var timebetween = 500;
     var timeout = source.timeout || 5000; 
     while (this.getValuesString() == cellValuesBefore){
       SpreadsheetApp.flush();
       Logger.log('Sleeping while waiting for data from source ' + i.toString() + ': ' + waitSeconds.toString() + ' seconds.');
       Utilities.sleep(timebetween);
       wait += timebetween;
       var waitSeconds = wait / 1000;
       if (wait >= timeout){
         Logger.log(this.getValuesString());
         Logger.log('ERROR: Source data for ' + source.formula + ' still empty after ' + waitSeconds.toString() + ' seconds.');
         var critical = source.critical || false;
         if (critical){
           throw new Error('Source data for critical source ' + source.formula + ' still empty after ' + waitSeconds.toString() + ' seconds.');
         }
         else{
           break;
         }
       }
     }
     source.lastColumn = this.dataSheet.getLastColumn()
     source.lastRow = this.dataSheet.getLastRow();
     this.sources[i] = source;
     if ( Number(i) == 0){
       this.firstRow = source.firstRow;
       this.firstColumn = source.firstColumn;
     }
     if ( Number(i) == this.sources.length - 1){
       this.lastColumn = source.lastColumn;
       this.lastRow = source.lastRow;
       this.numColumns = this.lastColumn - this.firstColumn + 1; 
       this.numRows = this.lastRow - this.firstRow + 1;
     }
   }
   if (this.range){   
       this.range[2] = this.range[2] || this.numRows - this.range[0] + 1;
       this.range[3] = this.range[3] || this.numColumns - this.range[1] + 1;
       var range = this.range;
       this.dataRange = this.dataSheet.getRange(this.range[0], this.range[1], this.range[2], this.range[3]);
     }
     else{
       this.dataRange = this.dataSheet.getDataRange();
     }    
   this.dataRange.setNumberFormat(this.dataSheetValueFormat);
   this.allValues = this.dataSheet.getDataRange().getValues();
   var newValuesStr = this.allValues.join('');
   if (newValuesStr == ''){
     throw new Error('All data sources produced empty data. Check the URLs and table indexes.');
   }
   if ( oldValuesStr !=  newValuesStr ){
     Logger.log('Finished making data sheet. Changes detected');
     return true;
   }
   else{
     Logger.log('Finished making data sheet. No changes detected');
     if (this.firstRun){
       return true;
     }
     else{
       return false;
     }
   }
 }
 this.makeDiffSheet = function(){
    Logger.log('Making Difference Sheet');
    this.diffSheet = sheetsextra.getOrCreateSheet(this.spreadsheet, this.diffSheetName);
    if (this.archive){
      sheetsextra.renameSheet(this.spreadsheet, this.diffSheet, this.diffSheetName + ' ' + this.formattedName);
      sheetsextra.moveSheet(this.spreadsheet, this.diffSheet, this.root + 3);
      this.diffSheet = sheetsextra.getOrCreateSheet(this.spreadsheet, this.diffSheetName);
    }
    if ( this.firstRun ){
      this.diffSheet.getRange(1,1,1,2).setValues([[this.a1Title, this.b1Title]]);
    }
    this.getChanges();
   
    var dataChanged = this.changedValues.join('').replace(/-/g, '');
    if ( dataChanged != '') {
      this.updateChanges();
      this.updateHistorySheet();
      if (this.mailto){
        this.sendEmail();
      }
    }
    else{
        Logger.log('No changes detected');
      }
    Logger.log('Finished making difference sheet');
  }
  this.makeHistorySheet = function(){
    this.historySheet = sheetsextra.getOrCreateSheet(this.spreadsheet, this.historySheetName);
    if (this.archiveTime){
      var createdDate = this.historySheet.getRange(2,1).getValue(); 
      if (createdDate){
        var period = this.date - createdDate;
        if ( period > this.archiveTime){
          Logger.log('Archiving');
          var oldHistorySheetName = this.historySheetName + ' ' + this.formattedDate;
          sheetsextra.renameSheet(this.spreadsheet, this.historySheet, this.historySheetName + ' ' + dformat, false);
          sheetsextra.moveSheet(this.spreadsheet, this.historySheet, this.root + 3);
          this.historySheet = sheetsextra.getOrCreateSheet(this.spreadsheet, this.historySheetName);
          this.archived = true;
          this.firstRun = true;
        }
      }
    }
    if ( this.firstRun ){
      Objects1d.getRow(this.historySheet, 0, 1).setValues(this.historyHeaders);
    }
  }
  this.postFormatting = function(){
    if ( this.defaultFormatting ){
      this.runDefaultFormatting();
    }
    if ( this.sortDiff ){
      sheetsextra.sortAll(this.diffSheet, [{column: 1, ascending: true}], 1);
    }
    if ( this.sortHist ){
      sheetsextra.sortAll(this.historySheet, [{column: 1, ascending: false}], 1);
    }
  }
  this.runCheck = function(dataSheetOnly){
    var dataSheetOnly = dataSheetOnly || false;
    Logger.log('Running difference check for ' + this.name);
    var isDifferent = this.makeDataSheet();
    if ( isDifferent && ! dataSheetOnly ){
      this.makeHistorySheet();
      this.makeDiffSheet();
      this.postFormatting();  
    }
  }
  this.runDefaultFormatting = function(){
    if ( this.firstRun ){
      Logger.log('Running firstRun default formatting');
      sheetsextra.freezeHeaders(this.diffSheet, 1, 2);
      sheetsextra.freezeHeaders(this.historySheet, 1, 3);
      this.diffSheet.autoResizeColumn(1);
      this.diffSheet.autoResizeColumn(2);
      sheetsextra.moveSheet(this.spreadsheet, this.diffSheet, 1);
    }
    else{
      Logger.log('Running default formatting');
    }
    SpreadsheetApp.setActiveSheet(this.diffSheet);
    var lastColumn = this.diffSheet.getLastColumn();
    this.diffSheet.autoResizeColumn(lastColumn);
    sheetsextra.autoResizeAll(this.historySheet)
    sheetsextra.boldHeaders(this.diffSheet);
    this.historySheet.getRange(1,1,1,3).setFontWeight('bold');
    Logger.log('Default formatting done');
  }
  this.sendEmail = function(){
    var subject = this.writeEmailSubject(); 
    var mailContents = this.writeEmail();
    Logger.log('Sending e-mail to: ' + this.mailto + ' with subject ' + subject);
    Logger.log('E-mail contents: ' + mailContents);
    try{
      MailApp.sendEmail(this.mailto, subject, mailContents)
    }
    catch(e){
     Logger.log(e); 
    }
  }
  this.updateChanges = function(){
    Logger.info('Updating data on diff sheet');
    this.diffSheetNames = this.cellNamesCol.values;
    var cellNames = Objects1d.to1d(this.latestCellNames);
    for ( row=1; row <= this.diffSheetNames.length - 1; row++ ){
       var name = this.diffSheetNames[row];
       if (name){
         var filtered = Objects1d.filterByValue(cellNames, name);
         if (filtered.length == 0){
            if (this.previousValues[row] != this.removedTextString){
              this.changedValues[row] = this.removedTextString;
              this.previousValues[row] = this.removedTextString;
            }
          }
       }
       else{
         this.diffSheet.deleteRow(row + 1);
       }
    }
    Logger.log('Adding Changes to sheet');
    this.cellNamesCol.commit();
    if (this.valueFormat){
       var numColumns = this.changesCol.column -1;
       var numRows = this.previousValues.length -1;
       this.diffSheet.getRange(2,2, numRows, numColumns).setNumberFormat(this.valueFormat);
    }
    this.previousCol.setValues(this.previousValues); 
    if (this.alwaysFill){
       Logger.log('alwaysFill paramater has been selected'); 
       this.previousValues[0] = this.date;
       this.changesCol.setValues(this.previousValues);
    }
    else{ 
       this.changedValues[0] = this.date;
       this.changesCol.setValues(this.changedValues);
    }
    this.changesCol.getCell(0).setNumberFormat(this.dateFormat).setHorizontalAlignment("left");
    Logger.log('Finished updating changes');
  }
  this.updateHistorySheet = function(){
    var historyRows = [];
    for ( var row=1; row <= this.changedValues.length; row++ ){
      var name = this.diffSheetNames[row];
      var value = this.changedValues[row];
      if (name && value != "-"){
        historyRows.push([this.date, name, value]);
      }
    }
    var numRows = historyRows.length;
    if (numRows > 0){
      var historyNextRow = this.historySheet.getLastRow() + 1;
      if (this.valueFormat){
        this.historySheet.getRange(historyNextRow, 3, numRows).setNumberFormat(this.valueFormat);  
        this.historySheet.getRange(historyNextRow, 1, numRows, 3).setValues(historyRows);
        this.historySheet.getRange(historyNextRow, 1, numRows).setNumberFormat(this.dateFormat).setHorizontalAlignment("left");
      }
    }
  }
  this.filterChangedRows = function(){
    this.not_changed = [];
    var changesList = []
    for ( var i in this.changedValues ){
      if (this.changedValues[i] != '-'){
        Logger.info(i + ' changed');
        changesList.push(this.diffSheetNames[i] + ': ' + this.changedValues[i]);
      }
    }
    Logger.info('Changed Rows : ' + changesList);
    return changesList;
  }
  this.writeEmail = function(){
    changesList = this.filterChangedRows().join('\n\r');
    return changesList;
  }
  this.writeEmailSubject = function(){
    return this.subject;
  }
}

function runDiffCheck(jsonConfigArray, dataSheetOnly, sendLogFileTo) {
  var dataSheetOnly = dataSheetOnly || false;
  var sendLogFileTo = sendLogFileTo || false;
  Logger.log('Running difference check');
  var dataTrackers = [];
  for ( var i in jsonConfigArray ) {
    var config = jsonConfigArray[i];
    var dataTracker = new DataTracker(config);
    dataTracker.runCheck(dataSheetOnly);
    dataTrackers.push(dataTracker);
  }
  Logger.log('Finished running difference check');
  if (sendLogFileTo){
    sheetsextra.sendLogFile(sendLogFileTo, 'Difference Check');
  }
  return dataTrackers;
}

function diffCheck(){
  runDiffCheck(spreadsheetsJSON, false);
}

function getDataSheet(jsonConfig){
  var dataTracker = new DataTracker(jsonConfig);
  this.diffSheet = dataTracker.spreadsheet.getSheetByName(sheetTracker.dataSheetName);
  return sheet;
}
function getDiffSheet(jsonConfig){
  var dataTracker = new DataTracker(jsonConfig);
  this.diffSheet = dataTracker.spreadsheet.getSheetByName(sheetTracker.diffSheetName);
  return sheet;
}
function getHistorySheet(jsonConfig){
  var dataTracker = new DataTracker(jsonConfig);
  this.diffSheet = dataTracker.spreadsheet.getSheetByName(sheetTracker.historySheetName);
  return sheet;
}
function filterRow(jsonConfig, row){
  var sheet = getDiffSheet(jsonConfig);
  var row = Objects1d.getRow(sheet, 0, row, 1, 1);
  var columnsToHide = row.filter('-');
  for ( var col in columnsToHide ){
    sheet.hideColumns(columnsToHide[col] + 1);
  }
  sheet.setActiveRange(row.getRange());
}
function showAll(jsonConfig){
  var sheet = getDiffSheet(jsonConfig);
  sheet.showColumns(1, sheet.getLastColumn());
  sheet.showRows(1, sheet.getLastRow());
}
