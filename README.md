# Google Apps Script Service Attendant

This library adds some missing functionality to Google Apps Script services. The library is written in Coffeescript and compiled to Javascript. Please use the Javascript version of the library. :) 

## Usage


Copy the contents of the compiled javascript file into a new file in your project. 
 
Replace the top level Google App Script service objects with their Attendant equivalent. Then use the services as you normally would, except now you can take advantage of the extras that the service attendants provide.

```
Logger.log(SpreadsheetApp.getActiveSheet().getName());
```

becomes

```
LoggerAttendant.info(SpreadsheetAppAttendant.getActiveSheet().getName());
```

**Note: You only have to replace the top level service object. The subsequent objects will automatically be wrapped with their Attendant equivalents.**
 
## Current Service Attendants

* [SpreadsheetAppAttendant](#spreadsheetappattendant)
* [SpreadsheetAttendant](#spreadsheetattendant)
* [SheetAttendant](#sheetattendant)
* [RangeAttendant](#rangeattendant)
* PropertiesServiceAttendant
* ScriptPropertiesAttendant
* UserPropertiesAttendant
* DocumentPropertiesAttendant
* LoggerAttendant

## SpreadsheetAppAttendant

This attendant is only used to wrap the other services though a bit of meta programming magic. No extras are available at the moment.

## SpreadsheetAttendant  
### Methods

| Method | Return Type | Brief description |
| ------ | ----------- | ----------------- |
| [appendRowReturnRange(data)](#spreadsheetattendantappendrowreturnrangedata) | [RangeAttendant](#rangeattendant) | Appends data to the active sheet, searches the sheet for the appended data, and returns the range of the appended row. |
| [columnIterator()](#spreadsheetattendantcolumniterator) | [RangeAttendantIterator](#rangeattendantiterator) | Get a range iterator that iterates over all the columns in the active sheet of the spreadsheet. |
| [dataColumnIterator()](#spreadsheetattendantdatacolumniterator) | [RangeAttendantIterator](#rangeattendantiterator) | Get a range iterator that iterates over all the data columns in the active sheet of the spreadsheet. |
| [dataRowIterator()](#spreadsheetattendantdatarowiterator) | [RangeAttendantIterator](#rangeattendantiterator) | Get a range iterator that iterates over all the data rows in the active sheet of the spreadsheet. |
| [eachColumn(callback)](#spreadsheetattendanteachcolumncallback) | [SpreadsheetAttendant](#spreadsheetattendant) | Execute callback for each column in the active sheet of the spreadsheet starting from the left and moving right. |
| [eachColumnReverse(callback)](#spreadsheetattendanteachcolumnreversecallback) | [SpreadsheetAttendant](#spreadsheetattendant) | Execute callback for each column in the active sheet of the spreadsheet starting from the right and moving left. |
| [eachDataColumn(callback)](#spreadsheetattendanteachdatacolumncallback) | [SpreadsheetAttendant](#spreadsheetattendant) | Execute callback for each data column in the active sheet of the spreadsheet starting from the left and moving right. |
| [eachDataColumnReverse(callback)](#spreadsheetattendanteachdatacolumnreversecallback) | [SpreadsheetAttendant](#spreadsheetattendant) | Execute callback for each data column in the active sheet of the spreadsheet starting from the right and moving left. |
| [eachDataRow(callback)](#spreadsheetattendanteachdatarowcallback) | [SpreadsheetAttendant](#spreadsheetattendant) | Execute callback for each data row in the active sheet of the spreadsheet starting from the top and moving down. |
| [eachDataRowReverse(callback)](#spreadsheetattendanteachdatarowreversecallback) | [SpreadsheetAttendant](#spreadsheetattendant) | Execute callback for each data row in the active sheet of the spreadsheet starting from the bottom and moving up. |
| [eachRow(callback)](#spreadsheetattendanteachrowcallback) | [SpreadsheetAttendant](#spreadsheetattendant) | Execute callback for each row in the active sheet of the spreadsheet starting from the top and moving down. |
| [eachRowReverse(callback)](#spreadsheetattendanteachrowreversecallback) | [SpreadsheetAttendant](#spreadsheetattendant) | Execute callback for each row in the active sheet of the spreadsheet starting from the bottom and moving up. |
| [getEntireRange()](#spreadsheetattendantgetentirerange) | [RangeAttendant](#rangeattendant) | Get a range that contains all the columns and rows of the active sheet. |
| [rowIterator()](#spreadsheetattendantrowiterator) | [RangeAttendantIterator](#rangeattendantiterator) | Get a range iterator that iterates over all the rows in the active sheet of the spreadsheet. |


### SpreadsheetAttendant.appendRowReturnRange(data)

Appends data to the active sheet, searches the sheet for the appended data starting at the bottom, and returns the range of the appended row.

```
  function moveRow(row) {
    var newRow = SpreadsheetAppAttendant.getSpreadsheet().appendRowReturnRange(row.getValues()[0]);
    row.clear();
    SpreadsheetAppAttendant.getSpreadsheet().setActiveRange(newRow);
  }  
```

##### Return
[RangeAttendant](#rangeattendant)  

### SpreadsheetAttendant.columnIterator()

Get a range iterator that iterates over all the columns in the active sheet of the spreadsheet.

```
  var iterator = SpreadsheetAppAttendant.getSpreadsheet().columnIterator();
  while (iterator.hasNext()) {
    LoggerAttendant.info(iterator.next().getA1Notation());
  }
```

##### Return
[RangeAttendantIterator](#rangeattendantiterator)  

### SpreadsheetAttendant.dataColumnIterator()

Get a range iterator that iterates over all the data columns in the active sheet of the spreadsheet.

```
  var iterator = SpreadsheetAppAttendant.getSpreadsheet().dataColumnIterator();
  while (iterator.hasNext()) {
    LoggerAttendant.info(iterator.next().getA1Notation());
  }
```

##### Return
[RangeAttendantIterator](#rangeattendantiterator)   

### SpreadsheetAttendant.dataRowIterator()

Get a range iterator that iterates over all the data rows in the active sheet of the spreadsheet.

```
  var iterator = SpreadsheetAppAttendant.getSpreadsheet().dataRowIterator();
  while (iterator.hasNext()) {
    LoggerAttendant.info(iterator.next().getA1Notation());
  }
```

##### Return
[RangeAttendantIterator](#rangeattendantiterator)        

### SpreadsheetAttendant.eachColumn(callback)

Execute callback for each column in the active sheet of the spreadsheet starting from the left and moving right.

#### Parameters

| Name | Type | Description |
|------|------|-------------|
|callback | Function | The function to call with each column in the active sheet. | 

```
  SpreadsheetAppAttendant.getSpreadsheet().eachColumn(function (column) {
    LoggerAttendant.info(column.getA1Notation());
  });
```

##### Return
[SpreadsheetAttendant](#spreadsheetattendant) - for chaining

### SpreadsheetAttendant.eachColumnReverse(callback)

Execute callback for each column in the active sheet of the spreadsheet starting from the right and moving left.

#### Parameters

| Name | Type | Description |
|------|------|-------------|
|callback | Function | The function to call with each column in the active sheet. | 

```
  SpreadsheetAppAttendant.getSpreadsheet().eachColumnReverse(function (column) {
    LoggerAttendant.info(column.getA1Notation());
  });
```

##### Return
[SpreadsheetAttendant](#spreadsheetattendant) - for chaining

### SpreadsheetAttendant.eachDataColumn(callback)

Execute callback for each data column in the active sheet of the spreadsheet starting from the left and moving right.

#### Parameters

| Name | Type | Description |
|------|------|-------------|
|callback | Function | The function to call with each data column in the active sheet. | 

```
  SpreadsheetAppAttendant.getSpreadsheet().eachDataColumn(function (column) {
    LoggerAttendant.info(column.getA1Notation());
  });
```

##### Return
[SpreadsheetAttendant](#spreadsheetattendant) - for chaining

### SpreadsheetAttendant.eachDataColumnReverse(callback)

Execute callback for each data column in the active sheet of the spreadsheet starting from the right and moving left.

#### Parameters

| Name | Type | Description |
|------|------|-------------|
|callback | Function | The function to call with each data column in the active sheet. | 

```
  SpreadsheetAppAttendant.getSpreadsheet().eachDataColumnReverse(function (column) {
    LoggerAttendant.info(column.getA1Notation());
  });
```

##### Return
[SpreadsheetAttendant](#spreadsheetattendant) - for chaining   

### SpreadsheetAttendant.eachDataRow(callback)

Execute callback for each data row in the active sheet of the spreadsheet starting from the top and moving down.

#### Parameters

| Name | Type | Description |
|------|------|-------------|
|callback | Function | The function to call with each data row in the active sheet. | 

```
  SpreadsheetAppAttendant.getSpreadsheet().eachDataRow(function (row) {
    LoggerAttendant.info(row.getA1Notation());
  });
```

##### Return
[SpreadsheetAttendant](#spreadsheetattendant) - for chaining

### SpreadsheetAttendant.eachDataRowReverse(callback)

Execute callback for each data row in the active sheet of the spreadsheet starting from the bottom and moving up.

#### Parameters

| Name | Type | Description |
|------|------|-------------|
|callback | Function | The function to call with each data row in the active sheet. | 

```
  SpreadsheetAppAttendant.getSpreadsheet().eachDataRowReverse(function (row) {
    LoggerAttendant.info(row.getA1Notation());
  });
```

##### Return
[SpreadsheetAttendant](#spreadsheetattendant) - for chaining  

### SpreadsheetAttendant.eachRow(callback)

Execute callback for each row in the active sheet of the spreadsheet starting from the top and moving down.

#### Parameters

| Name | Type | Description |
|------|------|-------------|
|callback | Function | The function to call with each row in the active sheet. | 

```
  SpreadsheetAppAttendant.getSpreadsheet().eachRow(function (row) {
    LoggerAttendant.info(row.getA1Notation());
  });
```

##### Return
[SpreadsheetAttendant](#spreadsheetattendant) - for chaining

### SpreadsheetAttendant.eachRowReverse(callback)

Execute callback for each row in the active sheet of the spreadsheet starting from the bottom and moving up.

#### Parameters

| Name | Type | Description |
|------|------|-------------|
|callback | Function | The function to call with each row in the active sheet. | 

```
  SpreadsheetAppAttendant.getSpreadsheet().eachRowReverse(function (row) {
    LoggerAttendant.info(row.getA1Notation());
  });
```

##### Return
[SpreadsheetAttendant](#spreadsheetattendant) - for chaining

### SpreadsheetAttendant.getEntireRange()

Get a range that contains all the columns and rows of the active sheet.

```javascript
  var range = SpreadsheetAppAttendant.getSpreadsheet().getEntireRange();
```

##### Return
[RangeAttendant](#rangeattendant)

### SpreadsheetAttendant.rowIterator()

Get a range iterator that iterates over all the rows in the active sheet of the spreadsheet.

```
  var iterator = SpreadsheetAppAttendant.getSpreadsheet().rowIterator();
  while (iterator.hasNext()) {
    LoggerAttendant.info(iterator.next().getA1Notation());
  }
```

##### Return
[RangeAttendantIterator](#rangeattendantiterator)

## SheetAttendant
### Methods

| Method | Return Type | Brief description |
| ------ | ----------- | ----------------- |
| [appendRowReturnRange(data)](#sheetattendantappendrowreturnrangedata) | [RangeAttendant](#rangeattendant) | Appends data to the sheet, searches the sheet for the appended data, and returns the range of the appended row. |
| [columnIterator()](#sheetattendantcolumniterator) | [RangeAttendantIterator](#rangeattendantiterator) | Get a range iterator that iterates over all the columns in the sheet. |
| [dataColumnIterator()](#sheetattendantdatacolumniterator) | [RangeAttendantIterator](#rangeattendantiterator) | Get a range iterator that iterates over all the data columns in the sheet. |
| [dataRowIterator()](#sheetattendantdatarowiterator) | [RangeAttendantIterator](#rangeattendantiterator) | Get a range iterator that iterates over all the data rows in the sheet. |
| [eachColumn(callback)](#sheetattendanteachcolumncallback) | [SheetAttendant](#sheetattendant) | Execute callback for each column in the sheet starting from the left and moving right. |
| [eachColumnReverse(callback)](#sheetattendanteachcolumnreversecallback) | [SheetAttendant](#sheetattendant) | Execute callback for each column in the sheet starting from the right and moving left. |
| [eachDataColumn(callback)](#sheetattendanteachdatacolumncallback) | [SheetAttendant](#sheetattendant) | Execute callback for each data column in the sheet starting from the left and moving right. |
| [eachDataColumnReverse(callback)](#sheetattendanteachdatacolumnreversecallback) | [SheetAttendant](#sheetattendant) | Execute callback for each data column in the sheet starting from the right and moving left. |
| [eachDataRow(callback)](#sheetattendanteachdatarowcallback) | [SheetAttendant](#sheetattendant) | Execute callback for each data row in the sheet starting from the top and moving down. |
| [eachDataRowReverse(callback)](#sheetattendanteachdatarowreversecallback) | [SheetAttendant](#sheetattendant) | Execute callback for each data row in the sheet starting from the bottom and moving up. |
| [eachRow(callback)](#sheetattendanteachrowcallback) | [SheetAttendant](#sheetattendant) | Execute callback for each row in the sheet starting from the top and moving down. |
| [eachRowReverse(callback)](#sheetattendanteachrowreversecallback) | [SheetAttendant](#sheetattendant) | Execute callback for each row in the sheet starting from the bottom and moving up. |
| [getEntireRange()](#sheetattendantgetentirerange) | [RangeAttendant](#rangeattendant) | Get a range that contains all the columns and rows of the sheet. |
| [rowIterator()](#sheetattendantrowiterator) | [RangeAttendantIterator](#rangeattendantiterator) | Get a range iterator that iterates over all the rows in the sheet. |


### SheetAttendant.getEntireRange()

Get a range that contains all the columns and rows of the sheet.

```javascript
  var range = SpreadsheetAppAttendant.getActiveSheet().getEntireRange();
```

##### Return
[RangeAttendant](#rangeattendant)

## RangeAttendant
### Methods

| Method | Return Type | Brief description |
| ------ | ----------- | ----------------- |
| [includeAllColumns()](#rangeattendantincludeallcolumns) | [RangeAttendant](#rangeattendant) | Expand the range to include all the columns of the rows in the range. |
| [isBlank()](#rangeattendantisblank) | Boolean | Sometimes Range.isBlank() throws errors. This provides a backup implementation. |
| [removeHeader()](#rangeattendantremoveheader) | [RangeAttendant](#rangeattendant) | Remove the header row from the range if it is included. If only the header row is included in the range return null. |
| [sliceRows(start, length)](#rangeattendantslicerows) | [RangeAttendant](#rangeattendant) | Get a subset of a range including the row at the *start* index through *length* rows or end of the range. |
| [sliceColumns(start, length)](#rangeattendantslicecolumns) | [RangeAttendant](#rangeattendant) | Get a subset of a range including the columns at the *start* index through *length* rows or end of the range. |
| [slice(startRow, startColumn, rowLength, columnLength)](#rangeattendantslicecolumns) | [RangeAttendant](#rangeattendant) | Get a subset of a range including the rows and columns at the *startRow* index and *startColumn* index through *rowLength* rows and *columnLength* columns or end of the range. |
| [rowIterator()](#rangeattendantrowiterator) | [RangeAttendantIterator](#rangeattendantiterator) | Get a row iterator. |
| [columnIterator()](#rangeattendantcolumniterator) | [RangeAttendantIterator](#rangeattendantiterator) | Get a column iterator. |
| [eachRow(callback)](#rangeattendanteachrow) | [RangeAttendant](#rangeattendant) | Execute callback for each row in the range starting from the top and moving down. |
| [eachRowReverse(callback)](#rangeattendanteachrowreverse) | [RangeAttendant](#rangeattendant) | Execute callback for each row in the range starting from the bottom and moving up. |
| [eachColumn(callback)](#rangeattendanteachcolumn) | [RangeAttendant](#rangeattendant) | Execute callback for each column in the range starting from the left and moving right. |
| [eachColumnReverse(callback)](#rangeattendanteachcolumnreverse) | [RangeAttendant](#rangeattendant) | Execute callback for each column in the range starting from the right and moving left. |

### RangeAttendant.includeAllColumns()

Expand the range to include all the columns of the rows in the range.

```javascript
  var cell = SpreadsheetAppAttendant.getActive().getRange('A1');
  var firstRow = cell.includeAllColumns();
  LoggerAttendant.info(firstRow.getA1Notation());
```

##### Return
[RangeAttendant](#rangeattendant)

### RangeAttendant.isBlank()

Sometimes Range.isBlank() throws errors. This provides a backup implementation.

```javascript
  var cell = SpreadsheetAppAttendant.getActive().getRange('A1');
  LoggerAttendant.info(cell.isBlank());
```

##### Return
Boolean

## RangeAttendantIterator
### Methods

| Method | Return Type | Brief description |
| ------ | ----------- | ----------------- |
| [reverse()](#rangeattendantiteratorreverse) | [RangeAttendantIterator](#rangeattendantiterator) | Iterate over the range in reverse order. |
| [hasNext()](#rangeattendantiteratorhasnext) | Boolean | Returns true if there are more rows or columns left in the range. |
| [getSize()](#rangeattendantiteratorgetsize) | Number | Returns the total number of rows or columns in the range. |
| [startAt()](#rangeattendantiteratorstartat) | [RangeAttendantIterator](#rangeattendantiterator) | Starts at the row or column at the passed in index. |
| [next()](#rangeattendantiteratornext) | [RangeAttendant](#rangeattendant) | Returns the next row or column in the range. |



## TODO
* Add more docs
* Add examples
