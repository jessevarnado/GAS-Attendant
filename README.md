# Google Apps Script Attendant

This library adds some missing functionality to Google Apps Script services. The library is written in Coffeescript and compiled to Javascript. Please use the Javascript version of the library. :) 

## To use:


Copy the contents of the compiled javascript file into a new file in your project. 
 
Replace the top level Google App Script service objects with their Attendant equivalent. Then use the services as you normally would, except now you can take advantage of the extras that the service attendants provide.

```
SpreadsheetApp.getActiveSpreadsheet()
```

becomes

```
SpreadsheetAppAttendant.getActiveSpreadsheet()
```

### Note: You only have to replace the top level service object. The subsequent objects will automatically be wrapped with their Attendant equivalents.
 
## Current Service Attendants

* SpreadsheetAppAttendant
* [#SpreadsheetAttendant]
* [#SheetAttendant]
* RangeAttendant
* PropertiesServiceAttendant
* ScriptPropertiesAttendant
* UserPropertiesAttendant
* DocumentPropertiesAttendant
* LoggerAttendant


## SpreadsheetAttendant 
## SheetAttendant
### Methods

| Method | Return Type | Brief description |
| ------ | ----------- | ----------------- |
| [getEntireRange()](#getEntireRange) | RangeAttendant | Get a range that contains all the columns and rows of a sheet |


### getEntireRange

Get a range that contains all the columns and rows of a sheet

```javascript
  var range = SpreadsheetAppAttendant.getSpreadsheet.getEntireRange();
```

**Return**
(RangeAttendant)

## RangeAttendant


