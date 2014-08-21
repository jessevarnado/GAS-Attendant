root = exports ? this

class TypeHelper
  @isRange: (object)->
    object.toString() is 'Range'
  @isSpreadsheet: (object)->
    object.toString() is 'Spreadsheet'
  @isSheet: (object)->
    object.toString() is 'Sheet'
  @isScriptProperties: (object)->
    object.toString() is 'ScriptProperties'
  @isUserProperties: (object)->
    object.toString() is 'UserProperties'
  @isDocumentProperties: (object)->
    object.toString() is 'DocumentProperties'

class AttendantOverrides
  @override: (object)->
    switch
      when TypeHelper.isRange(object)
        new RangeAttendant(object)
      when TypeHelper.isSpreadsheet(object)
        new SpreadsheetAttendant(object)
      when TypeHelper.isSheet(object)
        new SheetAttendant(object)
      when TypeHelper.isScriptProperties(object)
        new ScriptPropertiesAttendant(object)
      when TypeHelper.isUserProperties(object)
        new UserPropertiesAttendant(object)
      when TypeHelper.isDocumentProperties(object)
        new DocumentPropertiesAttendant(object)
      else
        object

class BaseAttendant
  constructor: (@object)->

  __noSuchMethod__: (id, args)->
    throw new TypeError unless @object[id]?
    returnObject = @object[id].apply(@object, args)
    AttendantOverrides.override(returnObject)

class SheetIterator
  eachRow: (callback)->
    @getEntireRange().eachRow(callback)
    @

  eachRowReverse: (callback)->
    @getEntireRange().eachRowReverse(callback)
    @

  eachColumn: (callback)->
    @getEntireRange().eachColumn(callback)
    @

  eachColumnReverse: (callback)->
    @getEntireRange().eachColumnReverse(callback)
    @

  rowIterator: ->
    @getEntireRange().rowIterator()

  columnIterator: ->
    @getEntireRange().columnIterator()

  dataRowIterator: ->
    @getDataRange().rowIterator()

  dataColumnIterator: ->
    @getDataRange().columnIterator()

  eachDataRow: (callback)->
    @getDataRange().eachRow(callback)
    @

  eachDataRowReverse: (callback)->
    @getDataRange().eachRowReverse(callback)
    @

  eachDataColumn: (callback)->
    @getDataRange().eachColumn(callback)
    @

  eachDataColumnReverse: (callback)->
    @getDataRange().eachColumnReverse(callback)
    @

class SheetAppender
  appendRowReturnRange: (data)->
    sheet = @appendRow(data)
    rowIterator = sheet.rowIterator().reverse()
    while rowIterator.hasNext()
      row = rowIterator.next()
      values = row.getValues()[0]
      finder = (value, index)->
        return true unless data[index]?
        value.valueOf() is data[index].valueOf()
      if values.every finder
        return row
    null

class SpreadsheetAttendant extends Utilities.mixOf BaseAttendant, SheetIterator, SheetAppender
  getEntireRange: ->
    sheet = @getActiveSheet()
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())

  toString: ->
    "SpreadsheetAttendant"


class SheetAttendant extends Utilities.mixOf BaseAttendant, SheetIterator, SheetAppender
  getEntireRange: ->
    @getRange(1, 1, @getMaxRows(), @getMaxColumns())

  toString: ->
    "SheetAttendant"

class RangeAttendant extends BaseAttendant
  isBlank: ->
    try
      @object.isBlank()
    catch error
      LoggerAttendant.debug('Built in Range.isBlank() failed trying backup')
      values = @getValues()
      for row in values
        for value in row
          return false if value isnt ''
      true

  includeAllColumns: ->
    sheet = @getSheet()
    rowIndex = @getCell(1, 1).getRow()
    sheet.getRange(rowIndex, 1, @getNumRows(), sheet.getMaxColumns())

  removeHeader: ->
    if @getRow() is 1
      if @getNumRows() is 1
        null
      else
        @sliceRows(2)
    else
      @

  sliceRows: (start, length)->
    @slice(start, 1, length)

  sliceColumns: (start, length)->
    @slice(1, start, null, length)

  slice: (startRow,
          startColumn,
          rowLength = @getNumRows() - startRow + 1,
          columnLength = @getNumColumns() - startColumn + 1)->
    return @ if startRow > @getNumRows() or startColumn > @getNumColumns()
    rowLength = @getNumRows() - startRow + 1 if rowLength + startRow > @getNumRows()
    columnLength = @getNumColumns() - startColumn + 1 if columnLength + startColumn > @getNumColumns()
    startCell = @getCell(startRow, startColumn)
    @getSheet().getRange(startCell.getRow(), startCell.getColumn(), rowLength, columnLength)

  rowIterator: ->
    new RangeRowIterator(@)

  columnIterator: ->
    new RangeColumnIterator(@)

  eachRow: (callback)->
    rowIterator = @rowIterator()
    while rowIterator.hasNext()
      callback(rowIterator.next(), rowIterator.currentIndex)
    @

  eachRowReverse: (callback)->
    rowIterator = @rowIterator().reverse()
    while rowIterator.hasNext()
      callback(rowIterator.next(), rowIterator.currentIndex)
    @

  eachColumn: (callback)->
    columnIterator = @columnIterator()
    while columnIterator.hasNext()
      callback(columnIterator.next(), columnIterator.currentIndex)
    @

  eachColumnReverse: (callback)->
    columnIterator = @columnIterator().reverse()
    while columnIterator.hasNext()
      callback(columnIterator.next(), columnIterator.currentIndex)
    @

class RangeAttendantIterator
  constructor: (@range)->
    @currentIndex = 1
    @reversed = false

  reverse: ->
    @currentIndex = @getSize()
    @reversed = not @reversed
    @

  hasNext: ->
    if @reversed then @currentIndex > 0 else @currentIndex <= @getSize()

  getSize: ->
  next: ->

  startAt: (index)->
    @currentIndex = index if 0 < index <= @getSize()
    @

class RangeRowIterator extends RangeAttendantIterator
  getSize: ->
    @range.getNumRows()

  next: ->
    sheet = @range.getSheet()
    firstCell = @range.getCell(@currentIndex, 1)
    rowIndex = firstCell.getRow()
    columnIndex = firstCell.getColumn()
    if @reversed
      @currentIndex--
    else
      @currentIndex++
    sheet.getRange(rowIndex, columnIndex, 1, @range.getNumColumns())

class RangeColumnIterator extends RangeAttendantIterator
  getSize: ->
    @range.getNumColumns()

  next: ->
    sheet = @range.getSheet()
    firstCell = @range.getCell(1, @currentIndex)
    rowIndex = firstCell.getRow()
    columnIndex = firstCell.getColumn()
    if @reversed
      @currentIndex--
    else
      @currentIndex++
    sheet.getRange(rowIndex, columnIndex, @range.getNumRows(), 1)

class PropertiesAttendant
  getJSONProperty: (key)->
    JSON.parse(PropertiesService.getScriptProperties().getProperty(key))

  fetchDeepJSONProperty: (key, fields...)->
    property = @getJSONProperty(key)
    Utilities.fetchDeep(property, fields...)

  setJSONProperty: (key, value)->
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(value))
    @

  mergePropertyMapping: (key, map)->
    mapping = @getJSONProperty(key)
    if mapping?
      Utilities.merge(mapping, map)
      @setJSONProperty(key, mapping)
    else
      @setJSONProperty(key, map)
    @

class ScriptPropertiesAttendant extends Utilities.mixOf BaseAttendant, PropertiesAttendant
class UserPropertiesAttendant extends Utilities.mixOf BaseAttendant, PropertiesAttendant
class DocumentPropertiesAttendant extends Utilities.mixOf BaseAttendant, PropertiesAttendant

class PropertiesServiceAttendant
  @__noSuchMethod__: (id, args)->
    throw new TypeError unless PropertiesService[id]?
    returnObject = PropertiesService[id].apply(PropertiesService, args)
    AttendantOverrides.override(returnObject)

class SpreadsheetAppAttendant

  @DataValidationCriteria = SpreadsheetApp.DataValidationCriteria

  @__noSuchMethod__: (id, args)->
    throw new TypeError unless SpreadsheetApp[id]?
    returnObject = SpreadsheetApp[id].apply(SpreadsheetApp, args)
    AttendantOverrides.override(returnObject)

class LoggerAttendant
  @SEVERITY =
    UNKNOWN: 5
    FATAL: 4
    ERROR: 3
    WARN: 2
    INFO: 1
    DEBUG: 0

  level = LoggerAttendant.SEVERITY.INFO

  @getLevel: ->
    level

  @setLevel: (value)->
    level = value if LoggerAttendant.SEVERITY.DEBUG <= value <= LoggerAttendant.SEVERITY.UNKNOWN

  @isDebug: ->
    LoggerAttendant.getLevel() <= LoggerAttendant.SEVERITY.DEBUG

  @isInfo: ->
    LoggerAttendant.getLevel() <= LoggerAttendant.SEVERITY.INFO

  @isWarn: ->
    LoggerAttendant.getLevel() <= LoggerAttendant.SEVERITY.WARN

  @isError: ->
    LoggerAttendant.getLevel() <= LoggerAttendant.SEVERITY.ERROR

  @isFatal: ->
    LoggerAttendant.getLevel() <= LoggerAttendant.SEVERITY.FATAL

  @_log: (severity = LoggerAttendant.SEVERITY.UNKNOWN, message = '', args...)->
    return if severity < LoggerAttendant.getLevel()
    formattedMessage = LoggerAttendant._formatMessage(severity, message)
    Logger.log(formattedMessage, args...)
    @

  @_formatMessage: (severity, message)->
    formattedLevel = switch severity
      when LoggerAttendant.SEVERITY.DEBUG then 'DEBUG'
      when LoggerAttendant.SEVERITY.INFO then 'INFO'
      when LoggerAttendant.SEVERITY.WARN then 'WARN'
      when LoggerAttendant.SEVERITY.ERROR then 'ERROR'
      when LoggerAttendant.SEVERITY.FATAL then 'FATAL'
      else 'UNKNOWN'
    "#{formattedLevel}: #{message}"

  @debug: (message, args...)->
    LoggerAttendant._log(LoggerAttendant.SEVERITY.DEBUG, message, args...)

  @info: (message, args...)->
    LoggerAttendant._log(LoggerAttendant.SEVERITY.INFO, message, args...)

  @warn: (message, args...)->
    LoggerAttendant._log(LoggerAttendant.SEVERITY.WARN, message, args...)

  @error: (message, args...)->
    LoggerAttendant._log(LoggerAttendant.SEVERITY.ERROR, message, args...)

  @fatal: (message, args...)->
    LoggerAttendant._log(LoggerAttendant.SEVERITY.FATAL, message, args...)

class Utilities
  @merge: (left, right)->
    unless left?
      left = right
      return left
    for property of right
      if Utilities.type(right[property]) is 'object'
        if Utilities.type(left[property]) is 'object'
          left[property] = Utilities.merge(left[property], right[property]);
        else
          left[property] = right[property];
      else
        left[property] = right[property];
    left

  @reverseMerge: (left, right)->
    Utilities.merge(right, left)

  @type: (obj) ->
    if obj == undefined or obj == null
      return String obj
    classToType = {
      '[object Boolean]': 'boolean',
      '[object Number]': 'number',
      '[object String]': 'string',
      '[object Function]': 'function',
      '[object Array]': 'array',
      '[object Date]': 'date',
      '[object RegExp]': 'regexp',
      '[object Object]': 'object'
    }
    classToType[Object.prototype.toString.call(obj)]

  @fetchDeep: (obj, fields...)->
    reducer = (prev, curr)->
      prev[curr]
    fields.reduce(reducer, obj)

  @mixOf: (base, mixins...) ->
    class Mixed extends base
    for mixin in mixins by -1 #earlier mixins override later ones
      for name, method of mixin::
        Mixed::[name] = method
    Mixed

root.SpreadsheetAppAttendant = SpreadsheetAppAttendant
root.PropertiesServiceAttendant = PropertiesServiceAttendant
root.LoggerAttendant = LoggerAttendant
