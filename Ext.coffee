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

class ExtOverridesFactory
  @override: (object)->
    switch
      when TypeHelper.isRange(object)
        new RangeExt(object)
      when TypeHelper.isSpreadsheet(object)
        new SpreadsheetExt(object)
      when TypeHelper.isSheet(object)
        new SheetExt(object)
      when TypeHelper.isScriptProperties(object)
        new ScriptPropertiesExt(object)
      when TypeHelper.isUserProperties(object)
        new UserPropertiesExt(object)
      when TypeHelper.isDocumentProperties(object)
        new DocumentPropertiesExt(object)
      else
        object

class BaseExt
  constructor: (@object)->

  __noSuchMethod__: (id, args)->
    throw new TypeError unless @object[id]?
    returnObject = @object[id].apply(@object, args)
    ExtOverridesFactory.override(returnObject)

  @staticNoSuchMethodGenerator: (thisArg, wrappedObject)->
    thisArg['__noSuchMethod__'] = (id, args)->
      throw new TypeError unless wrappedObject[id]?
      returnObject = wrappedObject[id].apply(wrappedObject, args)
      ExtOverridesFactory.override(returnObject)

class SheetIterator
  eachRow: (callback)->
    @getEntireRange().eachRow(callback)

  eachRowReverse: (callback)->
    @getEntireRange().eachRowReverse(callback)

  eachColumn: (callback)->
    @getEntireRange().eachColumn(callback)

  eachColumnReverse: (callback)->
    @getEntireRange().eachColumnReverse(callback)

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

  eachDataRowReverse: (callback)->
    @getDataRange().eachRowReverse(callback)

  eachDataColumn: (callback)->
    @getDataRange().eachColumn(callback)

  eachDataColumnReverse: (callback)->
    @getDataRange().eachColumnReverse(callback)

class SheetAppender
  appendRowRuturnRange: (data)->
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

class SpreadsheetExt extends Utilities.mixOf BaseExt, SheetIterator, SheetAppender
  getEntireRange: ->
    sheet = @getActiveSheet()
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())

  toString: ->
    "SpreadsheetExt"


class SheetExt extends Utilities.mixOf BaseExt, SheetIterator, SheetAppender
  getEntireRange: ->
    @getRange(1, 1, @getMaxRows(), @getMaxColumns())

  toString: ->
    "SheetExt"

class RangeExt extends BaseExt
  isBlank: ->
    try
      @object.isBlank()
    catch error
      #LoggerExt.debug('Built in Range.isBlank() failed trying backup')
      values = @includeAllColumns().getValues()
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

  eachRowReverse: (callback)->
    rowIterator = @rowIterator().reverse()
    while rowIterator.hasNext()
      callback(rowIterator.next(), rowIterator.currentIndex)

  eachColumn: (callback)->
    columnIterator = @columnIterator()
    while columnIterator.hasNext()
      callback(columnIterator.next(), columnIterator.currentIndex)

  eachColumnReverse: (callback)->
    columnIterator = @columnIterator().reverse()
    while columnIterator.hasNext()
      callback(columnIterator.next(), columnIterator.currentIndex)

class RangeExtIterator
  constructor: (@range)->
    @currentIndex = 1
    @reversed = false

  reverse: ->
    @currentIndex = @getSize()
    @reversed = true
    @

  hasNext: ->
    if @reversed then @currentIndex > 0 else @currentIndex <= @getSize()

  getSize: ->
  next: ->

  startAt: (index)->
    @currentIndex = index if 0 < index <= @getSize()
    @

class RangeRowIterator extends RangeExtIterator
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

class RangeColumnIterator extends RangeExtIterator
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

class PropertiesExt
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

class ScriptPropertiesExt extends Utilities.mixOf BaseExt, PropertiesExt
class UserPropertiesExt extends Utilities.mixOf BaseExt, PropertiesExt
class DocumentPropertiesExt extends Utilities.mixOf BaseExt, PropertiesExt

class PropertiesServiceExt
  @__noSuchMethod__: (id, args)->
    throw new TypeError unless PropertiesService[id]?
    returnObject = PropertiesService[id].apply(PropertiesService, args)
    ExtOverridesFactory.override(returnObject)

class SpreadsheetAppExt

  @DataValidationCriteria = SpreadsheetApp.DataValidationCriteria

  @__noSuchMethod__: (id, args)->
    throw new TypeError unless SpreadsheetApp[id]?
    returnObject = SpreadsheetApp[id].apply(SpreadsheetApp, args)
    ExtOverridesFactory.override(returnObject)

class LoggerExt
  @SEVERITY =
    UNKNOWN: 5
    FATAL: 4
    ERROR: 3
    WARN: 2
    INFO: 1
    DEBUG: 0

  level = LoggerExt.SEVERITY.WARN

  @getLevel: ->
    level

  @setLevel: (value)->
    level = value if LoggerExt.SEVERITY.DEBUG <= value <= LoggerExt.SEVERITY.UNKNOWN

  @isDebug: ->
    LoggerExt.getLevel() <= LoggerExt.SEVERITY.DEBUG

  @isInfo: ->
    LoggerExt.getLevel() <= LoggerExt.SEVERITY.INFO

  @isWarn: ->
    LoggerExt.getLevel() <= LoggerExt.SEVERITY.WARN

  @isError: ->
    LoggerExt.getLevel() <= LoggerExt.SEVERITY.ERROR

  @isFatal: ->
    LoggerExt.getLevel() <= LoggerExt.SEVERITY.FATAL

  @log: (severity = LoggerExt.SEVERITY.UNKNOWN, message = '', args...)->
    return if severity < LoggerExt.getLevel()
    formattedMessage = LoggerExt.formatMessage(severity, message)
    Logger.log(formattedMessage, args...)

  @formatMessage: (severity, message)->
    formattedLevel = switch severity
      when LoggerExt.SEVERITY.DEBUG then 'DEBUG'
      when LoggerExt.SEVERITY.INFO then 'INFO'
      when LoggerExt.SEVERITY.WARN then 'WARN'
      when LoggerExt.SEVERITY.ERROR then 'ERROR'
      when LoggerExt.SEVERITY.FATAL then 'FATAL'
      else 'UNKNOWN'
    "#{formattedLevel}: #{message}"

  @debug: (message, args...)->
    LoggerExt.log(LoggerExt.SEVERITY.DEBUG, message, args...)

  @info: (message, args...)->
    LoggerExt.log(LoggerExt.SEVERITY.INFO, message, args...)

  @warn: (message, args...)->
    LoggerExt.log(LoggerExt.SEVERITY.WARN, message, args...)

  @error: (message, args...)->
    LoggerExt.log(LoggerExt.SEVERITY.ERROR, message, args...)

  @fatal: (message, args...)->
    LoggerExt.log(LoggerExt.SEVERITY.FATAL, message, args...)

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

root.SpreadsheetAppExt = SpreadsheetAppExt
root.PropertiesServiceExt = PropertiesServiceExt
root.LoggerExt = LoggerExt
