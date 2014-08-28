root = exports ? this

class AttendantUtilities
  @merge: (left, right)->
    unless left?
      left = right
      return left
    for property of right
      if AttendantUtilities.type(right[property]) is 'object'
        if AttendantUtilities.type(left[property]) is 'object'
          left[property] = AttendantUtilities.merge(left[property], right[property]);
        else
          left[property] = right[property];
      else
        left[property] = right[property];
    left

  @reverseMerge: (left, right)->
    AttendantUtilities.merge(right, left)

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

class AttendantAdapter
  @override: (object)->
    return object unless object?
    switch object.toString()
      when 'Range'
        new RangeAttendant(object)
      when 'Spreadsheet'
        new SpreadsheetAttendant(object)
      when 'Sheet'
        new SheetAttendant(object)
      when 'ScriptProperties'
        new ScriptPropertiesAttendant(object)
      when 'UserProperties'
        new UserPropertiesAttendant(object)
      when 'DocumentProperties'
        new DocumentPropertiesAttendant(object)
      else
        object

  @proxyMethod: (object, method, args)->
    throw new TypeError unless object[method]?
    castArguments = []
    for arg in args
      if AttendantUtilities.type(arg) is 'object' and '_baseObject' of arg
        castArguments.push arg['_baseObject']
      else
        castArguments.push arg
    returnObject = object[method].apply(object, castArguments)

    AttendantAdapter.override(returnObject)


class BaseAttendant
  constructor: (@_baseObject)->

  __noSuchMethod__: (id, args)->
    AttendantAdapter.proxyMethod(@_baseObject, id, args)

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

class SpreadsheetAttendant extends AttendantUtilities.mixOf BaseAttendant, SheetIterator, SheetAppender
  getEntireRange: ->
    sheet = @getActiveSheet()
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())

  toString: ->
    "SpreadsheetAttendant"


class SheetAttendant extends AttendantUtilities.mixOf BaseAttendant, SheetIterator, SheetAppender
  getEntireRange: ->
    @getRange(1, 1, @getMaxRows(), @getMaxColumns())

  toString: ->
    "SheetAttendant"

class RangeAttendant extends BaseAttendant
  isBlank: ->
    try
      @_baseObject.isBlank()
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
    AttendantUtilities.fetchDeep(property, fields...)

  setJSONProperty: (key, value)->
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(value))
    @

  mergePropertyMapping: (key, map)->
    mapping = @getJSONProperty(key)
    if mapping?
      AttendantUtilities.merge(mapping, map)
      @setJSONProperty(key, mapping)
    else
      @setJSONProperty(key, map)
    @

class ScriptPropertiesAttendant extends AttendantUtilities.mixOf BaseAttendant, PropertiesAttendant
class UserPropertiesAttendant extends AttendantUtilities.mixOf BaseAttendant, PropertiesAttendant
class DocumentPropertiesAttendant extends AttendantUtilities.mixOf BaseAttendant, PropertiesAttendant

class PropertiesServiceAttendant
  @__noSuchMethod__: (id, args)->
    AttendantAdapter.proxyMethod(PropertiesService, id, args)

class SpreadsheetAppAttendant

  @DataValidationCriteria = SpreadsheetApp.DataValidationCriteria

  @__noSuchMethod__: (id, args)->
    AttendantAdapter.proxyMethod(SpreadsheetApp, id, args)

class Enum
  @_size: 0
  @_VALUES: {}

  @values: () ->
    values = new Array()
    for value in Object.keys(@_VALUES)
      values.push @_VALUES[value]
    values

  @valueOf: (name) ->
    @_VALUES[name]

  _name: undefined
  _ordinal: undefined

  constructor: () ->
    Class = @getSuperclass()
    @_name = Object.keys(Class._VALUES)[Class._size]
    @_ordinal = Class._size
    Class._size += 1
    Class._VALUES[@_name] = this

  name: () ->
    @_name

  ordinal: () ->
    @_ordinal

  compareTo: (other) ->
    @_ordinal - other._ordinal

  equals: (other) ->
    this == other

  toString: () ->
    @_name

  getClass: () ->
    this.constructor

  getSuperclass: () ->
    @getClass().__super__.constructor

class Severity extends Enum
  @_VALUES = {@DEBUG, @INFO, @WARN, @ERROR, @FATAL}

  @DEBUG = new (class extends Severity)
  @INFO = new (class extends Severity)
  @WARN = new (class extends Severity)
  @ERROR = new (class extends Severity)
  @FATAL = new (class extends Severity)

class LoggerAttendant
  @SEVERITY = Severity

  _level = LoggerAttendant.SEVERITY.INFO

  @getLevel: ->
    _level

  @setLevel: (value)->
    _level = value if LoggerAttendant.SEVERITY.valueOf(value)?
    @

  @isDebug: ->
    LoggerAttendant.level.compareTo(LoggerAttendant.SEVERITY.DEBUG) <= 0

  @isInfo: ->
    LoggerAttendant.level.compareTo(LoggerAttendant.SEVERITY.INFO) <= 0

  @isWarn: ->
    LoggerAttendant.level.compareTo(LoggerAttendant.SEVERITY.WARN) <= 0

  @isError: ->
    LoggerAttendant.level.compareTo(LoggerAttendant.SEVERITY.ERROR) <= 0

  @isFatal: ->
    LoggerAttendant.level.compareTo(LoggerAttendant.SEVERITY.FATAL) <= 0

  @_log: (severity = LoggerAttendant.SEVERITY.FATAL, message = '', args...)->
    return @ unless LoggerAttendant.level.compareTo(severity) <= 0
    formattedMessage = LoggerAttendant._formatMessage(severity, message)
    Logger.log(formattedMessage, args...)
    @

  @_formatMessage: (severity, message)->
    "#{severity.toString()}: #{message}"

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

  @__noSuchMethod__: (id, args)->
    AttendantAdapter.proxyMethod(Logger, id, args)

Object.defineProperty(LoggerAttendant, 'level', { get: LoggerAttendant.getLevel, set: LoggerAttendant.setLevel })

root.SpreadsheetAppAttendant = SpreadsheetAppAttendant
root.PropertiesServiceAttendant = PropertiesServiceAttendant
root.LoggerAttendant = LoggerAttendant
