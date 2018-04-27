Attribute VB_Name = "a_Enums"
Option Private Module

Enum ParameterTypes
     adEmpty = 0             'No value
     adSmallInt = 2          'A 2-byte signed integer.
     adInteger = 3           'A 4-byte signed integer.
     adSingle = 4            'A single-precision floating-point value.
     adDouble = 5            'A double-precision floating-point value.
     adCurrency = 6          'A currency value
     adDate = 7              'The number of days since December 30, 1899 + the fraction of a day.
     adBSTR = 8              'A null-terminated character string.
     adIDispatch = 9         'A pointer to an IDispatch interface on a COM object. Note: Currently not supported by ADO.
     adError = 10            'A 32-bit error code
     adBoolean = 11          'A boolean value.
     adVariant = 12          'An Automation Variant. Note: Currently not supported by ADO.
     adIUnknown = 13         'A pointer to an IUnknown interface on a COM object. Note: Currently not supported by ADO.
     adDecimal = 14          'An exact numeric value with a fixed precision and scale.
     adTinyInt = 16          'A 1-byte signed integer.
     adUnsignedTinyInt = 17  'A 1-byte unsigned integer.
     adUnsignedSmallInt = 18 'A 2-byte unsigned integer.
     adUnsignedInt = 19      'A 4-byte unsigned integer.
     adBigInt = 20           'An 8-byte signed integer.
     adUnsignedBigInt = 21   'An 8-byte unsigned integer.
     adFileTime = 64         'The number of 100-nanosecond intervals since January 1,1601
     adGUID = 72             'A globally unique identifier (GUID)
     adBinary = 128          'A binary value.
     adChar = 129            'A string value.
     adWChar = 130           'A null-terminated Unicode character string.
     adNumeric = 131         'An exact numeric value with a fixed precision and scale.
     adUserDefined = 132     'A user-defined variable.
     adDBDate = 133          'A date value (yyyymmdd).
     adDBTime = 134          'A time value (hhmmss).
     adDBTimeStamp = 135     'A date/time stamp (yyyymmddhhmmss plus a fraction in billionths).
     adChapter = 136         'A 4-byte chapter value that identifies rows in a child rowset
     adPropVariant = 138     'An Automation PROPVARIANT.
     adVarNumeric = 139      'A numeric value (Parameter object only).
     adVarChar = 200         'A string value (Parameter object only).
     adLongVarChar = 201     'A long string value.
     adVarWChar = 202        'A null-terminated Unicode character string.
     adLongVarWChar = 203    'A long null-terminated Unicode string value.
     adVarBinary = 204       'A binary value (Parameter object only).
     adLongVarBinary = 205   'A long binary value.
     'adArray = 0x2000       'A flag value combined with another data type constant. Indicates an array of that other data type.
End Enum

Enum AdoCommandTypes
     adCmdUnspecified = -1   'Does not specify the command type argument.
     adCmdText = 1           'Evaluates CommandText as a textual definition of a command or stored procedure call.
     adCmdTable = 2          'Evaluates CommandText as a table name whose columns are all returned by an internally generated SQL query.
     adCmdStoredProc = 4     'Evaluates CommandText as a stored procedure name.
     adCmdUnknown = 8        'Default. Indicates that the type of command in the CommandText property is not known.
     adCmdFile = 256         'Evaluates CommandText as the file name of a persistently stored Recordset. Used with Recordset.Open or Requery only.
     adCmdTableDirect = 512  'Evaluates CommandText as a table name whose columns are all returned. Used with Recordset.Open or Requery only.
                             'To use the Seek method, the Recordset must be opened with adCmdTableDirect. This value cannot be combined with the ExecuteOptionEnum value adAsyncExecute.
End Enum

Enum ParameterDirections
     adParamUnknown = 0      'Direction is unknown
     adParamInput = 1        'Input parameter
     adParamOutput = 2       'Output parameter
     adParamInputOutput = 3  'Both input and output parameter
     adParamReturnValue = 4  'Return value
End Enum

Enum QueryTypes
     CRUD = 1
     StoredProcedure = 2
     TableValuedFunction = 3
     ScalarValuedFunction = 4
End Enum

Enum CustomOperators
     Yes = True
     No = False
     Not_Applicable = True
End Enum

Enum CursorTypes
     xlDefault = -4143
     xlIBeam = 3
     xlNorthwestArrow = 1
     xlWait = 2
End Enum

Enum CalculationTypes
    xlCalculationAutomatic = -4105
    xlCalculationManual = -4135
    xlCalculationSemiautomatic = 2
End Enum

Enum CustomErrors
    [_First] = vbObjectError + 512
    Connection_Time_Out = vbObjectError + 513
    Empty_RecordSet = vbObjectError + 514
    Limit_Access = vbObjectError + 515
    [_Last] = vbObjectError + 516
End Enum
