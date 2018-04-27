Attribute VB_Name = "z_DB_Functions"
Option Private Module

Private Const ModuleName As String = "z_DB_Functions"

Function CreateNewTableOrGetTheFirstOne(sheetName As String, TableName As String) As String
Dim newTableName As String

    If TableName <> "" Then
        CreateNewTableOrGetTheFirstOne = TableName
        Exit Function
    End If

    newTableName = "Table" & (ThisWorkbook.Sheets(sheetName).ListObjects.Count + 1)
    ThisWorkbook.Sheets(sheetName).ListObjects.Add(xlSrcRange, GetNewTableRange(sheetName), , xlYes).Name = newTableName
    CreateNewTableOrGetTheFirstOne = ThisWorkbook.Sheets(sheetName).ListObjects(newTableName).Name
    
End Function

Function GetNewTableRange(sheetName As String) As Range
Dim startPoint As Integer
Dim endPoint As Integer

    startPoint = ThisWorkbook.Sheets(sheetName).Range("ZZ1").End(xlToLeft).Offset(0, 2).Column
    If startPoint = 3 Then startPoint = 1
    endPoint = ThisWorkbook.Sheets(sheetName).Range("ZZ1").End(xlToLeft).Offset(0, 12).Column
    Set GetNewTableRange = ThisWorkbook.Sheets(sheetName).Range(Cells(1, startPoint), Cells(1, endPoint))
    
End Function

Function GetTableFirstCell(sheetName As String, sheetTable As String) As String

GetTableFirstCell = ThisWorkbook.Sheets(sheetName).ListObjects(sheetTable).HeaderRowRange.Columns(1).Offset(1, 0).Address

End Function

Function AddNewSheetOrGetCurrentName(sheetName As String) As String
Dim ws As Worksheet

    If sheetName <> "" Then
       AddNewSheetOrGetCurrentName = sheetName
       Exit Function
    End If
    
    With ThisWorkbook
        Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
    End With
    
    AddNewSheetOrGetCurrentName = ws.Name
    
End Function

Function GetQueryType(queryType As Integer) As Integer

    Select Case queryType
           Case 1
                GetQueryType = AdoCommandTypes.adCmdText
           Case 2
                GetQueryType = AdoCommandTypes.adCmdStoredProc
           Case 3
                GetQueryType = AdoCommandTypes.adCmdTable
           Case 4
                GetQueryType = AdoCommandTypes.adCmdText
    End Select
    
End Function

Function GetCustomOperatorKey(IsVisibleProperty As Integer) As String

    Select Case IsVisibleProperty
           Case -1
                GetCustomOperatorKey = "Yes"
           Case 0
                GetCustomOperatorKey = "No"
    End Select
    
End Function

Function GetCustomOperatorValue(IsVisibleProperty As String) As Integer

    Select Case LCase(IsVisibleProperty)
           Case LCase("Yes")
                GetCustomOperatorValue = CustomOperators.Yes
           Case LCase("No")
                GetCustomOperatorValue = CustomOperators.No
    End Select
    
End Function

Function GetParameterType(typeName As String) As Integer

    Select Case LCase(typeName)
           Case LCase("Boolean")          '16-bit True or False value type
                GetParameterType = ParameterTypes.adBoolean
           Case LCase("Byte")             '8-bit binary value type
           Case LCase("Char")             '16-bit character value type
           Case LCase("Date")             '64-bit date and time value type
                GetParameterType = ParameterTypes.adDBTimeStamp
           Case LCase("DBNull")           'Reference type indicating missing or nonexistent data
           Case LCase("Null")             'Manually added
                GetParameterType = ParameterTypes.adVarWChar
           Case LCase("Decimal")          '128-bit fixed-point numeric value type
                GetParameterType = ParameterTypes.adDecimal
           Case LCase("Double")           '64-bit floating-point numeric value type
                GetParameterType = ParameterTypes.adDouble
           Case LCase("Integer")          '32-bit integer value type
                GetParameterType = ParameterTypes.adInteger
           Case LCase("Object")           'Reference type pointing to an unspecialized object
           Case LCase("objectclass")      'Reference type pointing to a specialized object created from class objectclass
           Case LCase("Long")             '64-bit integer value type
           Case LCase("Nothing")          'Reference type with no object currently assigned to it
                GetParameterType = ParameterTypes.adEmpty
           Case LCase("SByte")            '8-bit signed integer value type
           Case LCase("Short")            '16-bit integer value type
           Case LCase("Single")           '32-bit floating-point numeric value type
                GetParameterType = ParameterTypes.adSingle
           Case LCase("String")           'Reference type pointing to a string of 16-bit characters
                GetParameterType = ParameterTypes.adVarWChar
           Case LCase("UInteger")         '32-bit unsigned integer value type
           Case LCase("ULong")            '64-bit unsigned integer value type
           Case LCase("UShort")           '16-bit unsigned integer value type
    End Select
    
End Function

Function Wait(ByVal sec As Integer)

timeStamp = Now + TimeSerial(0, 0, sec)
Do
  DoEvents
Loop Until Now > timeStamp

End Function

Sub Accelerate(Optional ByVal boolScreen As Boolean = True, _
               Optional ByVal boolEvents As Boolean = True, _
               Optional ByVal calculationType As CalculationTypes = xlCalculationAutomatic, _
               Optional ByVal cursorType As CursorTypes = -4143)

Application.ScreenUpdating = boolScreen
Application.EnableEvents = boolEvents
Application.Calculation = calculationType
Application.Cursor = cursorType

End Sub
