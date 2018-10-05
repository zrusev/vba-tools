Attribute VB_Name = "Functions"
Option Explicit
'Option Private Module 'Public because of the shortcuts

Private Const Delimiter = "|"
Private Const GW_HWNDNEXT = 2

Public Function Wait(ByVal sec As Integer)
Dim timeStamp As Date
timeStamp = Now + TimeSerial(0, 0, sec)

Do
  DoEvents
Loop Until Now > timeStamp
End Function

Public Function TextToConstant(choice As String)
Select Case choice
    Case "vbKeyLButton": TextToConstant = KeyCodeConstants.vbKeyLButton
    Case "vbKeyRButton": TextToConstant = KeyCodeConstants.vbKeyRButton
    Case "vbKeyCancel": TextToConstant = KeyCodeConstants.vbKeyCancel
    Case "vbKeyMButton": TextToConstant = KeyCodeConstants.vbKeyMButton
    Case "vbKeyMButton": TextToConstant = KeyCodeConstants.vbKeyMButton
    Case "vbKeyBack": TextToConstant = KeyCodeConstants.vbKeyBack
    Case "vbKeyTab": TextToConstant = KeyCodeConstants.vbKeyTab
    Case "vbKeyClear": TextToConstant = KeyCodeConstants.vbKeyClear
    Case "vbKeyReturn": TextToConstant = KeyCodeConstants.vbKeyReturn
    Case "vbKeyShift": TextToConstant = KeyCodeConstants.vbKeyShift
    Case "vbKeyControl": TextToConstant = KeyCodeConstants.vbKeyControl
    Case "vbKeyMenu": TextToConstant = KeyCodeConstants.vbKeyMenu
    Case "vbKeyPause": TextToConstant = KeyCodeConstants.vbKeyPause
    Case "vbKeyCapital": TextToConstant = KeyCodeConstants.vbKeyCapital
    Case "vbKeyEscape": TextToConstant = KeyCodeConstants.vbKeyEscape
    Case "vbKeySpace": TextToConstant = KeyCodeConstants.vbKeySpace
    Case "vbKeyPageUp": TextToConstant = KeyCodeConstants.vbKeyPageUp
    Case "vbKeyPageDown": TextToConstant = KeyCodeConstants.vbKeyPageDown
    Case "vbKeyEnd": TextToConstant = KeyCodeConstants.vbKeyEnd
    Case "vbKeyHome": TextToConstant = KeyCodeConstants.vbKeyHome
    Case "vbKeyLeft": TextToConstant = KeyCodeConstants.vbKeyLeft
    Case "vbKeyUp": TextToConstant = KeyCodeConstants.vbKeyUp
    Case "vbKeyRight": TextToConstant = KeyCodeConstants.vbKeyRight
    Case "vbKeyDown": TextToConstant = KeyCodeConstants.vbKeyDown
    Case "vbKeySelect": TextToConstant = KeyCodeConstants.vbKeySelect
    Case "vbKeyPrint": TextToConstant = KeyCodeConstants.vbKeyPrint
    Case "vbKeyExecute": TextToConstant = KeyCodeConstants.vbKeyExecute
    Case "vbKeySnapshot": TextToConstant = KeyCodeConstants.vbKeySnapshot
    Case "vbKeyInsert": TextToConstant = KeyCodeConstants.vbKeyInsert
    Case "vbKeyDelete": TextToConstant = KeyCodeConstants.vbKeyDelete
    Case "vbKeyHelp": TextToConstant = KeyCodeConstants.vbKeyHelp
    Case "vbKeyNumlock": TextToConstant = KeyCodeConstants.vbKeyNumlock
End Select
End Function

Public Function SplitInput(value As String) As Variant: SplitInput = Split(value, Delimiter, , vbTextCompare): End Function

Public Function ReplaceWithValues(strValue, strToSearchFor, strToReplaceWith) As String: ReplaceWithValues = Replace(strValue, strToSearchFor, strToReplaceWith, 1, , vbTextCompare): End Function

Public Sub GetCursorPosition()
Attribute GetCursorPosition.VB_ProcData.VB_Invoke_Func = "Q\n14"
Dim pos As POINTAPI
GetCursorPos pos

Dim wb As Workbook: Set wb = ThisWorkbook
Dim ws As Worksheet: Set ws = wb.Sheets("Settings")

ws.Range("F5") = pos.x & "|" & pos.y
End Sub

Public Sub ClearTable()
Dim wb As Workbook: Set wb = ThisWorkbook
Dim ws As Worksheet: Set ws = wb.Sheets("Settings")
Dim rawDataTbl As ListObject: Set rawDataTbl = ws.ListObjects("Table1")

With rawDataTbl
    If Not .DataBodyRange Is Nothing Then
       .DataBodyRange.Resize(.DataBodyRange.Rows.count, .DataBodyRange.Columns.count).Rows.Delete
    End If
End With
End Sub

Public Sub ListWins(Optional Title = "*", Optional Class = "*")
    Dim hWndThis As Long: hWndThis = FindWindowA(vbNullString, vbNullString)
    
    While hWndThis
        Dim sTitle As String, sClass As String
        
        sTitle = Space$(255)
        sTitle = left$(sTitle, GetWindowText(hWndThis, sTitle, Len(sTitle)))
        
        sClass = Space$(255)
        sClass = left$(sClass, GetClassName(hWndThis, sClass, Len(sClass)))
        
        If sTitle Like Title And sClass Like Class Then
            Debug.Print sTitle, sClass
        End If
        hWndThis = GetWindow(hWndThis, GW_HWNDNEXT)
        
        Dim rec As rect
        Dim result As Long: result = GetWindowRect(hWndThis, rec)
    Wend
End Sub

Public Sub ReturnTasks()
Dim wb As Workbook: Set wb = ThisWorkbook
Dim wsSettings As Worksheet: Set wsSettings = wb.Sheets("Settings")
Dim apps() As Variant: apps() = Tasks.EnumWindows

Dim colName As String: colName = "J"
wsSettings.Range(colName & "10000:" & colName & Range(colName & "1").End(xlDown).Offset(1, 0).row).Clear

Dim k As Variant
For Each k In apps()
    wsSettings.Range(colName & "10000").End(xlUp).Offset(1, 0).value = k
Next k

wsSettings.Range(colName & Range(colName & "1").End(xlDown).Offset(1, 0).row & ":" & colName & Range(colName & "10000").End(xlUp).row).Interior.Color = RGB(220, 230, 241)
End Sub

Public Sub PopulateSelectorCoordinates(rec As rect)
Dim wb As Workbook: Set wb = ThisWorkbook
Dim wsSettings As Worksheet: Set wsSettings = wb.Sheets("Settings")

Dim colName As String: colName = "H"
wsSettings.Range(colName & "5").value = rec.left & "|" & rec.top & "|" & rec.right & "|" & rec.bottom
End Sub

Public Sub CallSelector(): UserForm1.Show: End Sub
Attribute CallSelector.VB_ProcData.VB_Invoke_Func = "A\n14"
