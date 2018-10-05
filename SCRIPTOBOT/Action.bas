Attribute VB_Name = "Action"
Option Explicit
Option Private Module

Private Function ExecuteCommand(ByRef object As Object, ByRef strAction As String, ByRef parameters As Variant)
ExecuteCommand = VBA.CallByName(object, strAction, VbMethod, parameters)
End Function

Private Sub PrintResult(result As Long, command As String)
Debug.Print "Executed command: " & command & "; Result: " & result
End Sub

Private Sub CheckResult(result As Long, command As String, ByVal rowNumber As Integer)
rowNumber = rowNumber + 4 'Offset the first table's row
If result = 0 Then
    MsgBox "The following command failed: " & command & " on row: " & rowNumber, vbCritical, "System"
    End
End If
End Sub

Private Sub ResetTableBodyColor(tbl As ListObject): tbl.DataBodyRange.Interior.ColorIndex = xlColorIndexNone: End Sub

Private Sub FillTableBodyRowColor(tbl As ListObject, rw As Long): tbl.DataBodyRange.Rows(rw).Interior.Color = RGB(217, 217, 217): End Sub

Public Sub Execution()
Dim wb As Workbook: Set wb = ThisWorkbook
Dim wsSettings As Worksheet: Set wsSettings = wb.Sheets(1)
Dim wsData As Worksheet: Set wsData = wb.Sheets(2)
Dim tbl As ListObject: Set tbl = wsSettings.ListObjects(1)
Dim rawDataTbl As ListObject: Set rawDataTbl = wsData.ListObjects(1) 'Ignore table's name

Dim rw As Object
For Each rw In tbl.DataBodyRange.Rows
    Dim cl As Object
    For Each cl In rw.Columns
        If cl = "" Then
            MsgBox "Please fill all table cells", vbCritical, "Empty field"
            Exit Sub
        End If
    Next cl
Next rw

Dim command As New cCommand

Dim i As Long
For i = 1 To rawDataTbl.DataBodyRange.Rows.count
    ResetTableBodyColor tbl
    
    With tbl.DataBodyRange
        Dim y As Long
        For y = 1 To .Rows.count
            Dim inputMethod As String: inputMethod = .Rows(y).Columns(1).value
            Dim strAction As String: strAction = .Rows(y).Columns(2).value
            Dim strValues As String: strValues = .Rows(y).Columns(3).value
            
            Dim parameters As Variant
            
            With rawDataTbl
                Dim a As Long
                For a = 1 To .HeaderRowRange.Columns.count
                    Dim colName As String: colName = .HeaderRowRange.Columns(a)
                    Dim strToReplaceWith As String: strToReplaceWith = .DataBodyRange(a)(i)
                    
                    strValues = ReplaceWithValues(strValues, "[" & colName & "]", strToReplaceWith)
                Next a
            End With
            
            If GetKeyState(VK_SHIFT) < 0 Then TerminateExecution
            
            parameters = SplitInput(strValues)
          
            Dim action As String: action = command.MethodsLibrary(strAction)
            Dim output As Long: output = ExecuteCommand(command, action, parameters)

            PrintResult output, action
            CheckResult output, action, y

            FillTableBodyRowColor tbl, y
        Next y
    End With
Next i

MsgBox "End!", vbInformation, "System"
End Sub

Public Sub TerminateExecution(): MsgBox "Execution terminated manually!", vbCritical, "System": End: End Sub
