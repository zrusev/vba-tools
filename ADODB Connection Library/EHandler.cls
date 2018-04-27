VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlErrorID As Long
Private msProcedureName As String
Private msModuleName As String

Public Property Let ModuleName(ByVal sModuleName As String): msModuleName = sModuleName: End Property
Private Property Get ModuleName() As String: ModuleName = msModuleName: End Property
Public Property Let ProcedureName(ByVal sProcedureName As String): msProcedureName = sProcedureName: End Property
Private Property Let ErrorID(ByVal lErrorID As Long): mlErrorID = lErrorID: End Property
Private Property Get ErrorID() As Long: ErrorID = mlErrorID: End Property
Private Property Get ProcedureName() As String: ProcedureName = msProcedureName: End Property
Private Property Get GetLocation() As String: GetLocation = "[" & ThisWorkbook.Name & "].[" & ModuleName & "].[" & ProcedureName & "]": End Property
Private Property Get ErrDescription() As String: ErrDescription = "[" & "Error " & Err.Number & ": " & Err.Description & "]": End Property

Private Sub Class_Terminate()

If Err.Number <> 0 Then
    'ThisWorkbook.Sheets("Errors").Range("A100000").End(xlUp).Offset(1, 0).Value = "[" & Format(Now, "dd mmm yy hh:mm:ss") & "] - " & GetLocation & " - " & ErrDescription
End If

End Sub

Public Sub StopExecution(errNumber As CustomErrors, strMsg As String, strTitle As String)

Err.Raise errNumber, strTitle, strMsg

End Sub

Public Sub DisplayError(Optional msgDescription As String, Optional msgTitle As String)

Dim messageDescription As String
Dim messageTitle As String
Dim messageErrorNumber As Long

If msgDescription <> "" Then messageDescription = msgDescription Else messageDescription = Err.Description
If msgTitle <> "" Then messageTitle = msgTitle Else messageTitle = Err.Source
If Err.Number >= 262144 Then messageErrorNumber = Err.Number Else messageErrorNumber = Err.Number - vbObjectError 'Above 262144 generates Overflow error

'Handle unexpected errors
If CustomErrors.[_First] <= messageErrorNumber And messageErrorNumber <= CustomErrors.[_Last] Then
    'Already caught within the modules
Else
    MsgBox "Internal error." & vbNewLine & messageDescription, vbCritical, messageTitle
End If

End Sub
