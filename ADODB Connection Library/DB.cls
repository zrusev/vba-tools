VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const ModuleName As String = "DB"

Private Const User = "" 'username
Private Const server = "" 'server instance
Private Const Password = "" 'password
Private Const database = "" 'database name
Private Const schema = "" 'schema name

Private Const CONN_TIMEOUT = 15
Private Const COMM_TIMEOUT = 30

Private connDB As ADODB.Connection
Private cmd As ADODB.Command
Private recSet As ADODB.Recordset
Private strQuery As String
Private cmdType As Integer
Private msgBoxMessage As String
Private AdoTable As ITable
Private pvtTable As ITable
Private clsError As EHandler
Private customParameter As AdoDBParameter
Private ignoreEmptyFlag As Boolean
Private singleResultFlag As Boolean
Private sRes As Variant

Public Property Get SingleResult() As Variant: SingleResult = sRes: End Property
Private Property Let SingleResult(ByVal sResult As Variant): sRes = sResult: End Property
Private Property Get ErrMessage() As String: ErrMessage = msgBoxMessage: End Property
Private Property Let ErrMessage(ByVal eMsg As String): msgBoxMessage = eMsg: End Property
Private Property Get Recordset() As ADODB.Recordset: Set Recordset = recSet: End Property
Private Property Let Recordset(ByVal rcdSet As ADODB.Recordset): Set recSet = rcdSet: End Property
Public Property Get Connection() As ADODB.Connection: Set Connection = connDB: End Property
Private Property Let Connection(ByVal adoConn As ADODB.Connection): Set connDB = adoConn: End Property
Private Property Get CommandType() As Integer: CommandType = cmdType: End Property
Public Property Let CommandType(ByVal cmd As Integer): cmdType = GetQueryType(cmd): End Property
Private Property Get Query() As String: Query = strQuery: End Property
Public Property Let Query(ByVal strSQL As String): strQuery = ReplaceDb(strSQL): End Property

Private Property Get IsConnected() As Boolean

If Connection.State = adStateOpen Then IsConnected = True Else IsConnected = False

End Property

Private Property Get IsPivotTableSet() As Boolean

If Not pvtTable Is Nothing Then IsPivotTableSet = True

End Property

Private Property Get ReplaceDb(queryString As String) As String

ReplaceDb = Replace(queryString, "[database]", "[" & database & "]")
ReplaceDb = Replace(ReplaceDb, "[schema]", "[" & schema & "]")

End Property

Private Sub Class_Initialize()

Set cmd = Nothing
Set cmd = New ADODB.Command

End Sub

Private Sub Class_Terminate()

If Not connDB Is Nothing Then
    If connDB.State = adStateOpen Then connDB.Close
End If

Set connDB = Nothing
Set cmd = Nothing

End Sub

Private Sub ExecuteProcedure()
Dim clsError As New EHandler

clsError.ModuleName = ModuleName
clsError.ProcedureName = "ExecuteProcedure"

On Error GoTo errorhandler
cmd.Execute

ErrorExit: Exit Sub

errorhandler:
Set clsError = Nothing
TryReconnect
Execute
GoTo ErrorExit

End Sub

Private Sub TryReconnect()
Static intConnectionAttempts As Integer

Dim clsError As New EHandler

clsError.ModuleName = ModuleName
clsError.ProcedureName = "TryReconnect"

connDB.Close
Wait 2

If intConnectionAttempts >= 0 Then
    Set connDB = Nothing
    intConnectionAttempts = 0
    MsgBox "The server is currently unavailable." & vbNewLine & "Please try again later or contact support.", vbInformation, ErrorSystemTitle
    clsError.StopExecution Connection_Time_Out, Format(Now, "dd mmm yy hh:mm:ss") & " Unable to connect to server. Execution terminated.", "Connection Timed out"
End If

intConnectionAttempts = intConnectionAttempts + 1
'ThisWorkbook.Sheets("Errors").Range("A100000").End(xlUp).Offset(1, 0).Value = "[" & Format(Now, "dd mmm yy hh:mm:ss") & "]" & " - Unable to connect to server. Attempt No. " & intConnectionAttempts + 1 & ". Connecting..."

OpenConnection

End Sub

Private Sub GetDataFromRecordSet()
Dim clsError As New EHandler

clsError.ModuleName = ModuleName
clsError.ProcedureName = "GetDataFromRecordSet"

On Error GoTo errorhandler
Recordset = cmd.Execute

If IsRecordSetEmpty Then
    If ignoreEmptyFlag Then GoTo IgnoreEmptyLabel:
    On Error GoTo 0
    MsgBox ErrMessage & vbNewLine & ErrorMessageTwo, vbInformation, ErrorSystemTitle
    clsError.StopExecution Empty_RecordSet, "Sheet(""" & AdoTable.sheetName & """).Table(""" & AdoTable.TableName & """) cound not be filled. Empty RecordSet.", "Empty Table"
End If

If singleResultFlag Then
    SingleResult = Recordset.Fields(0).Value
    singleResultFlag = False
Else
    With AdoTable
        .ResizeTable Recordset
        .GetDataFromRecordSet Recordset
        .FormatTable
    End With
End If

IgnoreEmptyLabel:
ignoreEmptyFlag = False

Recordset.Close
Recordset = Nothing

If IsPivotTableSet Then pvtTable.RefreshTable
ErrorExit: Exit Sub

errorhandler:
Set clsError = Nothing
TryReconnect
Execute
GoTo ErrorExit

End Sub

Private Sub RemoveAllParameters()
Dim y As Integer

For y = cmd.parameters.Count - 1 To 0 Step -1
    cmd.parameters.Delete (y)
Next y

End Sub

Private Function IsRecordSetEmpty() As Boolean

If (Recordset.BOF = True And Recordset.EOF = True) Then IsRecordSetEmpty = True

End Function

Public Sub IgnoreEmptyRecordSet(): ignoreEmptyFlag = True: End Sub

Public Sub OpenConnection()

Set connDB = Nothing

Connection = New ADODB.Connection
Connection.ConnectionTimeout = CONN_TIMEOUT
Connection.CommandTimeout = COMM_TIMEOUT
Connection.Open "Provider = sqloledb;" & "Data Source=" & server & ";" & "User ID=" & User & ";" & "Password=" & Password & ";"

If Not IsConnected Then TryReconnect

End Sub

Public Sub Execute()

With cmd
    .ActiveConnection = Connection
    .CommandType = CommandType
    .CommandText = Query
    If recSet Is Nothing Then ExecuteProcedure Else GetDataFromRecordSet
End With

End Sub

Public Sub SetParameters(ParamArray arrayOfParameters() As Variant)
Dim i As Integer

RemoveAllParameters
For i = LBound(arrayOfParameters) To UBound(arrayOfParameters)
    Set customParameter = New AdoDBParameter
    cmd.parameters.Append customParameter.Add(arrayOfParameters(i))
    Set customParameter = Nothing
Next i

End Sub

Public Sub SetSingleResult()

Set recSet = Nothing
Set recSet = New ADODB.Recordset

singleResultFlag = True

End Sub

Public Sub SetTable(sheetName As String, TableName As String, Optional ErrorMessage As String)

ErrMessage = ErrorMessage

Set recSet = Nothing
Set recSet = New ADODB.Recordset

Set AdoTable = Nothing
Set AdoTable = New AdoDBTable

With AdoTable
    .SetObject sheetName, TableName
    .ClearTableContent
End With

End Sub

Public Sub SetPivotTable(sheetName As String, PivotTableName As String)

Set pvtTable = Nothing
Set pvtTable = New AdoDBPivotTable

pvtTable.SetObject sheetName, PivotTableName

End Sub
