VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Property Get sheetName() As String:: End Property

Property Let sheetName(ByVal sName As String):: End Property

Property Get TableName() As String:: End Property

Property Let TableName(ByVal tName As String):: End Property

Sub SetObject(shtName As String, tblName As String): End Sub

Sub ClearTableContent(): End Sub

Sub ResizeTable(ByRef adoRecSet As ADODB.Recordset): End Sub

Sub RefreshTable(): End Sub

Sub GetDataFromRecordSet(ByRef adoRecSet As ADODB.Recordset): End Sub

Sub FormatTable(): End Sub
