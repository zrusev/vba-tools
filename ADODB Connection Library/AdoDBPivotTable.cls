VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoDBPivotTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITable

Private shtName As String
Private pvtName As String
Private pvt As Excel.PivotTable

Private Property Get ITable_SheetName() As String: ITable_SheetName = shtName: End Property
Private Property Let ITable_SheetName(ByVal RHS As String): shtName = RHS: End Property
Private Property Get ITable_TableName() As String: ITable_TableName = pvtName: End Property
Private Property Let ITable_TableName(ByVal RHS As String): pvtName = RHS: End Property
Private Property Get PivotTable() As Excel.PivotTable: Set PivotTable = pvt: End Property
Private Property Let PivotTable(ByVal pTable As Excel.PivotTable): Set pvt = pTable: End Property
Private Sub ITable_ClearTableContent(): End Sub
Private Sub ITable_ResizeTable(adoRecSet As ADODB.Recordset): End Sub
Private Sub ITable_FormatTable(): End Sub
Private Sub ITable_GetDataFromRecordSet(adoRecSet As ADODB.Recordset): End Sub

Private Sub ITable_RefreshTable()

Application.EnableEvents = False
PivotTable.RefreshTable
Application.EnableEvents = True

End Sub

Private Sub ITable_SetObject(shtName As String, tblName As String)

ITable_SheetName = shtName
ITable_TableName = tblName

PivotTable = ThisWorkbook.Sheets(ITable_SheetName).PivotTables(ITable_TableName)

End Sub

