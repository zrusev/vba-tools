VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoDBTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITable

Private shtName As String
Private tblName As String
Private firstTableBodyCell As String
Private tbl As Excel.ListObject

Private Property Get ITable_SheetName() As String: ITable_SheetName = shtName: End Property
Private Property Let ITable_SheetName(ByVal RHS As String): shtName = RHS: End Property
Private Property Get ITable_TableName() As String: ITable_TableName = tblName: End Property
Private Property Let ITable_TableName(ByVal RHS As String): tblName = RHS: End Property
Private Property Get firstCell() As String: firstCell = firstTableBodyCell: End Property
Private Property Let firstCell(ByVal RHS As String): firstTableBodyCell = RHS: End Property
Public Property Get Table() As Excel.ListObject: Set Table = tbl: End Property
Private Property Let Table(ByVal t As Excel.ListObject): Set tbl = t: End Property

Private Sub ITable_ClearTableContent()

With Table
    If Not .DataBodyRange Is Nothing Then
       .DataBodyRange.Resize(.DataBodyRange.Rows.Count, .DataBodyRange.Columns.Count).Rows.Delete
    End If
End With

End Sub

Private Sub ITable_ResizeTable(ByRef adoRecSet As ADODB.Recordset)
Dim i As Integer

With Table
    .Resize .Range.Resize(, adoRecSet.Fields.Count)
    For i = 1 To adoRecSet.Fields.Count
        .HeaderRowRange(i) = adoRecSet.Fields(i - 1).Name
    Next i
End With

firstCell = Table.HeaderRowRange.Columns(1).Offset(1, 0).Address

End Sub

Private Sub ITable_RefreshTable(): End Sub

Private Sub ITable_GetDataFromRecordSet(ByRef adoRecSet As ADODB.Recordset)

ThisWorkbook.Sheets(ITable_SheetName).Range(firstCell).CopyFromRecordset adoRecSet

End Sub

Public Sub ITable_SetObject(shtName As String, tblName As String)

ITable_SheetName = shtName
ITable_TableName = tblName

Table = ThisWorkbook.Sheets(ITable_SheetName).ListObjects(ITable_TableName)

End Sub

Private Sub ITable_FormatTable()

With Table.DataBodyRange
    .Rows.Interior.TintAndShade = -0.149998474074526
    .Cells.Locked = False
    .Columns.HorizontalAlignment = xlLeft
End With

End Sub

