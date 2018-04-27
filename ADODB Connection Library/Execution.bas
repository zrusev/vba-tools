Attribute VB_Name = "Execution"
Option Private Module
Option Explicit

Private Const ModuleName As String = "Execution"

Sub Test_Query()
Dim clsError As New EHandler

clsError.ModuleName = ModuleName
clsError.ProcedureName = "ProcedureName"

On Error GoTo errorhandler

Dim DB As New DB

With DB
    .OpenConnection
    .CommandType = QueryTypes.CRUD
    .Query = "SELECT ID_Login FROM [database].[schema].[tbl_Login] WHERE UserName = ?"
    .SetParameters "zrusev"
    .SetSingleResult
    '.SetTable "Sheet1", "Table1", "User has not been created yet."
    .Execute
End With

Debug.Print DB.SingleResult

'Test a stored procedure and wrap the error around the caller
Call Test_StoredProcedure

ErrorExit: Exit Sub

errorhandler:
Set clsError = Nothing

Accelerate True, True, xlCalculationAutomatic, xlDefault
clsError.DisplayError
Resume ErrorExit

End Sub

Private Sub Test_StoredProcedure()
Dim clsError As New EHandler

clsError.ModuleName = ModuleName
clsError.ProcedureName = "ProcedureName"

Dim DB As New DB

With DB
    If DB.Connection Is Nothing Then .OpenConnection
    .CommandType = QueryTypes.StoredProcedure
    .Query = "[schema].StoredProcedureName"
    .SetParameters "if any"
    .Execute
End With

End Sub

