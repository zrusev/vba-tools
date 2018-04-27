Attribute VB_Name = "z_Adjustments"
Option Private Module
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Private Const GWL_STYLE As Long = (-16)
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_SIZEBOX = &H40000

Public Sub FormatUserForm(UserFormCaption As String)
Dim Frmhdl As Long
Dim lStyle As Long

Frmhdl = FindWindow(vbNullString, UserFormCaption)

lStyle = GetWindowLong(Frmhdl, GWL_STYLE)
lStyle = lStyle Or WS_SYSMENU
lStyle = lStyle Or WS_MINIMIZEBOX
lStyle = lStyle Or WS_MAXIMIZEBOX
lStyle = lStyle Or WS_SIZEBOX

SetWindowLong Frmhdl, GWL_STYLE, (lStyle)
DrawMenuBar Frmhdl

End Sub
