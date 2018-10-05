Attribute VB_Name = "Forms"
Option Explicit
Option Private Module

Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_CAPTION As Long = &HC00000
Private Const GWL_STYLE As Long = -16

Public Sub MakeFormResizableAndRemoveHeader()
Dim Hwnd As Long: Hwnd = GetForegroundWindow

Dim lStyle As Long: lStyle = GetWindowLong(Hwnd, GWL_STYLE) Or WS_THICKFRAME

Dim RetVal: RetVal = SetWindowLong(Hwnd, GWL_STYLE, lStyle)

lStyle = lStyle And (Not WS_CAPTION)

RetVal = SetWindowLong(Hwnd, -16, lStyle)

DrawMenuBar lStyle
End Sub

Public Function GetFormPosition() As rect
Dim Hwnd As Long: Hwnd = GetForegroundWindow

Dim rec As rect
Dim result As Long: result = GetWindowRect(Hwnd, rec)

GetFormPosition = rec
End Function
