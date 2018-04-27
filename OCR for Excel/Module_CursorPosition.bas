Attribute VB_Name = "Module_CursorPosition"
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10

Type POINTAPI
    x As Long
    y As Long
End Type

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Dim pos As POINTAPI

Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long


Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassname As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const WM_CLOSE = &H10

Public posX As Long
Public posY As Long

Sub Get_TextFromCursorPosition()

GetCursorPos pos
    posX = pos.x
    posY = pos.y
'Debug.Print pos.x & ", " & pos.y
Call PrivateTest

End Sub

Private Sub PrivateTest()
Attribute PrivateTest.VB_ProcData.VB_Invoke_Func = "q\n14"
Dim strRes

strRes = OCR(posX, posY, 800, 70)

MsgBox strRes

End Sub
