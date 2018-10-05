Attribute VB_Name = "API"
Option Explicit
Option Private Module

Public Declare PtrSafe Function ShowWindow Lib "User32.dll" _
     (ByVal Hwnd As Long, _
      ByVal nCmdShow As SHOW_WINDOW) As Boolean

Public Declare PtrSafe Function SetForegroundWindow Lib "user32" _
     (ByVal Hwnd As Long) As Integer

Public Declare PtrSafe Function GetForegroundWindow Lib "user32" _
     () As Integer

Public Declare PtrSafe Function FindWindowA Lib "User32.dll" _
     (ByVal lpszClass As String, _
      ByVal lpszWindow As String) As Long

Public Declare PtrSafe Function FindWindowEx Lib "User32.dll" Alias "FindWindowExA" _
     (ByVal hWndParent As Long, _
      ByVal hWndChildAfter As Long, _
      ByVal lpszClass As String, _
      ByVal lpszWindow As String) As Long

Public Declare PtrSafe Function IsIconic Lib "User32.dll" _
     (ByVal Hwnd As Long) As Long

Public Declare PtrSafe Function SetCursorPos Lib "User32.dll" _
     (ByVal x As Long, _
      ByVal y As Long) As Long

Public Declare PtrSafe Function GetWindowRect Lib "User32.dll" _
     (ByVal Hwnd As Long, _
      ByRef lpRECT As rect) As Long

Public Declare Function GetWindow Lib "user32" _
     (ByVal Hwnd As Long, _
      ByVal wCmd As Long) As Long

Public Declare PtrSafe Function ClientToScreen Lib "User32.dll" _
     (ByVal Hwnd As Long, _
      ByRef lpPoint As POINTAPI) As Long

Public Declare PtrSafe Function ScreenToClient Lib "User32.dll" _
     (ByVal Hwnd As Long, _
      ByRef lpPoint As POINTAPI) As Long

Public Declare PtrSafe Function GetSystemMetrics Lib "User32.dll" _
     (ByVal nIndex As Long) As Long

Public Declare PtrSafe Function SendInput Lib "User32.dll" _
     (ByVal nInputs As LongPtr, _
      pInputs As Any, _
      ByVal cbSize As LongPtr) As LongPtr
      
Public Declare PtrSafe Function VkKeyScan Lib "User32.dll" Alias "VkKeyScanA" _
     (ByVal cChar As Byte) As Integer

Public Declare PtrSafe Sub Sleep Lib "kernel32.dll" _
     (ByVal dwMilliseconds As Long)

Public Declare PtrSafe Function GetCursorPos Lib "user32" _
     (lpPoint As POINTAPI) As Long
     
Public Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" _
     (ByVal Hwnd As Long, _
      ByVal lpClassname As String, _
      ByVal nMaxCount As Long) As Long

Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
     (ByVal Hwnd As Long, _
      ByVal lpString As String, _
      ByVal cch As Long) As Long
      
Public Declare PtrSafe Function GetKeyState Lib "User32.dll" _
     (ByVal nVirtKey As Long) As Integer

Public Declare PtrSafe Function GetWindowLong Lib "User32.dll" Alias "GetWindowLongA" _
     (ByVal Hwnd As Long, _
      ByVal nIndex As Long) As Long

Public Declare PtrSafe Function SetWindowLong Lib "User32.dll" Alias "SetWindowLongA" _
     (ByVal Hwnd As Long, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As Long) As Long

Public Declare PtrSafe Function DrawMenuBar Lib "user32" _
     (ByVal Hwnd As Long) As Long

Public Declare PtrSafe Function GetDesktopWindow Lib "user32" _
     () As Long

