Attribute VB_Name = "Enums"
Option Explicit
Option Private Module

Public Enum SHOW_WINDOW
    SW_HIDE = 0
    SW_SHOWNORMAL = 1
    SW_NORMAL = 1
    SW_SHOWMINIMIZED = 2
    SW_SHOWMAXIMIZED = 3
    SW_MAXIMIZE = 3
    SW_SHOWNOACTIVATE = 4
    SW_SHOW = 5
    SW_MINIMIZE = 6
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_RESTORE = 9
    SW_SHOWDEFAULT = 10
    SW_FORCEMINIMIZE = 11
    SW_MAX = 11
End Enum

Public Enum MOUSEEVENTF
    Move = &H1
    LEFTDOWN = &H2
    LEFTUP = &H4
    RIGHTDOWN = &H8
    RIGHTUP = &H10
    MIDDLEDOWN = &H20
    MIDDLEUP = &H40
    XDOWN = &H80
    XUP = &H100
    VIRTUALDESK = &H400
    WHEEL = &H800
    ABSOLUTE = &H8000
End Enum

Public Enum KEYEVENTF
    EXTENDEDKEY = 1
    KeyUp = 2
    [UNICODE] = 4
    SCANCODE = 8
End Enum
