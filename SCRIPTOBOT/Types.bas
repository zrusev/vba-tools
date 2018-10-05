Attribute VB_Name = "Types"
Option Explicit
'Option Private Module

Public Type POINTAPI
   x As Long
   y As Long
End Type
 
Public Type rect
   left As Long         ' Screen X position
   top As Long          ' Screen Y position
   right As Long
   bottom As Long
End Type

Public Type MOUSECOMMAND
   iType As Long        ' 0 for mouse, 1 for kbd
   iDx As Long          ' rel movt in pixels (unless ABSOLUTE)
   iDy As Long
   iWheelData As Long   ' we don't use this
   iFlags As Long       ' we use this (see MOUSEEVENT flags below)
   iTime As Long        ' don't use this
   iXtra As Long        ' or this
End Type

Public Type KEYBOARDCOMMAND
   dwType As Long       'input type (keyboard or mouse)
   wVk As Integer       'the key to press/release as ASCSI scan code
   wScan As Integer     'not required
   dwFlags As Long      'specify if key is pressed or released
   dwTime As Long       'not required
   dwExtraInfo As Long  'not required
   dwPadding As Currency 'only required for mouse inputs
End Type

Public Type HARDWAREINPUTCOMMAND
   uMsg As Long
   wParamL As Long
   wParamH As Long
End Type
