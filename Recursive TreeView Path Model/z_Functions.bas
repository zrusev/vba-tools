Attribute VB_Name = "z_Functions"
Option Private Module

Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Const HWND_DESKTOP As Long = 0
Const LOGPIXELSX As Long = 88
Const LOGPIXELSY As Long = 90

'--------------------------------------------------
Function TwipsPerPixelX() As Single
'--------------------------------------------------
'Returns the width of a pixel, in twips.
'--------------------------------------------------
Dim lngDC As Long
lngDC = GetDC(HWND_DESKTOP)
TwipsPerPixelX = 1440& / GetDeviceCaps(lngDC, LOGPIXELSX)
ReleaseDC HWND_DESKTOP, lngDC
End Function

'--------------------------------------------------
Function TwipsPerPixelY() As Single
'--------------------------------------------------
'Returns the height of a pixel, in twips.
'--------------------------------------------------
Dim lngDC As Long
lngDC = GetDC(HWND_DESKTOP)
TwipsPerPixelY = 1440& / GetDeviceCaps(lngDC, LOGPIXELSY)
ReleaseDC HWND_DESKTOP, lngDC
End Function

