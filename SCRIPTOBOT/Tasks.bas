Attribute VB_Name = "Tasks"
Option Explicit
Option Private Module

Private Const mcGWCHILD = 5
Private Const mcGWHWNDNEXT = 2
Private Const mcGWLSTYLE = (-16)
Private Const mcWSVISIBLE = &H10000000
Private Const mconMAXLEN = 255

Public Function EnumWindows() As Variant
Dim lngx As Long: lngx = GetDesktopWindow()
Dim lngLen As Long: lngx = GetWindow(lngx, mcGWCHILD)

Dim oDic As Object, a() As Variant
Set oDic = CreateObject("Scripting.Dictionary")

Dim lngStyle As Long, strCaption As String
Do While Not lngx = 0
    strCaption = GetCaption(lngx)
        If Len(strCaption) > 0 Then
            lngStyle = GetWindowLong(lngx, mcGWLSTYLE)
            If lngStyle And mcWSVISIBLE Then 'enum visible windows only
                Dim caption As String: caption = GetCaption(lngx)
                If Not oDic.Exists(caption) Then oDic.Add caption, caption
            End If
        End If
    lngx = GetWindow(lngx, mcGWHWNDNEXT)
Loop
EnumWindows = oDic.keys
End Function

Private Function GetCaption(Hwnd As Long) As String
Dim strBuffer As String: strBuffer = String$(mconMAXLEN - 1, 0)
Dim intCount As Integer: intCount = GetWindowText(Hwnd, strBuffer, mconMAXLEN)
If intCount > 0 Then
    GetCaption = left$(strBuffer, intCount)
End If
End Function
