Attribute VB_Name = "z_Functions"
Declare Function Get_User_Name Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Function GetUserName() As String
Dim lpBuff As String * 25

Get_User_Name lpBuff, 25
GetUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

End Function
Public Function Wait(ByVal sec As Integer)

timeStamp = Now + TimeSerial(0, 0, sec)
Do
  DoEvents
Loop Until Now > timeStamp

End Function
