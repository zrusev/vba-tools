VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal dst As Long, ByVal src As Long, ByVal Length As Long)

Private Const DEFAULT_CAPACITY As Long = &H10
Private m_currLen As Long
Private m_stringBuffer() As Byte

Private Sub Class_Initialize()
    
ReDim m_stringBuffer(0 To (DEFAULT_CAPACITY * 2) - 1) 'Each unicode character is 2 bytes

End Sub

Public Function Append(strString As String) As clsStringBuilder

On Error GoTo derp

If m_currLen + LenB(strString) < UBound(m_stringBuffer) Then
    CopyMemory VarPtr(m_stringBuffer(m_currLen)), StrPtr(strString), LenB(strString)
Else
    If m_currLen + LenB(strString) < UBound(m_stringBuffer) * 2 Then
        Expand
    Else
        Expand m_currLen + LenB(strString)
    End If
    CopyMemory VarPtr(m_stringBuffer(m_currLen)), StrPtr(strString), LenB(strString)
End If

m_currLen = m_currLen + LenB(strString)
Set Append = Me

Exit Function

derp:
    Stop
    Resume

End Function

Public Function AppendLine(strString As String) As clsStringBuilder

On Error GoTo derp

strString = strString & vbNewLine

If m_currLen + LenB(strString) < UBound(m_stringBuffer) Then
    CopyMemory VarPtr(m_stringBuffer(m_currLen)), StrPtr(strString), LenB(strString)
Else
    If m_currLen + LenB(strString) < UBound(m_stringBuffer) * 2 Then
        Expand
    Else
    Expand m_currLen + LenB(strString)
End If
    CopyMemory VarPtr(m_stringBuffer(m_currLen)), StrPtr(strString), LenB(strString)
End If

m_currLen = m_currLen + LenB(strString)
Set AppendLine = Me

Exit Function

derp:
    Stop
    Resume

End Function

Public Property Get Length() As Long

Length = m_currLen / 2

End Property

Public Property Get Capacity() As Long

Capacity = UBound(m_stringBuffer)

End Property

Private Sub Expand(Optional newSize As Long = 0)

If newSize <> 0 Then
    ReDim Preserve m_stringBuffer(0 To newSize - 1)
Else
    ReDim Preserve m_stringBuffer(0 To (UBound(m_stringBuffer) * 2) + 1)
End If

End Sub

Public Function toString() As String

toString = Mid(m_stringBuffer, 1, m_currLen / 2)

End Function
