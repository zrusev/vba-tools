Attribute VB_Name = "References"
Option Explicit
Option Private Module

Private Const ReferenceName As String = "ImageContainer"
Private Const ReferencePath As String = "C:\Output\ImageContainer.tlb"

Public Sub ReferenceCheck()
Dim VBAEditor As VBIDE.VBE: Set VBAEditor = Application.VBE
Dim vbProj As VBIDE.VBProject: Set vbProj = ActiveWorkbook.VBProject

Dim chkRef As VBIDE.Reference
For Each chkRef In vbProj.References
    If chkRef.Name = ReferenceName Then
        Dim BoolExists As Boolean: BoolExists = True
        GoTo CleanUp
    End If
Next

If Not FileThere(ReferencePath) Then
    MsgBox "The following file: " & ReferencePath & " could not be found!", vbCritical, "System"
    GoTo Finish:
End If

On Error Resume Next
vbProj.References.AddFromFile ReferencePath

CleanUp:
If BoolExists = True Then
    Debug.Print "Reference to " & ReferencePath & " already exists."
Else
    Debug.Print "Reference to " & ReferencePath & " added successfully."
End If

Finish:
Set vbProj = Nothing
Set VBAEditor = Nothing
End Sub

Private Function FileThere(FileName As String) As Boolean: FileThere = (Dir(FileName) > ""): End Function
