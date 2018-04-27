Attribute VB_Name = "PathBuilder"
Option Private Module
Option Explicit

Sub LoopDatabase()

Dim previousCellKey As String, previousCellValue As String, activeCellKey As String, activeCellValue As String
Dim arr As Variant

Dim colFiles As New Collection
RecursiveDir colFiles, DirectoryPath, FileExtension, True

Dim y As Long
For y = 1 To colFiles.Count
    arr = Split(MasterPathName & Application.PathSeparator & Replace(colFiles(y), TrailingSlash(DirectoryPath), "", , , vbTextCompare), Application.PathSeparator)
    
    previousCellKey = ""
    previousCellValue = ""
    activeCellKey = ""
    activeCellValue = ""
    
    Dim coll As Long
    For coll = LBound(arr) + 1 To UBound(arr)
        previousCellKey = previousCellKey & Application.PathSeparator & arr(coll - 1)
        previousCellValue = arr(coll - 1)
        activeCellKey = previousCellKey & Application.PathSeparator & arr(coll)
        activeCellValue = arr(coll)
            
        If activeCellValue <> "" Then
            UserForm1.AddToNode activeCellKey, activeCellValue, previousCellKey, previousCellValue
        End If
    Next coll
Next y

End Sub

Private Function RecursiveDir(ByRef colFiles As Collection, strFolder As String, strFileSpec As String, bIncludeSubfolders As Boolean)
Dim colFolders As New Collection

strFolder = TrailingSlash(strFolder)

Dim strTemp As String
strTemp = Dir(strFolder & strFileSpec)

Do While strTemp <> vbNullString
    colFiles.Add strFolder & strTemp
    strTemp = Dir
Loop

If bIncludeSubfolders Then

    strTemp = Dir(strFolder, vbDirectory)
    Do While strTemp <> vbNullString
        If (strTemp <> ".") And (strTemp <> "..") Then
            If Len(strFolder & strTemp) >= 255 Then
                MsgBox "The file's name is too long." & vbNewLine & _
                       "The lenght should not exceed 255 symbols." & vbNewLine & _
                       "'" & strFolder & strTemp & "'", vbInformation, "System"
                End
            End If
            If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0 Then
                colFolders.Add strTemp
            End If
        End If
        strTemp = Dir
    Loop
            
    Dim vFolderName As Variant
    For Each vFolderName In colFolders
        Call RecursiveDir(colFiles, strFolder & vFolderName, strFileSpec, True)
    Next vFolderName
End If

End Function

Private Function TrailingSlash(strFolder As String) As String

If Len(strFolder) > 0 Then
    If Right(strFolder, 1) = Application.PathSeparator Then
        TrailingSlash = strFolder
    Else
        TrailingSlash = strFolder & Application.PathSeparator
    End If
End If

End Function
