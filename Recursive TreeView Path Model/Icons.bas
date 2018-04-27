Attribute VB_Name = "Icons"
Option Private Module
Option Explicit

Function ReturnImageList(imL As ImageList) As ImageList

With imL.ListImages
    .Add , "folder", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\Folder.bmp"))
    .Add , "word", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\Word.bmp"))
    .Add , "excel", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\Excel.bmp"))
    .Add , "text", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\Text.bmp"))
    .Add , "link", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\Link.bmp"))
    .Add , "url", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\Url.bmp"))
    .Add , "pdf", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\PDF.bmp"))
    .Add , "ppt", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\PPT.bmp"))
End With

End Function

Function GetFileType(strValue As String) As String

If InStr(1, strValue, ".doc", vbTextCompare) > 0 Then
    GetFileType = "word"
    Exit Function
ElseIf InStr(1, strValue, ".dotx", vbTextCompare) > 0 Then
    GetFileType = "word"
    Exit Function
ElseIf InStr(1, strValue, ".xls", vbTextCompare) > 0 Then
    GetFileType = "excel"
    Exit Function
ElseIf InStr(1, strValue, ".pdf", vbTextCompare) > 0 Then
    GetFileType = "pdf"
    Exit Function
ElseIf InStr(1, strValue, ".ppt", vbTextCompare) > 0 Then
    GetFileType = "ppt"
    Exit Function
ElseIf InStr(1, strValue, ".txt", vbTextCompare) > 0 Then
    GetFileType = "text"
    Exit Function
ElseIf InStr(1, strValue, ".lnk", vbTextCompare) > 0 Then
    GetFileType = "link"
    Exit Function
ElseIf InStr(1, strValue, ".url", vbTextCompare) > 0 Then
    GetFileType = "url"
    Exit Function
End If

GetFileType = "folder"

End Function

Function FileExists(strDir As String) As String

If Dir(strDir) <> "" Then
    MsgBox strDir & " does not exist." & vbNewLine & vbNewLine & "Please make sure you have all icons loaded.", vbInformation, "System"
    End
End If

FileExists = strDir

End Function
