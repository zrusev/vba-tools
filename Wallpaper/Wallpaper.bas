Attribute VB_Name = "Wallpaper"
Option Private Module
Option Explicit

Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPIF_SENDWININICHANGE = &H2

Private Const REG_SZ = 1

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
    ByVal uAction As Long, _
    ByVal uParam As Long, _
    ByVal lpvParam As Any, _
    ByVal fuWinIni As Long) As Long

Public Enum REG_TOPLEVEL_KEYS
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" ( _
    ByVal Hkey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
    ByVal Hkey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, lpData As Any, _
    ByVal cbData As Long) As Long

Sub PictureDir()
Dim dskPath As String

With Application.FileDialog(msoFileDialogFilePicker)
    .Title = "Select the previous report:"
    .InitialFileName = GetDesktop
    .Show
    
    On Error GoTo exitSub:
    dskPath = .SelectedItems(1)
End With

ChangeWallpaper dskPath, False
Exit Sub

exitSub:
End Sub

Private Function GetDesktop() As String
Dim oWSHShell As Object

Set oWSHShell = CreateObject("WScript.Shell")
GetDesktop = oWSHShell.SpecialFolders("Desktop")
Set oWSHShell = Nothing

End Function

Public Function ChangeWallpaper(ImageFile As String, Tile As Boolean)

'Returns true if successful, false otherwise
'If you want to tile, set Tile to True

Dim lRet As Long
On Error Resume Next

If Tile Then WriteStringToRegistry HKEY_CURRENT_USER, "Control Panel\desktop", "TileWallpaper", "1"
  
lRet = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, ImageFile, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)

ChangeWallpaper = lRet <> 0 And Err.LastDllError = 0

End Function

Private Function WriteStringToRegistry(Hkey As REG_TOPLEVEL_KEYS, _
                                       strPath As String, _
                                       strValue As String, _
                                       strdata As String) As Boolean
Dim bAns As Boolean

On Error GoTo ErrorHandler
   
Dim keyhand As Long
Dim r As Long

r = RegCreateKey(Hkey, strPath, keyhand)
If r = 0 Then
    r = RegSetValueEx(keyhand, strValue, 0, _
    REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End If
    
WriteStringToRegistry = (r = 0)

Exit Function

ErrorHandler:
WriteStringToRegistry = False
Exit Function
    
End Function
