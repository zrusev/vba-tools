Attribute VB_Name = "TreeBuilder"
Option Private Module
Option Explicit

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Const INFINITE = &HFFFF 'Infinite timeout
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const PROCESS_ALL_ACCESS = &H1F0FFF

Public Sub Load_File(strPath As String)

Dim lonTaskID As Long
Dim lonProcessID As Long

SetAttr strPath, vbReadOnly
lonTaskID = Shell("explorer.exe " & strPath, vbMaximizedFocus)
    
''Below is optional, to wait for notepad to close before resetting attributes to normal
'lonProcessID = lonProcessID = OpenProcess(PROCESS_ALL_ACCESS, True, lonTaskID)
'WaitForSingleObject lonProcessID, INFINITE
'SetAttr strPath, vbNormal

End Sub

Public Sub Load_Form()

UserForm1.Show False

End Sub
