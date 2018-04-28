Attribute VB_Name = "mod_Socket"
Public Const ScpiPort = 804

Public sendMessage As String
Private Hostname$
Private WorkbookPath$

Sub Get_hostname()
    
Hostname$ = "127.0.0.1"

End Sub

Sub Get_WorkbookPath()

WorkbookPath$ = ThisWorkbook.Path & Application.PathSeparator & ThisWorkbook.Name

End Sub

Function GetUserform1() As UserForm1

Set GetUserform1 = UserForm1

End Function

Sub ListenToServer()
Dim VBC As Object
Dim y As String

Call Get_WorkbookPath
Set VBC = Application.Run("'" & WorkbookPath$ & "'!GetUserform1")

Do

y = RecvStrTO(Application.StatusBar)

If y <> "" Then
    VBC.Hide
    VBC.TextBox2.Text = y
    'Debug.Print y
    VBC.Show modeless
End If

'Wait 1
Loop Until Sheets(1).Range("A1") = "exit"

End Sub
Sub test()

UserForm1.Show vbModeless
Do While sendMessage = ""
DoEvents
Loop

End Sub
Sub ConnectToSocket()
Dim x As Long
Dim fieldText As String
Dim currentUser As String
Dim lineBeginning As String

sendMessage = ""
currentUser = GetUserName
lineBeginning = "<<" & currentUser & ">> "

Call StartIt
Call Get_hostname
x = OpenSocket(Hostname$, ScpiPort)

LoopAgain:
UserForm1.Show vbModaless
Do While sendMessage = ""
    y = RecvStrTO(socketId)
    UserForm1.TextBox2.Text = y
    Wait 3
Loop

x = sendCommand(lineBeginning & sendMessage)
If x = -1 Then Exit Sub
UserForm1.TextBox2.Text = ""

If Not sendMessage = "exit" Then GoTo LoopAgain:

Call CloseConnection
Call EndIt

End Sub
