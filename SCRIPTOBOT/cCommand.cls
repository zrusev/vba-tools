VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CompareLevelConst As Double = 0.95
Private Const SimilarityThresholdConst As Double = 0.5

Private pPOINTAPI As POINTAPI
Private rRECT As rect
Private mCommand As MOUSECOMMAND
Private kCommand As KEYBOARDCOMMAND
Private keysArray() As KEYBOARDCOMMAND
Private hCommand As HARDWAREINPUTCOMMAND

Private iContainer As cImageContainer

Public MethodsLibrary As Dictionary

Private Sub Class_Initialize()
Set iContainer = New cImageContainer
Set MethodsLibrary = New Dictionary

With MethodsLibrary
    .CompareMode = TextCompare
    .Add "Left Click", "LeftClick"
    .Add "Right Click", "RightClick"
    .Add "Set Position", "SetCursorPosition"
    .Add "Send Keys", "SendKeyboardCombination"
    .Add "Copy Text", "PutInClipboard"
    .Add "Paste Text", "SendKeystrokes"
    .Add "Sleep", "Sleep"
    '.Add "Compare Images", "CompareImages" 'public Boolean CompareImages(Bitmap image, string targetImage, double compareLevel, float similarityThreshold, string filepath = DefaultSaveToPath)
    .Add "Capture Desktop", "CaptureDesktop" 'public Bitmap CaptureDesktop()
    .Add "Capture Active Window", "CaptureActiveWindow" 'public Bitmap CaptureActiveWindow()
    .Add "Capture Window", "CaptureWindow" 'public Bitmap CaptureWindow(string name)
    .Add "Bring Window To Front", "BringWindowToFront" 'public int BringWindowToFront(string name) => ScreenCapture.BringWindowToFront(name);
    '.Add "Capture Window By Coordinates", "CaptureWindowByCoordinates" 'public Bitmap CaptureWindowByCoordinates(int left, int top, int right, int bottom)
    '.Add "GetBitmapBytes", "GetBitmapBytes" 'public byte[] GetBitmapBytes(Bitmap image)
    .Add "Wait For Image", "WaitForImageToAppear"
    .Add "Wait For Text", "WaitForTextToAppear"
    .Add "Copy Text From Image", "ReadTextFromImage"
    .Add "Copy Text From Position", "ReadTextFromPosition"
    .Add "Cut Text", "CutText"
End With
End Sub

Private Sub SetPointer(x As Long, y As Long): pPOINTAPI.x = x: pPOINTAPI.y = y: End Sub

Private Function CompareImages(image As Object, targetImage As String, compareLevel As Double, similarityThreshold As Double, Optional filepath As String) As Boolean
CompareImages = iContainer.clWrapper.CompareImages(image, targetImage, compareLevel, similarityThreshold, filepath)
End Function

Private Function CaptureWindowByCoordinates(parameters As Variant) As Object
Dim left As Integer: left = CInt(parameters(0))
Dim top As Integer: top = CInt(parameters(1))
Dim right As Integer: right = CInt(parameters(2))
Dim bottom As Integer: bottom = CInt(parameters(3))

Set CaptureWindowByCoordinates = iContainer.clWrapper.CaptureWindowByCoordinates(left, top, right, bottom)
End Function

Private Function GetFromClipBoard() As String
Dim objData As New DataObject

objData.GetFromClipBoard
GetFromClipBoard = objData.GetText()
End Function

Public Function SetCursorPosition(parameters As Variant) As Long
SetPointer CLng(parameters(0)), CLng(parameters(1))
SetCursorPosition = SetCursorPos(pPOINTAPI.x, pPOINTAPI.y)
End Function

Public Function ClientScreen(hWks As Long) As Long: ClientScreen = ClientToScreen(hWks, pPOINTAPI): End Function

Public Function LeftClick(parameters As Variant) As Long
mCommand.iType = INPUT_MOUSE

Dim i As Integer
For i = 1 To CInt(parameters(0))
    mCommand.iFlags = MOUSEEVENTF_LEFTDOWN
    LeftClick = SendInput(1&, mCommand, Len(mCommand))

    mCommand.iFlags = MOUSEEVENTF_LEFTUP
    LeftClick = SendInput(1&, mCommand, Len(mCommand))
Next i
End Function

Public Function RightClick(parameters As Variant) As Long
mCommand.iType = INPUT_MOUSE

Dim i As Integer
For i = 1 To CInt(parameters(0))
    mCommand.iFlags = MOUSEEVENTF_RIGHTDOWN
    RightClick = SendInput(1&, mCommand, Len(mCommand))

    mCommand.iFlags = MOUSEEVENTF_RIGHTUP
    RightClick = SendInput(1&, mCommand, Len(mCommand))
Next i
End Function

Public Function SendKeyboardCombination(parameters As Variant) As Long
Dim stepsCount As Integer: stepsCount = UBound(parameters) + 1

'twice the count of parameters as we press and release
ReDim keysArray(0 To stepsCount * 2 - 1)

'first press control key
With keysArray(LBound(keysArray))
    .dwType = INPUT_KEYBOARD 'operation type
    .wVk = TextToConstant(CStr(parameters(LBound(keysArray)))) 'press CTRL key
End With

Dim pCount As Integer: pCount = 1

'press and release each letter key
Dim i As Integer
For i = LBound(keysArray) + 1 To UBound(keysArray) - 2 Step 2
    With keysArray(i)
        .dwType = INPUT_KEYBOARD 'operation type
        .wVk = VkKeyScan(Asc(parameters(pCount))) 'press chr key
    End With
    
    With keysArray(i + 1)
        .dwType = INPUT_KEYBOARD 'operation type
        .wVk = VkKeyScan(Asc(parameters(pCount)))
        .dwFlags = KEYEVENTF.KeyUp 'release chr key
    End With
    pCount = pCount + 1
Next i

'last release control key
With keysArray(UBound(keysArray))
    .dwType = INPUT_KEYBOARD ' operation type
    .wVk = TextToConstant(CStr(parameters(LBound(keysArray))))
    .dwFlags = KEYEVENTF.KeyUp 'release CTRL Key
End With

SendKeyboardCombination = SendInput(UBound(keysArray) + 1, keysArray(0), Len(keysArray(0)))

Erase keysArray
End Function

Public Function Sleep(parameters As Variant) As Integer: Wait CInt(parameters(0)): Sleep = 1: End Function

Public Function PutInClipboard(parameters As Variant) As Long
Dim MyDataObj As New DataObject

MyDataObj.SetText CStr(parameters(0))
MyDataObj.PutInClipboard

PutInClipboard = 1
End Function

Public Function SendKeystrokes(parameters As Variant) As Integer: Application.SendKeys "^v": SendKeystrokes = 1: End Function

Public Function CaptureDesktop(parameters As Variant) As Boolean
Dim sWindow As Object: Set sWindow = iContainer.clWrapper.CaptureDesktop
CaptureDesktop = CompareImages(sWindow, CStr(parameters(1)), CompareLevelConst, SimilarityThresholdConst)
End Function

Public Function CaptureActiveWindow(parameters As Variant) As Boolean
Dim sWindow As Object: Set sWindow = iContainer.clWrapper.CaptureActiveWindow
CaptureActiveWindow = CompareImages(sWindow, CStr(parameters(1)), CompareLevelConst, SimilarityThresholdConst)
End Function

Public Function CaptureWindow(parameters As Variant) As Boolean
Dim sWindow As Object: Set sWindow = iContainer.clWrapper.CaptureWindow(CStr(parameters(0)))
CaptureWindow = CompareImages(sWindow, CStr(parameters(1)), CompareLevelConst, SimilarityThresholdConst)
End Function

Public Function BringWindowToFront(parameters As Variant) As Integer: BringWindowToFront = iContainer.clWrapper.BringWindowToFront(CStr(parameters(0))): End Function

Public Function ReadTextFromImage(parameters As Variant) As Integer
Dim container(0) As String: container(0) = OCR.GetText(CStr(parameters(0)))

PutInClipboard container

ReadTextFromImage = 1
End Function

Public Function ReadTextFromPosition(parameters As Variant) As Integer
Dim sWindow As Object: Set sWindow = CaptureWindowByCoordinates(parameters)
Dim bArray() As Byte: bArray = iContainer.clWrapper.GetBitmapBytes(sWindow)
Dim container(0) As String: container(0) = OCR.GetText("", bArray, sWindow)

PutInClipboard container

ReadTextFromPosition = 1
End Function

Public Function WaitForImageToAppear(parameters As Variant) As Integer
Dim result As Boolean: result = False
Dim checksBetweenSeconds As Integer: checksBetweenSeconds = 2
Dim maxSecondsForChecks As Integer: maxSecondsForChecks = 60
Dim targetImage As String: targetImage = CStr(parameters(4))

Dim startStamp As Date: startStamp = Now
Dim currentStamp As Date
Dim elapsedSeconds As Long

Do
    Dim sWindow As Object: Set sWindow = CaptureWindowByCoordinates(parameters)
    result = CompareImages(sWindow, targetImage, CompareLevelConst, SimilarityThresholdConst)
    
    Wait checksBetweenSeconds
    Debug.Print "Executed command: Waiting for " & checksBetweenSeconds & " seconds"
        
    If GetKeyState(VK_SHIFT) < 0 Then action.TerminateExecution
    
    currentStamp = Now
    elapsedSeconds = DateDiff("s", startStamp, currentStamp)

    If elapsedSeconds >= maxSecondsForChecks Then
        WaitForImageToAppear = 0
        Exit Function
    End If
Loop Until result

WaitForImageToAppear = 1
End Function

Public Function WaitForTextToAppear(parameters As Variant) As Integer
Dim result As Boolean: result = False
Dim checksBetweenSeconds As Integer: checksBetweenSeconds = 2
Dim maxSecondsForChecks As Integer: maxSecondsForChecks = 60
Dim targetText As String: targetText = CStr(parameters(4))

Dim startStamp As Date: startStamp = Now
Dim currentStamp As Date
Dim elapsedSeconds As Long

Do
    Dim sWindow As Object: Set sWindow = CaptureWindowByCoordinates(parameters)
    Dim bArray() As Byte: bArray = iContainer.clWrapper.GetBitmapBytes(sWindow)
    Dim resultText As String: resultText = OCR.GetText("", bArray, sWindow)
    
    If InStr(1, resultText, targetText, vbTextCompare) > 0 Then Exit Do
    
    Wait checksBetweenSeconds
    Debug.Print "Executed command: Waiting for " & checksBetweenSeconds & " seconds"
    
    If GetKeyState(VK_SHIFT) < 0 Then action.TerminateExecution
    
    currentStamp = Now
    elapsedSeconds = DateDiff("s", startStamp, currentStamp)

    If elapsedSeconds >= maxSecondsForChecks Then
        WaitForTextToAppear = 0
        Exit Function
    End If
Loop Until result

WaitForTextToAppear = 1
End Function

Public Function CutText(parameters As Variant) As Integer
Dim startChar As Integer: startChar = CInt(parameters(0))
Dim endChar As Integer: endChar = CInt(parameters(1))

Dim container(0) As String: container(0) = GetFromClipBoard

If container(0) <> "" Then
    container(0) = Mid(container(0), startChar, endChar)
    Debug.Print container(0)
End If

PutInClipboard container

CutText = 1
End Function
