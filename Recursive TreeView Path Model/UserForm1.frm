VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Library"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7725
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click(): Unload Me: End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
Dim objNode As Node
Dim lngItem As Long, lngCount As Long

Node.ForeColor = vbBlack
Node.BackColor = RGB(255, 255, 255)

lngCount = Node.Children

If Not lngCount = 0 Then

    Set objNode = Node.Child

    For lngItem = 1 To lngCount
        If objNode.Expanded = True Then objNode.Expanded = False
        Set objNode = objNode.Next
    Next
End If
Set objNode = Nothing

End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)

Node.BackColor = RGB(0, 143, 255)
Node.ForeColor = vbWhite

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
Dim fileToOpen As String

If InStr(1, Node.Key, ".", vbTextCompare) > 0 Then
    fileToOpen = DirectoryPath & Right(Node.Key, Len(Node.Key) - 7)
    Load_File fileToOpen
Else
    Node.Expanded = True
End If

End Sub

Private Sub UserForm_Initialize()
Dim imL As ImageList

FormatUserForm Me.Caption

Set imL = New ImageList1
ReturnImageList imL

TreeView1.ImageList = imL

LoopDatabase

End Sub

Public Sub AddToNode(ByVal activeCellKey As String, ByVal activeCellValue As String, ByVal previousCellKey As String, ByVal previousCellValue As String)

If NodeExists(previousCellKey) = False Then
    TreeView1.Nodes.Add , , previousCellKey, previousCellValue, GetFileType(previousCellKey)
End If

If NodeExists(activeCellKey) = False Then
    TreeView1.Nodes.Add previousCellKey, tvwChild, activeCellKey, activeCellValue, GetFileType(activeCellKey)
End If

End Sub

Private Function NodeExists(ByVal strKey As String) As Boolean

Dim Node As MSComctlLib.Node
    
On Error Resume Next
    
Set Node = TreeView1.Nodes(strKey)
Select Case Err.Number
    Case 0
        NodeExists = True
    Case Else
        NodeExists = False
End Select

End Function

Private Sub UserForm_Resize()

With Me.TreeView1
    .Move 0, 0, Me.Width - 10, Me.Height
End With
    
End Sub
