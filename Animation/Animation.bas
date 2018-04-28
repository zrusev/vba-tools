Attribute VB_Name = "Animation"
Option Private Module
Option Explicit

Sub Moving()
Dim rep_count As Long: rep_count = 0
Dim shape1 As Shape, shape2 As Shape

Set shape1 = ThisWorkbook.Sheets(1).Shapes("Rectangle 1")
Set shape2 = ThisWorkbook.Sheets(1).Shapes("Rectangle 2")

Do
    DoEvents
    rep_count = rep_count + 1
    
    If rep_count = 1 Then
        shape1.Rotation = 0
        shape2.Fill.Transparency = 1
    End If
    
    With shape1
        .Left = 900
        .Top = 500 - rep_count
        .Height = rep_count
        .Width = rep_count
    End With
    
    If rep_count >= 20 Then
        shape1.IncrementRotation (1)
    End If
    
    Timeout (0.01)
Loop Until rep_count = 70

Dim spTop As Double, spLeft As Double, spHeight As Double, spWeight As Double

With shape1
    spTop = .Top
    spLeft = .Left
    spHeight = .Height
    spWeight = .Width
End With

rep_count = 0

Do
    DoEvents
    rep_count = rep_count + 1
    
    shape2.Fill.Transparency = (rep_count + 30) / 100
    
    With shape1
        .Left = spLeft + rep_count
        .Top = spTop + rep_count
        .Height = spHeight - rep_count
        .Width = spWeight - rep_count
    End With
    
    If rep_count >= 20 Then
        shape1.IncrementRotation (1)
    End If
    
    Timeout (0.01)
Loop Until rep_count = 70

End Sub

Sub Timeout(duration_ms As Double)
Dim start_time As Date

start_time = Timer

Do
    DoEvents
Loop Until (Timer - start_time) >= duration_ms

End Sub
