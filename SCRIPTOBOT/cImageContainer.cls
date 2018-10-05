VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cImageContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Reference to ImageContainer.tlb is required

Private clWrapField As Wrapper

Private Property Let clWrapper(ByVal vNewValue As Variant): Set clWrapField = vNewValue: End Property

Public Property Get clWrapper() As Variant: Set clWrapper = clWrapField: End Property

Private Sub Class_Initialize(): clWrapper = New Wrapper: End Sub
