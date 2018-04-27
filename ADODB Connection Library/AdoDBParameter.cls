VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoDBParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private param As ADODB.Parameter
Private parameterDataTypeValue As Integer
Private parameterSizeValue As Integer
Private parameterDirectionValue As Integer

Private Property Get DirectionValue() As Integer: DirectionValue = parameterDirectionValue: End Property
Private Property Let DirectionValue(ByVal pDirectionValue As Integer): parameterDirectionValue = pDirectionValue: End Property
Private Property Get SizeValue() As Integer: SizeValue = parameterSizeValue: End Property
Private Property Let SizeValue(ByVal pSizeValue As Integer): parameterSizeValue = pSizeValue: End Property
Private Property Get DataTypeValue() As Integer: DataTypeValue = parameterDataTypeValue: End Property
Private Property Let DataTypeValue(ByVal pDataTypeValue As Integer): parameterDataTypeValue = pDataTypeValue: End Property
Private Property Get Parameter() As ADODB.Parameter: Set Parameter = param: End Property
Private Property Let Parameter(ByVal adoParam As ADODB.Parameter): Set param = adoParam: End Property

Private Sub Class_Terminate(): Set param = Nothing: End Sub

Private Sub Class_Initialize()

Set param = Nothing
Set param = New ADODB.Parameter

End Sub

Public Function Add(ByVal parameterValue As Variant) As ADODB.Parameter

DataTypeValue = GetParameterType(typeName(parameterValue))
If LCase(typeName(parameterValue)) <> LCase("Null") Then SizeValue = Len(parameterValue)
If SizeValue = 0 Then SizeValue = 1
DirectionValue = ParameterDirections.adParamInput
If parameterValue = "" Then parameterValue = Null

Parameter.Type = DataTypeValue
Parameter.Direction = DirectionValue
Parameter.Size = SizeValue
Parameter.Value = parameterValue

Set Add = Parameter

End Function
