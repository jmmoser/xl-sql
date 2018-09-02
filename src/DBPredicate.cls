VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBPredicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pColumn As Long
Private pName As String
Private pOperator As OP
Private pParameters As Variant
Private pParameterCount As Long

Public Sub InitiateProperties(Name As String, Operator As OP, Optional Params As Variant)
    pName = Name
    pOperator = Operator
    pParameters = DB.ConvertToArray(Params)
    pParameterCount = DB.ArrayLength(pParameters)
    pColumn = -1
End Sub

Public Sub SetColumnNumber(colNumber)
    pColumn = colNumber
End Sub

Public Property Get Column() As Long
    Column = pColumn
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Get Operator() As OP
    Operator = pOperator
End Property

Public Property Get Parameter(idx As Long)
    If VarType(pParameters(idx)) = vbObject Then
        Set Parameter = pParameters(idx)
    Else
        Parameter = pParameters(idx)
    End If
End Property

Public Property Get ParameterCount() As Long
    ParameterCount = pParameterCount
End Property
