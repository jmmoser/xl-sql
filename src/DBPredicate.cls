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
Private pOperator As Long
Private pParameters As Variant
Private pParameterCount As Long

Public Sub InitiateProperties(Name As String, op As Long, Optional Params As Variant)
    pName = Name
    pOperator = op
    pParameters = DB.ConvertToArray(Params)
    pParameterCount = DB.ArrayLength(pParameters)
End Sub

Public Sub SetColumnNumber(colNumber)
    pColumn = colNumber
End Sub

Public Property Get Column()
    Column = pColumn
End Property

Public Property Get Name()
    Name = pName
End Property

Public Property Get Operator()
    Operator = pOperator
End Property

Public Property Get Parameter(idx As Long)
    Parameter = pParameters(idx)
End Property

Public Property Get ParameterCount()
    ParameterCount = pParameterCount
End Property
