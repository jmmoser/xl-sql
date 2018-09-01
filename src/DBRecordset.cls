VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pData As Variant
Private pDataUpperBound As Long
Private pColumnIndices As New Dictionary
Private pIndex As Long


Public Property Get Value(Name As Variant, Optional Index As Long = -1) As Variant
    Value = DB.NullValue()
    
    If pColumnIndices.Exists(Name) = False Then Exit Property
    
    If Index > 0 Then
        If Index <= pDataUpperBound Then
            Value = pData(Index, pColumnIndices(Name))
        End If
    ElseIf pIndex <= pDataUpperBound Then
        Value = pData(pIndex, pColumnIndices(Name))
    End If
End Property


Public Property Get EOF() As Boolean
    EOF = pIndex > pDataUpperBound
End Property


Public Sub MoveNext()
    pIndex = pIndex + 1
End Sub


Public Sub Cursor(Index As Long)
    If Index < 0 Then ResetIndex
    pIndex = Index
End Sub


Private Sub ResetIndex()
    If IsEmpty(pData) = False Then
        pIndex = LBound(pData, 1)
    Else
        pIndex = 0
    End If
End Sub


Public Sub Setup(vData As Variant, columnNames As Variant)
    Dim i As Long
    
    pData = vData
    
    If IsEmpty(vData) Then
        pIndex = 0
        pDataUpperBound = -1
    Else
        pIndex = LBound(vData, 1)
        pDataUpperBound = UBound(vData, 1)
        
        For i = LBound(vData, 2) To UBound(vData, 2)
            pColumnIndices.Add columnNames(i + LBound(columnNames, 1) - LBound(vData, 1)), i
        Next i
    End If
End Sub


