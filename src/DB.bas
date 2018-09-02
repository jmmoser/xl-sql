Attribute VB_Name = "DB"
Option Explicit
Const TABLE_PREFIX = "tb_"


Public Enum OP
    EQUAL
    NOT_EQUAL
    GREATER_THAN
    LESS_THAN
    GREATER_THAN_EQUAL
    LESS_THAN_EQUAL
    IS_IN
    IS_NOT_IN
    IS_NULL
    IS_NOT_NULL
    BETWEEN
End Enum

Private Function OperatorStringToOP(ByVal opStr As String) As OP
    opStr = VBA.Strings.UCase$(VBA.Strings.Trim$(opStr))
    If VBA.Strings.Len(opStr) = 0 Then
        Err.Raise Number:=1, Description:="Operator string is empty"
    End If
    Select Case opStr
        Case "="
            OperatorStringToOP = EQUAL
        Case "!="
        Case "<>"
            OperatorStringToOP = NOT_EQUAL
        Case ">"
            OperatorStringToOP = GREATER_THAN
        Case "<"
            OperatorStringToOP = LESS_THAN
        Case ">="
            OperatorStringToOP = GREATER_THAN_EQUAL
        Case "<="
            OperatorStringToOP = LESS_THAN_EQUAL
        Case "IN"
            OperatorStringToOP = IS_IN
        Case "NOT IN"
            OperatorStringToOP = IS_NOT_IN
        Case "IS NULL"
            OperatorStringToOP = IS_NULL
        Case "IS NOT NULL"
            OperatorStringToOP = IS_NOT_NULL
        Case "BETWEEN"
            OperatorStringToOP = BETWEEN
        Case Else
            Err.Raise Number:=1, Description:="Not a valid operator"
    End Select
End Function


Public Function ArrayLength(val As Variant) As Variant
    ArrayLength = UBound(val, 1) - LBound(val, 1) + 1
End Function


Public Function ConvertToArray(val As Variant) As Variant
    If IsArray(val) Then
        ConvertToArray = val
    Else
        ConvertToArray = Array(val)
    End If
End Function


Public Function CombineArrays(ByVal arr1 As Variant, arr2 As Variant) As Variant
    Dim i As Long
    Dim pArrResult As Variant
    
    ReDim pArrResult(0 To ArrayLength(arr1) + ArrayLength(arr2) - 1)
    
    For i = LBound(arr1, 1) To UBound(arr1, 1)
        pArrResult(i) = arr1(i)
    Next i
    
    For i = LBound(arr2, 1) To UBound(arr2, 1)
        pArrResult(i + UBound(arr1, 1) + 1) = arr2(i)
    Next i
    
    CombineArrays = pArrResult
End Function


Public Function Delete(tableName As String, Optional ByVal predicates As Variant) As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As Variant
    
    Delete = 0
    
    With GetTable(tableName)
        tmp = .UsedRange.Value
        SetupPredicateColumnNumbers tableName, tmp, predicates
        j = 2
        For i = 2 To UBound(tmp, 1)
            If RowMatchesPredicates(tmp, i, predicates) Then
                .Rows(j).Delete
                Delete = Delete + 1
            Else
                j = j + 1
            End If
        Next i
    End With
End Function


Public Function Insert(tableName As String, ByVal columnNames As Variant, ByVal values As Variant) As Boolean
    Dim i As Long
    Dim j As Long
    Dim colNumbers As Variant
    Dim tmp As Variant
    
    columnNames = ConvertToArray(columnNames)
    values = ConvertToArray(values)
    
    If ArrayLength(columnNames) <> ArrayLength(values) Then
        Err.Raise Number:=1, Description:="Array count of columns and values must match"
    End If
    
    With GetTable(tableName)
        tmp = .UsedRange.Value
        
        ReDim colNumbers(LBound(columnNames, 1) To UBound(columnNames, 1))
        For j = LBound(columnNames, 1) To UBound(columnNames, 1)
            colNumbers(j) = GetColumnNumber(tmp, columnNames(j))
            If colNumbers(j) < 1 Then
                Err.Raise Number:=1, Description:="Column '" & columnNames(j) & "' not found in table '" & tableName & "'"
            End If
        Next j
        
        i = UBound(tmp, 1) + 1
        
        For j = LBound(colNumbers, 1) To UBound(colNumbers, 1)
            .Cells(i, colNumbers(j)).Value = values(j)
        Next j
    End With
End Function


Public Function Update(tableName As String, ByVal values As Variant, Optional ByVal predicates As Variant) As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As Variant
    Dim colNumbers() As Long
    Dim rowMatches As New Collection
    
    Update = 0
    
    values = ConvertToArray(values)
    
    With GetTable(tableName)
        tmp = .UsedRange.Value
        
        ReDim colNumbers(LBound(values, 1) To UBound(values, 1), 1 To 1)
        For i = LBound(values, 1) To UBound(values, 1)
            colNumbers(i, 1) = GetColumnNumber(tmp, CStr(values(i)(0)))
            If colNumbers(i, 1) < 1 Then
                Err.Raise Number:=1, Description:="Column '" & values(i)(0) & "' not found in table '" & tableName & "'"
            End If
        Next i

        SetupPredicateColumnNumbers tableName, tmp, predicates
    
        For i = 2 To UBound(tmp, 1)
            If RowMatchesPredicates(tmp, i, predicates) Then rowMatches.Add i
        Next i
        
        If rowMatches.Count = 0 Then Exit Function
        
        Update = rowMatches.Count
        
        ReDim vals(1 To rowMatches.Count, 1 To UBound(colNumbers, 1) - LBound(colNumbers, 1) + 1)
        For i = 1 To rowMatches.Count
            For j = LBound(colNumbers, 1) To UBound(colNumbers, 1)
                .Cells(rowMatches(i), colNumbers(j, 1)) = values(j)(1)
            Next j
        Next i
    End With
End Function


Public Function Query(tableName As String, ByVal selectColumns As Variant, Optional ByVal predicates As Variant) As DBRecordset
    Dim i As Long
    Dim j As Long
    Dim rowMatches As New Collection
    Dim tmp As Variant
    Dim vals As Variant
    Dim colNumbers() As Long
    
    selectColumns = ConvertToArray(selectColumns)
    
    tmp = GetTable(tableName).UsedRange.Value
    
    ReDim colNumbers(LBound(selectColumns, 1) To UBound(selectColumns, 1), 1 To 1)
    For i = LBound(selectColumns, 1) To UBound(selectColumns, 1)
        colNumbers(i, 1) = GetColumnNumber(tmp, CStr(selectColumns(i)))
        If colNumbers(i, 1) < 1 Then
            Err.Raise Number:=1, Description:="Column '" & selectColumns(i) & "' not found in table '" & tableName & "'"
        End If
    Next i

    SetupPredicateColumnNumbers tableName, tmp, predicates
    
    For i = 2 To UBound(tmp, 1)
        If RowMatchesPredicates(tmp, i, predicates) Then rowMatches.Add i
    Next i
    
    If rowMatches.Count > 0 Then
        ReDim vals(1 To rowMatches.Count, 1 To UBound(colNumbers, 1) - LBound(colNumbers, 1) + 1)
        For i = 1 To rowMatches.Count
            For j = LBound(colNumbers, 1) To UBound(colNumbers, 1)
                vals(i, 1 + j - LBound(colNumbers, 1)) = tmp(rowMatches(i), colNumbers(j, 1))
            Next j
        Next i
    End If
    
    Set Query = New DBRecordset
    Query.Setup vals, selectColumns
End Function


Private Function RowMatchesPredicates(data As Variant, rowNum As Long, Optional predicates As Variant) As Boolean
    Dim i As Long
    Dim j As Long

    If IsMissing(predicates) = True Then
        RowMatchesPredicates = True
        Exit Function
    End If

    For i = LBound(predicates, 1) To UBound(predicates, 1)
        Select Case predicates(i).Operator
            Case OP.EQUAL
                RowMatchesPredicates = (CStr(data(rowNum, predicates(i).Column)) = predicates(i).Parameter(0))
            Case OP.NOT_EQUAL
                RowMatchesPredicates = (CStr(data(rowNum, predicates(i).Column)) <> predicates(i).Parameter(0))
            Case OP.GREATER_THAN
                RowMatchesPredicates = (data(rowNum, predicates(i).Column) > predicates(i).Parameter(0))
            Case OP.LESS_THAN
                RowMatchesPredicates = (data(rowNum, predicates(i).Column) < predicates(i).Parameter(0))
            Case OP.GREATER_THAN_EQUAL
                RowMatchesPredicates = (data(rowNum, predicates(i).Column) >= predicates(i).Parameter(0))
            Case OP.LESS_THAN_EQUAL
                RowMatchesPredicates = (data(rowNum, predicates(i).Column) <= predicates(i).Parameter(0))
            Case OP.IS_IN
                For j = 0 To predicates(i).ParameterCount() - 1
                    If (CStr(data(rowNum, predicates(i).Column)) = predicates(i).Parameter(j)) Then
                        RowMatchesPredicates = True
                        Exit For
                    End If
                Next j
            Case OP.IS_NOT_IN
                For j = 0 To predicates(i).ParameterCount() - 1
                    If (CStr(data(rowNum, predicates(i).Column)) = predicates(i).Parameter(j)) Then
                        RowMatchesPredicates = False
                        Exit For
                    End If
                Next j
            Case OP.IS_NULL
                RowMatchesPredicates = IsNullValue(data(rowNum, predicates(i).Column))
            Case OP.IS_NOT_NULL
                RowMatchesPredicates = Not IsNullValue(data(rowNum, predicates(i).Column))
            Case OP.BETWEEN
                RowMatchesPredicates = (data(rowNum, predicates(i).Column) >= predicates(i).Parameter(0)) And RowMatchesPredicates = (data(rowNum, predicates(i).Column) <= predicates(i).Parameter(1))
            Case Else
                RowMatchesPredicates = False
        End Select
        If RowMatchesPredicates = False Then Exit For
    Next i
End Function


Public Function Pred(Name As String, ByVal Operator As Variant, Optional Params As Variant) As DBPredicate
    Dim opVal As OP
    Set Pred = New DBPredicate
    If VarType(Operator) = vbString Then
        opVal = OperatorStringToOP(Operator)
    Else
        opVal = Operator
    End If
    Pred.InitiateProperties Name, opVal, Params
End Function


Private Sub SetupPredicateColumnNumbers(tableName As String, data As Variant, predicates As Variant)
    Dim i As Long
    If IsMissing(predicates) = False Then
        predicates = ConvertToArray(predicates)
        For i = LBound(predicates, 1) To UBound(predicates, 1)
            predicates(i).SetColumnNumber GetColumnNumber(data, predicates(i).Name)
            If predicates(i).Column < 1 Then
                Err.Raise Number:=1, Description:="Column '" & predicates(i).Name & "' not found in table '" & tableName & "'"
            End If
        Next i
    End If
End Sub


Private Function GetColumnNumber(tmp As Variant, columnName As Variant) As Long
    Dim i As Long
    
    GetColumnNumber = -1
    
    If VBA.Strings.Len(columnName) = 0 Then Exit Function
    
    For i = 1 To UBound(tmp, 2)
        If tmp(1, i) = columnName Then
            GetColumnNumber = i
            Exit For
        End If
    Next i
End Function


Private Function ToSheetTableName(tableName) As String
    If VBA.Strings.Left(tableName, VBA.Strings.Len(TABLE_PREFIX)) <> TABLE_PREFIX Then
        ToSheetTableName = TABLE_PREFIX & tableName
    Else
        ToSheetTableName = tableName
    End If
End Function


Public Function NullValue() As Variant
    NullValue = Empty
End Function


Private Function IsNullValue(val As Variant) As Boolean
    IsNullValue = (IsNull(val) Or IsEmpty(val))
End Function


Private Function GetTable(tableName As String) As Worksheet
    Set GetTable = ThisWorkbook.Sheets(ToSheetTableName(tableName))
End Function


Public Sub Sort(tableName As String, ByVal columnNames As Variant)
    Dim i As Long
    Dim tmp As Variant
    
    columnNames = ConvertToArray(columnNames)
    
    If ArrayLength(columnNames) < 1 Then Exit Sub
    
    With GetTable(tableName).UsedRange
        tmp = .Value
        For i = LBound(columnNames, 1) To UBound(columnNames, 1)
            columnNames(i) = GetColumnNumber(tmp, columnNames(i))
            If columnNames(i) < 1 Then Err.Raise Number:=1, Description:="Column '" & columnNames(i) & "' does not exist in table '" & tableName & "'"
        Next i
        
        On Error Resume Next ' Just incase
        Select Case ArrayLength(columnNames)
            Case 1
                .Sort Key1:=.Columns(columnNames(0)), Order1:=Excel.XlSortOrder.xlAscending, Header:=Excel.XlYesNoGuess.xlYes
            Case 2
                .Sort Key1:=.Columns(columnNames(0)), Order1:=Excel.XlSortOrder.xlAscending, Key2:=.Columns(columnNames(1)), Order2:=Excel.XlSortOrder.xlAscending, Header:=Excel.XlYesNoGuess.xlYes
            Case Else
                .Sort Key1:=.Columns(columnNames(0)), Order1:=Excel.XlSortOrder.xlAscending, Key2:=.Columns(columnNames(1)), Order2:=Excel.XlSortOrder.xlAscending, Key3:=.Columns(columnNames(2)), Order3:=Excel.XlSortOrder.xlAscending, Header:=Excel.XlYesNoGuess.xlYes
        End Select
    End With
End Sub

