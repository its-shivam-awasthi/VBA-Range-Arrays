VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Class: Arrays
Option Explicit

Public Enum vbMatchMode
    'Text
    vbAutoDetect = -1
    vbTextEquals = 0
    vbTextNotEqual = 1
    vbTextContains = 2
    vbTextDoesNotContains = 3
    vbTextStartsWith = 4
    vbTextEndsWith = 5
    vbTextWildcard = 6
    vbTextRegex = 7
    'Numbers
    vbNumberEquals = 11
    vbNumberNotEqual = 12
    vbNumberGreaterThan = 13
    vbNumberGreaterThanOrEqual = 14
    vbNumberLessThan = 15
    vbNumberLessThanOrEqual = 16
    vbNumberBetween = 17
    'Logic
    VbVarType = 21
    vbIsFunction = 22
End Enum

Public Enum vbConversionType
    vbToDate
    vbToCurrency
    vbToInteger
    vbToLong
    vbToDouble
    vbToString
End Enum

Public Enum vbReplaceOperator
    vbEquals
    vbContains
    vbStartsWith
    vbEndsWith
End Enum

Public Enum vbMathOperator
    vbAdd
    vbSubtract
    vbMultiply
    vbDivide
    vbPower
    vbMod
End Enum

Private Enum Conversion_Method
    vbToRow = 0
    vbToCol = 1
End Enum

Public Enum vbFilterLogic
    vbLogicAnd = 1
    vbLogicOr = 2
End Enum

Public Enum vbUserSortMethod
    vbUserQuickSort = 0
    vbUserMergeSort = 1
End Enum

Private C_Dict As Object
Private C_IndexDict As Object
Private C_Arr As Variant
Private C_Ref As String
Private ArrIndex As ArrayIndexes

Private Sub Class_Initialize()
    Set C_Dict = VBA.CreateObject("Scripting.Dictionary")
    Set C_IndexDict = VBA.CreateObject("Scripting.Dictionary")
    Set ArrIndex = New ArrayIndexes
End Sub

Private Sub Class_Terminate()
    If Not C_Dict Is Nothing Then Set C_Dict = Nothing
    If Not C_IndexDict Is Nothing Then Set C_IndexDict = Nothing
    If Not ArrIndex Is Nothing Then Set ArrIndex = Nothing
End Sub

Public Property Get Arrays() As Object
    Set Arrays = C_Dict
End Property

Public Property Set Arrays(ByRef vNewValue As Object)
    If TypeName(vNewValue) = "Dictionary" Then
        Set C_Dict = vNewValue
    Else
        Err.Raise vbObjectError + 1011, , "Invalid object type. Expected Dictionary."
    End If
End Property

Public Property Get Ref() As Variant
    Ref = C_Ref
End Property

Public Property Get RefArray(Optional Ref As Variant) As Variant
    If IsMissing(Ref) Then
        RefArray = C_Arr
    Else
        If C_Dict.Exists(Ref) Then
            RefArray = C_Dict(Ref)
        Else
            Err.Raise vbObjectError + 1012, , "Reference not found in dictionary: " & Ref
        End If
    End If
End Property

Public Property Get arr(Ref As Variant) As Arrays
    Set arr = UseRef(Ref)
End Property

Public Property Get AllRefs() As Variant
    AllRefs = C_Dict.keys()
End Property

Public Property Get AllArrays() As Variant
    AllArrays = C_Dict.items()
End Property

Public Property Get CurrentRef() As Variant
    CurrentRef = C_Ref
End Property

Public Property Get CurrentArray() As Variant
    CurrentArray = C_Arr
End Property

Public Property Get ArrayCount() As Long
    ArrayCount = C_Dict.Count
End Property

Public Function UseRef(Ref As Variant, Optional newArr As Variant) As Arrays

    If Not C_Dict.Exists(Ref) Then
        Err.Raise vbObjectError + 1010, , "Reference name not found: " & Ref
    End If

    ' Preserve current index if C_Ref was previously set
    If Not IsEmpty(C_Ref) And C_Dict.Exists(C_Ref) Then
        C_IndexDict(C_Ref) = ArrIndex
    End If

    ' Update reference
    C_Ref = Ref

    ' Optionally update array
    If Not IsMissing(newArr) Then
        If IsArray(newArr) And DimensionOfArray(newArr) = 2 Then
            C_Dict(Ref) = newArr
        Else
            Err.Raise vbObjectError + 999, , "Expected a 2D Array or empty Array: " & newArr
        End If
    End If

    ' Load array and index
    C_Arr = C_Dict(Ref)
    If C_IndexDict.Exists(Ref) Then
        ArrIndex = C_IndexDict(Ref)
    Else
        Set ArrIndex = New ArrayIndexes
        ArrIndex.InitialArray = C_Arr
        C_IndexDict.Add Ref, ArrIndex
    End If

    Set UseRef = Me

End Function

Public Property Get FirstRowIndex() As Long
    FirstRowIndex = LBound(C_Arr, 1)
End Property

Public Property Get LastRowIndex() As Long
    LastRowIndex = UBound(C_Arr, 1)
End Property

Public Property Get FirstColIndex() As Long
    FirstColIndex = LBound(C_Arr, 2)
End Property

Public Property Get LastColIndex() As Long
    LastColIndex = UBound(C_Arr, 2)
End Property

Public Property Get Indexes() As ArrayIndexes
    If Not C_Dict.Exists(C_Ref) Then
        Err.Raise vbObjectError + 1020, , "Reference not set."
    End If
    Set Indexes = ArrIndex
End Property

Public Property Get RowCount() As Long
    RowCount = LastRowIndex - FirstRowIndex + 1
End Property

Public Property Get ColumnCount() As Long
    ColumnCount = LastColIndex - FirstColIndex + 1
End Property

' Get Dimension
Public Property Get Dimension() As String
    Dimension = RowCount & "x" & ColumnCount
End Property

Public Property Get BoundsInfo() As String
    BoundsInfo = "Rows: " & LBound(C_Arr, 1) & " to " & UBound(C_Arr, 1) & ", Columns: " & LBound(C_Arr, 2) & " to " & UBound(C_Arr, 2)
End Property

Public Property Get CellValue(ByVal RowIndex As Long, ByVal ColIndex As Long) As Variant
    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2120, , "Array not initialized."
    End If

    If RowIndex < 1 Or RowIndex > UBound(C_Arr, 1) Or ColIndex < 1 Or ColIndex > UBound(C_Arr, 2) Then
        Err.Raise vbObjectError + 2121, , "Cell index out of bounds."
    End If

    CellValue = C_Arr(RowIndex, ColIndex)
End Property

Public Property Let CellValue(ByVal RowIndex As Long, ByVal ColIndex As Long, ByVal NewValue As Variant)
    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2120, , "Array not initialized."
    End If

    If RowIndex < 1 Or RowIndex > UBound(C_Arr, 1) Or ColIndex < 1 Or ColIndex > UBound(C_Arr, 2) Then
        Err.Raise vbObjectError + 2121, , "Cell index out of bounds."
    End If

    C_Arr(RowIndex, ColIndex) = NewValue
End Property

Public Property Get HasInput() As Boolean

    HasInput = (UBound(C_Arr, 1) > 1 Or Not IsEmpty(C_Arr(1, 1)))

End Property

Public Property Get HasEmptyCells() As Boolean
    Dim i As Long, j As Long

    If Not IsArray(C_Arr) Then Exit Property

    On Error Resume Next
    For i = LBound(C_Arr, 1) To UBound(C_Arr, 1)
        For j = LBound(C_Arr, 2) To UBound(C_Arr, 2)
            If IsEmpty(C_Arr(i, j)) Or CStr(C_Arr(i, j)) = "" Then
                HasEmptyCells = True
                Exit Property
            End If
        Next j
    Next i
    On Error GoTo 0
    
End Property

Public Property Get HasErrors() As Boolean
    Dim i As Long, j As Long

    If Not IsArray(C_Arr) Then Exit Property

    On Error Resume Next
    For i = LBound(C_Arr, 1) To UBound(C_Arr, 1)
        For j = LBound(C_Arr, 2) To UBound(C_Arr, 2)
            If IsError(C_Arr(i, j)) Then
                HasErrors = True
                Exit Property
            End If
        Next j
    Next i
    On Error GoTo 0
    
End Property

Public Property Get ErrorsCount(Optional RowIndex As Long, Optional ColIndex As Long) As Long
    
    If Not IsArray(C_Arr) Then Exit Property
    
    If IsMissing(RowIndex) Then RowIndex = LBound(C_Arr, 1) & ":" & UBound(C_Arr, 1)
    If IsMissing(ColIndex) Then ColIndex = LBound(C_Arr, 2) & ":" & UBound(C_Arr, 2)
    
    Dim rowIndexArr() As Variant, colIndexArr() As Variant
    Dim i As Long, j As Long, TotalErrors
    
    rowIndexArr = NormalizeToArray(RowIndex)
    colIndexArr = NormalizeToArray(ColIndex)
    
    On Error Resume Next
    For i = LBound(rowIndexArr, 1) To UBound(rowIndexArr, 1)
        For j = LBound(colIndexArr, 2) To UBound(colIndexArr, 2)
            If IsError(C_Arr(rowIndexArr(i), colIndexArr(j))) Then
                TotalErrors = TotalErrors + 1
            End If
        Next j
    Next i
    On Error GoTo 0
    ErrorsCount = TotalErrors
End Property

Public Function HasDuplicates(Optional Ref As Variant) As Boolean
    Dim OpArr As Variant, dict As Object
    Dim i As Long, j As Long, key As String

    If IsMissing(Ref) Then
        OpArr = C_Arr
    ElseIf C_Dict.Exists(Ref) Then
        OpArr = C_Dict(Ref)
    Else
        Err.Raise vbObjectError + 1012, , "Reference name not found: " & Ref
    End If

    Set dict = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    For i = LBound(OpArr, 1) To UBound(OpArr, 1)
        For j = LBound(OpArr, 2) To UBound(OpArr, 2)
            key = CStr(OpArr(i, j))
            If dict.Exists(key) Then
                HasDuplicates = True
                Exit Function
            Else
                dict.Add key, True
            End If
        Next j
    Next i
    HasDuplicates = False
    On Error GoTo 0
End Function

Public Function IsEmptyArray(Optional Ref As Variant) As Boolean
    Dim OpArr As Variant

    If IsMissing(Ref) Then
        OpArr = C_Arr
    Else
        If C_Dict.Exists(Ref) Then
            OpArr = C_Dict(Ref)
        Else
            Err.Raise vbObjectError + 1010, , "Reference name not found: " & Ref
        End If
    End If

    On Error Resume Next
    IsEmptyArray = Not IsArray(OpArr) Or (UBound(OpArr, 1) < LBound(OpArr, 1)) Or (UBound(OpArr, 2) < LBound(OpArr, 2))
    If Err.Number <> 0 Then
        IsEmptyArray = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Function IsBooleanArray(Optional Ref As Variant) As Boolean
    Dim OpArr As Variant, i As Long, j As Long

    If IsMissing(Ref) Then
        OpArr = C_Arr
    ElseIf C_Dict.Exists(Ref) Then
        OpArr = C_Dict(Ref)
    Else
        Err.Raise vbObjectError + 1016, , "Reference name not found: " & Ref
    End If

    On Error Resume Next
    For i = LBound(OpArr, 1) To UBound(OpArr, 1)
        For j = LBound(OpArr, 2) To UBound(OpArr, 2)
            If VarType(OpArr(i, j)) <> vbBoolean Then
                IsBooleanArray = False
                Exit Function
            End If
        Next j
    Next i
    IsBooleanArray = True
    On Error GoTo 0
End Function

Public Function IsDateArray(Optional Ref As Variant) As Boolean
    Dim OpArr As Variant, i As Long, j As Long

    If IsMissing(Ref) Then
        OpArr = C_Arr
    ElseIf C_Dict.Exists(Ref) Then
        OpArr = C_Dict(Ref)
    Else
        Err.Raise vbObjectError + 1013, , "Reference name not found: " & Ref
    End If

    On Error Resume Next
    For i = LBound(OpArr, 1) To UBound(OpArr, 1)
        For j = LBound(OpArr, 2) To UBound(OpArr, 2)
            If Not IsDate(OpArr(i, j)) Then
                IsDateArray = False
                Exit Function
            End If
        Next j
    Next i
    IsDateArray = True
    On Error GoTo 0
End Function

Public Function IsNumericArray(Optional Ref As Variant) As Boolean
    Dim OpArr As Variant, i As Long, j As Long

    If IsMissing(Ref) Then
        OpArr = C_Arr
    ElseIf C_Dict.Exists(Ref) Then
        OpArr = C_Dict(Ref)
    Else
        Err.Raise vbObjectError + 1011, , "Reference name not found: " & Ref
    End If

    On Error Resume Next
    For i = LBound(OpArr, 1) To UBound(OpArr, 1)
        For j = LBound(OpArr, 2) To UBound(OpArr, 2)
            If Not IsNumeric(OpArr(i, j)) Then
                IsNumericArray = False
                Exit Function
            End If
        Next j
    Next i
    IsNumericArray = True
    On Error GoTo 0
End Function

Public Function IsJaggedArray(ByVal arr As Variant) As Boolean
    Dim i As Long, baseCols As Long

    If Not IsArray(arr) Then
        IsJaggedArray = False
        Exit Function
    End If

    On Error Resume Next
    baseCols = -1
    For i = LBound(arr) To UBound(arr)
        If IsArray(arr(i)) Then
            If baseCols = -1 Then
                baseCols = UBound(arr(i)) - LBound(arr(i)) + 1
            ElseIf (UBound(arr(i)) - LBound(arr(i)) + 1) <> baseCols Then
                IsJaggedArray = True
                Exit Function
            End If
        Else
            IsJaggedArray = True
            Exit Function
        End If
    Next i
    IsJaggedArray = False
    On Error GoTo 0
End Function

Public Function IsTextArray(Optional Ref As Variant) As Boolean
    Dim OpArr As Variant, i As Long, j As Long

    If IsMissing(Ref) Then
        OpArr = C_Arr
    ElseIf C_Dict.Exists(Ref) Then
        OpArr = C_Dict(Ref)
    Else
        Err.Raise vbObjectError + 1015, , "Reference name not found: " & Ref
    End If

    On Error Resume Next
    For i = LBound(OpArr, 1) To UBound(OpArr, 1)
        For j = LBound(OpArr, 2) To UBound(OpArr, 2)
            If Not (VarType(OpArr(i, j)) = vbString) Then
                IsTextArray = False
                Exit Function
            End If
        Next j
    Next i
    IsTextArray = True
    On Error GoTo 0
End Function

Public Sub Update(ByVal NewArray As Variant)
    If Not IsArray(NewArray) Or DimensionOfArray(NewArray) < 2 Then
        Err.Raise vbObjectError + 2000, , "Input must be a 2D array."
    Else
        C_Arr = NewArray
        C_Dict(C_Ref) = C_Arr
    End If
End Sub

Public Sub Clear()
    If C_Dict.Exists(C_Ref) Then
        Dim fi As Long, fj As Long
        Dim tempArr As Variant
        tempArr = C_Dict(C_Ref)
        For fi = LBound(tempArr, 1) To UBound(tempArr, 1)
            For fj = LBound(tempArr, 2) To UBound(tempArr, 2)
                tempArr(fi, fj) = Empty
            Next fj
        Next fi
        
        C_Dict(C_Ref) = tempArr ' Clears the array associated with the reference
    Else
        Err.Raise vbObjectError + 1012, "Clear", "Reference doesn't exist: " & C_Ref
    End If
End Sub

Private Sub Swap(ByRef a As Long, ByRef b As Long)
    Dim temp As Long
    temp = a
    a = b
    b = temp
End Sub

Public Property Get Rows(Optional ByVal Index As Variant, Optional ByVal InvertMatch As Boolean = False) As Collection
    Dim i As Long
    Dim RowIndexes() As Variant
    Dim selectedRows As Object
    Dim ArrIndex As ArrayIndexes
    Dim Coll As New Collection

    ' Default to full range if Index is missing
    If IsMissing(Index) Then
        Index = LBound(C_Arr, 1) & ":" & UBound(C_Arr, 1)
    End If

    ' Normalize input to array
    RowIndexes = NormalizeToArray(Index)

    ' Validate and collect selected row indexes
    Set selectedRows = VBA.CreateObject("Scripting.Dictionary")
    For i = LBound(RowIndexes) To UBound(RowIndexes)
        If RowIndexes(i) < 1 Or RowIndexes(i) > UBound(C_Arr, 1) Then
            Err.Raise vbObjectError + 2001, , "Row index out of bounds."
        End If
        If selectedRows.Exists(RowIndexes(i)) Then
            Err.Raise vbObjectError + 2051, , "Duplicate row indexes."
        Else
            selectedRows.Add RowIndexes(i), True
        End If
    Next i

    ' Collect and build output rows
    For i = 1 To UBound(C_Arr, 1)
        Dim includeRow As Boolean
        includeRow = selectedRows.Exists(i)
        If InvertMatch Then includeRow = Not includeRow

        If includeRow Then
            Set ArrIndex = New ArrayIndexes
            Set ArrIndex.ParentArrays = Me
            ArrIndex.InitialArray = Me.CurrentArray
            ArrIndex.RowIndexes = Array(i)
            Coll.Add ArrIndex
        End If
    Next i

    Set Rows = Coll
End Property

Public Property Get Columns(Optional ByVal Index As Variant, Optional ByVal InvertMatch As Boolean = False) As Collection
    Dim i As Long
    Dim colIndexes() As Variant
    Dim selectedCols As Object
    Dim ArrIndex As ArrayIndexes
    Dim Coll As New Collection

    ' Default to full range if Index is missing
    If IsMissing(Index) Then
        Index = LBound(C_Arr, 2) & ":" & UBound(C_Arr, 2)
    End If

    ' Normalize input to array
    colIndexes = NormalizeToArray(Index)

    ' Validate and collect selected column indexes
    Set selectedCols = VBA.CreateObject("Scripting.Dictionary")
    For i = LBound(colIndexes) To UBound(colIndexes)
        If colIndexes(i) < 1 Or colIndexes(i) > UBound(C_Arr, 2) Then
            Err.Raise vbObjectError + 2002, , "Column index out of bounds."
        End If
        If selectedCols.Exists(colIndexes(i)) Then
            Err.Raise vbObjectError + 2052, , "Duplicate column indexes."
        Else
            selectedCols.Add colIndexes(i), True
        End If
    Next i

    ' Collect and build output columns
    For i = 1 To UBound(C_Arr, 2)
        Dim includeCol As Boolean
        includeCol = selectedCols.Exists(i)
        If InvertMatch Then includeCol = Not includeCol

        If includeCol Then
            Set ArrIndex = New ArrayIndexes
            Set ArrIndex.ParentArrays = Me
            ArrIndex.InitialArray = Me.CurrentArray
            ArrIndex.ColumnIndexes = Array(i)
            Coll.Add ArrIndex
        End If
    Next i

    Set Columns = Coll
End Property

Public Property Get Region(startIndex As Variant, endIndex As Variant) As Arrays
    Dim i As Long, j As Long
    Dim ArrIndex As New ArrayIndexes
    Dim RowIndexes As Variant, colIndexes As Variant
    Dim startRow As Long, startCol As Long, endRow As Long, endCol As Long

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2122, , "Array not initialized."
    End If

    If Not (IsArray(startIndex) And IsArray(endIndex)) Then
        Err.Raise vbObjectError + 2125, , "Expected 1D arrays of 2 elements."
    End If
    If DimensionOfArray(startIndex) > 1 Or DimensionOfArray(endIndex) > 1 Then
        Err.Raise vbObjectError + 2125, , "Expected 1D arrays of 2 elements."
    End If

    If UBound(startIndex) <> 1 Or UBound(endIndex) <> 1 Then
        Err.Raise vbObjectError + 2125, , "Expected 1D arrays of 2 elements."
    End If

    startRow = startIndex(0): startCol = startIndex(1)
    endRow = endIndex(0): endCol = endIndex(1)

    If startRow < 1 Or endRow > UBound(C_Arr, 1) Or startCol < 1 Or endCol > UBound(C_Arr, 2) Then
        Err.Raise vbObjectError + 2123, , "Region bounds out of range."
    End If

    RowIndexes = NormalizeToArray(startRow & ":" & endRow)
    colIndexes = NormalizeToArray(startCol & ":" & endCol)

    Set ArrIndex.ParentArrays = Me
    ArrIndex.InitialArray = Me.CurrentArray
    ArrIndex.RowIndexes = RowIndexes
    ArrIndex.ColumnIndexes = colIndexes
    
    Set Region = ArrIndex
    
End Property

Public Property Get Areas(ByVal regionSpecs As Variant) As Collection
    Dim i As Long
    Dim Region As Variant
    Dim ArrIndex As ArrayIndexes
    Dim Coll As New Collection
    Dim RowIndexes As Variant, colIndexes As Variant
    Dim startRow As Long, startCol As Long, endRow As Long, endCol As Long

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2130, , "Array not initialized."
    End If

    If Not IsArray(regionSpecs) Then
        Err.Raise vbObjectError + 2131, , "Areas expects an array of region specifications."
    End If

    For i = LBound(regionSpecs) To UBound(regionSpecs)
        Region = regionSpecs(i)
        If Not IsArray(Region) Or UBound(Region) <> 3 Then
            Err.Raise vbObjectError + 2132, , "Each region must be an array of 4 elements: startRow,startCol,endRow,endCol."
        End If

        startRow = Region(0)
        startCol = Region(1)
        endRow = Region(2)
        endCol = Region(3)
        If startRow < 1 Or endRow > UBound(C_Arr, 1) Or startCol < 1 Or endCol > UBound(C_Arr, 2) Then
            Err.Raise vbObjectError + 2133, , "Region bounds out of range."
        End If

        RowIndexes = NormalizeToArray(startRow & ":" & endRow)
        colIndexes = NormalizeToArray(startCol & ":" & endCol)

        Set ArrIndex = New ArrayIndexes
        Set ArrIndex.ParentArrays = Me
        ArrIndex.InitialArray = Me.CurrentArray
        ArrIndex.RowIndexes = RowIndexes
        ArrIndex.ColumnIndexes = colIndexes
        Coll.Add ArrIndex
    Next i

    Set Areas = Coll
    
End Property

Public Sub Filter(criteria As Variant, Field As Long, _
    Optional FilterMethod As XlSearchOrder = xlByRows, _
    Optional matchmode As Variant = vbTextEquals, _
    Optional CompareMethod As VbCompareMethod = vbBinaryCompare, _
    Optional Not_Found As Variant)

    Dim i As Long, j As Long, n As Long
    Dim ReturnArray As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2022, , "Array Object Not Set or is empty."
    End If

    Select Case FilterMethod
        Case xlByRows
            For i = LBound(C_Arr, 1) To UBound(C_Arr, 1)
                If IsMatch(C_Arr(i, Field), criteria, matchmode, CompareMethod) Then n = n + 1: dict.Add n, i
            Next i

            If n > 0 Then
                ReDim ReturnArray(1 To dict.Count, 1 To ColumnCount)
                For i = 1 To n
                    For j = 1 To ColumnCount
                        ReturnArray(i, j) = C_Arr(dict(i), j)
                    Next j
                Next i
            End If

        Case xlByColumns
            For j = LBound(C_Arr, 2) To UBound(C_Arr, 2)
                If IsMatch(C_Arr(Field, j), criteria, matchmode, CompareMethod) Then n = n + 1: dict.Add n, j
            Next j

            If n > 0 Then
                ReDim ReturnArray(1 To RowCount, 1 To dict.Count)
                For i = 1 To RowCount
                    For j = 1 To n
                        ReturnArray(i, j) = C_Arr(i, dict(j))
                    Next j
                Next i
            End If

        Case Else
            Err.Raise vbObjectError + 2023, , "Invalid Filter Method"
    End Select

    ' Handle result
    If C_Ref <> "" Then
        If n = 0 Then
            If IsMissing(Not_Found) Then
                C_Arr = Array()
            Else
                C_Arr = Not_Found
            End If
        Else
            C_Arr = ReturnArray
        End If
    End If
    C_Dict(C_Ref) = C_Arr
End Sub

Private Function ParseCondition(ByVal condition As String, ByRef op As String, ByRef operand As Variant) As Boolean
    Dim operators As Variant, i As Long
    operators = Array("<=", ">=", "<>", "=", "<", ">")

    For i = LBound(operators) To UBound(operators)
        If left(condition, Len(operators(i))) = operators(i) Then
            op = operators(i)
            operand = Trim(mid(condition, Len(op) + 1))
            ParseCondition = True
            Exit Function
        End If
    Next i

    ParseCondition = False
    
End Function

Private Function IsMatch(ByVal Value As Variant, _
                         ByVal criteria As Variant, _
                         ByVal matchmode As Variant, _
                         Optional CompareMethod As VbCompareMethod = vbBinaryCompare, _
                         Optional reCache As Object = Nothing) As Boolean
    Dim strValue As String, strCriteria As String
    Dim i As Long
    
    ' Null safety
    If IsNull(Value) Then
        strValue = ""
    Else
        strValue = CStr(Value)
    End If
    
    ' Handle multiple criteria (OR logic)
    If IsArray(criteria) Then
        If matchmode = vbNumberBetween Then
            If UBound(criteria) = 1 Then
                IsMatch = (Value >= criteria(0) And Value <= criteria(1))
                Exit Function
            Else
                IsMatch = False
                Exit Function
            End If
        End If
        
        For i = LBound(criteria) To UBound(criteria)
            Dim singleCriteria As Variant
            singleCriteria = criteria(i)
            If IsMatch(Value, singleCriteria, matchmode, CompareMethod, reCache) Then
                IsMatch = True
                Exit Function
            End If
        Next i
        IsMatch = False
        Exit Function
    End If
    
    ' Convert single criteria safely
    If IsNull(criteria) Then
        strCriteria = ""
    Else
        strCriteria = CStr(criteria)
    End If
    
    ' Auto-detect logic
    If matchmode = vbAutoDetect And VarType(criteria) = vbString Then
        Dim op As String, operand As Variant
        If ParseCondition(strCriteria, op, operand) Then
            Select Case op
                Case "=": IsMatch = (Value = operand)
                Case "<>": IsMatch = (Value <> operand)
                Case ">": IsMatch = (Value > operand)
                Case "<": IsMatch = (Value < operand)
                Case ">=": IsMatch = (Value >= operand)
                Case "<=": IsMatch = (Value <= operand)
            End Select
            Exit Function
        End If
        matchmode = IIf(IsNumeric(Value), vbNumberEquals, vbTextEquals)
    End If
    
    ' Main logic using Select Case True for speed
    Select Case True
        Case matchmode = vbTextEquals
            IsMatch = (StrComp(strValue, strCriteria, CompareMethod) = 0)
        Case matchmode = vbTextNotEqual
            IsMatch = (StrComp(strValue, strCriteria, CompareMethod) <> 0)
        Case matchmode = vbTextContains
            IsMatch = (InStr(1, strValue, strCriteria, CompareMethod) > 0)
        Case matchmode = vbTextDoesNotContains
            IsMatch = (InStr(1, strValue, strCriteria, CompareMethod) = 0)
        Case matchmode = vbTextStartsWith
            IsMatch = (Len(strValue) >= Len(strCriteria) And _
                       StrComp(left(strValue, Len(strCriteria)), strCriteria, CompareMethod) = 0)
        Case matchmode = vbTextEndsWith
            IsMatch = (Len(strValue) >= Len(strCriteria) And _
                       StrComp(right(strValue, Len(strCriteria)), strCriteria, CompareMethod) = 0)
        Case matchmode = vbTextWildcard
            IsMatch = (strValue Like strCriteria)
        Case matchmode = vbTextRegex
            Dim re As Object
            If reCache Is Nothing Then
                Set re = CreateObject("VBScript.RegExp")
                re.Pattern = strCriteria
                re.IgnoreCase = (CompareMethod = vbTextCompare)
                re.Global = False
                Set reCache = re
            Else
                Set re = reCache
            End If
            IsMatch = re.Test(strValue)
        Case matchmode = vbNumberEquals
            IsMatch = (Value = criteria)
        Case matchmode = vbNumberNotEqual
            IsMatch = (Value <> criteria)
        Case matchmode = vbNumberGreaterThan
            IsMatch = (Value > criteria)
        Case matchmode = vbNumberGreaterThanOrEqual
            IsMatch = (Value >= criteria)
        Case matchmode = vbNumberLessThan
            IsMatch = (Value < criteria)
        Case matchmode = vbNumberLessThanOrEqual
            IsMatch = (Value <= criteria)
        Case matchmode = VbVarType
            IsMatch = (VarType(Value) = criteria)
        Case matchmode = vbIsFunction
            Select Case LCase(strCriteria)
                Case "isdate": IsMatch = IsDate(Value)
                Case "isempty": IsMatch = IsEmpty(Value)
                Case "iserror": IsMatch = IsError(Value)
                Case "isnull": IsMatch = IsNull(Value)
                Case "isnumeric": IsMatch = IsNumeric(Value)
                Case Else: IsMatch = False
            End Select
        Case Else
            IsMatch = False
    End Select
End Function

Public Sub TransFielde()
    
    Dim i As Long, j As Long
    Dim transFieldedArr() As Variant
    
    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2024, , "Array not initialized."
    End If
    
    ReDim transFieldedArr(LBound(C_Arr, 2) To UBound(C_Arr, 2), LBound(C_Arr, 1) To UBound(C_Arr, 1))
    
    For i = LBound(C_Arr, 1) To UBound(C_Arr, 1)
        For j = LBound(C_Arr, 2) To UBound(C_Arr, 2)
            transFieldedArr(j, i) = C_Arr(i, j)
        Next j
    Next i
    
    ' Update internal array and dictionary
    C_Arr = transFieldedArr
    If C_Ref <> "" Then
        C_Dict(C_Ref) = C_Arr
    End If
    
End Sub

Private Function DimensionOfArray(ArrData As Variant) As Long
    Dim n As Long
    On Error GoTo Done
    Do
        n = n + 1
        Dim temp As Long
        temp = UBound(ArrData, n)
    Loop
Done:
    DimensionOfArray = n - 1
End Function

Private Function Convert2D(ArrData As Variant, ConvertBy As Conversion_Method) As Variant
    Dim newArr() As Variant
    Dim i As Long, j As Long

    If IsArray(ArrData) Then
        If DimensionOfArray(ArrData) = 1 Then
            Select Case ConvertBy
                Case vbToRow
                    ReDim newArr(1 To 1, LBound(ArrData) To UBound(ArrData))
                    For j = LBound(ArrData) To UBound(ArrData)
                        newArr(1, j) = ArrData(j)
                    Next j
                Case vbToCol
                    ReDim newArr(LBound(ArrData) To UBound(ArrData), 1 To 1)
                    For i = LBound(ArrData) To UBound(ArrData)
                        newArr(i, 1) = ArrData(i)
                    Next i
                Case Else
                    Err.Raise vbObjectError + 2030, , "Unsupported conversion method."
            End Select
        Else
            newArr = ArrData
        End If
    Else
        Err.Raise vbObjectError + 2029, , "Input is not an array."
    End If

    Convert2D = newArr
End Function

Public Sub AppendRows(ByVal NewData As Variant)
    Dim newArr() As Variant
    Dim i As Long, j As Long
    Dim newRows As Long, newCols As Long
    Dim totalRows As Long, totalCols As Long
    Dim resultArr() As Variant
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    
    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2026, , "Target array is not initialized."
    End If
    
    If Not IsArray(NewData) Then
        ReDim newArr(1 To 1, UBound(C_Arr, 2) To UBound(C_Arr, 2))
        For j = LBound(C_Arr, 2) To UBound(C_Arr, 2)
            newArr(1, j) = NewData
        Next j
    Else
        newArr = Convert2D(NewData, vbToRow)
    End If

    ' Validate column count
    newCols = UBound(newArr, 2) - LBound(newArr, 2) + 1
    totalCols = UBound(C_Arr, 2) - LBound(C_Arr, 2) + 1

    If newCols <> totalCols Then
        Err.Raise vbObjectError + 2028, , "Column count mismatch. Expected: " & totalCols & ", Got: " & newCols
    End If

    ' Prepare result array
    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)

    totalRows = r2 - r1 + 1
    newRows = UBound(newArr, 1) - LBound(newArr, 1) + 1

    ReDim resultArr(1 To totalRows + newRows, c1 To c2)

    ' Copy existing data
    For i = 1 To totalRows
        For j = c1 To c2
            resultArr(i, j) = C_Arr(r1 + i - 1, j)
        Next j
    Next i

    ' Append new data
    For i = 1 To newRows
        For j = c1 To c2
            resultArr(totalRows + i, j) = newArr(LBound(newArr, 1) + i - 1, j)
        Next j
    Next i
    ' Update internal array and dictionary
    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Sub AppendCols(ByVal NewData As Variant)
    Dim newArr() As Variant
    Dim i As Long, j As Long
    Dim newCols As Long, newRows As Long
    Dim totalCols As Long, totalRows As Long
    Dim resultArr() As Variant
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2030, , "Target array is not initialized."
    End If
    
    If Not IsArray(NewData) Then
        ReDim newArr(LBound(C_Arr, 1) To UBound(C_Arr, 1), 1 To 1)
        For i = LBound(C_Arr, 1) To UBound(C_Arr, 1)
            newArr(i, 1) = NewData
        Next i
    Else
        newArr = Convert2D(NewData, vbToCol)
    End If
    
    ' Validate row count
    newRows = UBound(newArr, 1) - LBound(newArr, 1) + 1
    totalRows = UBound(C_Arr, 1) - LBound(C_Arr, 1) + 1

    If newRows <> totalRows Then
        Err.Raise vbObjectError + 2032, , "Row count mismatch. Expected: " & totalRows & ", Got: " & newRows
    End If

    totalCols = UBound(C_Arr, 2) - LBound(C_Arr, 2) + 1
    newCols = UBound(newArr, 2) - LBound(newArr, 2) + 1

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)

    ReDim resultArr(r1 To r2, c1 To c2 + newCols)

    ' Copy existing data
    For i = r1 To r2
        For j = c1 To c2
            resultArr(i, j) = C_Arr(i, j)
        Next j
    Next i

    ' Append new columns
    For i = r1 To r2
        For j = 1 To newCols
            resultArr(i, c2 + j) = newArr(i, LBound(newArr, 2) + j - 1)
        Next j
    Next i

    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Sub Union(Ref_Args As Variant)
    If IsArray(Ref_Args) Then
        Dim i As Long
        For i = LBound(Ref_Args) To UBound(Ref_Args)
            If Not C_Dict.Exists(Ref_Args(i)) Then Err.Raise vbObjectError + 2100, , "Some or all refereces are not valid."
        Next i
    
        For i = LBound(Ref_Args) To UBound(Ref_Args)
            Call AppendRows(C_Dict(Ref_Args(i)))
        Next i
    Else
        If Not C_Dict.Exists(Ref_Args) Then Err.Raise vbObjectError + 2101, , "Reference is not valid"
        Call AppendRows(C_Dict(Ref_Args))
    End If
End Sub

Public Sub InsertRow(ByVal Index As Long, ByVal newRow As Variant)
    Dim i As Long, j As Long
    Dim totalRows As Long, totalCols As Long, insertAt As Long
    Dim resultArr() As Variant, NewRowArr As Variant
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    Dim n As Long

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2033, , "Target array is not initialized."
    End If

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)

    totalRows = r2 - r1 + 1
    totalCols = c2 - c1 + 1

    If Index < 1 Or Index > totalRows + 1 Then
        Err.Raise vbObjectError + 2034, , "Row index out of bounds."
    End If

    If Not IsArray(newRow) Then
        ReDim NewRowArr(1 To 1, c1 To c2)
        For j = LBound(C_Arr, 2) To UBound(C_Arr, 2)
            NewRowArr(1, j) = newRow
        Next j
    Else
        NewRowArr = Convert2D(newRow, vbToRow)
    End If

    n = UBound(NewRowArr, 1) - LBound(NewRowArr, 1) + 1

    If UBound(NewRowArr, 2) - LBound(NewRowArr, 2) + 1 <> totalCols Then
        Err.Raise vbObjectError + 2035, , "New row must match column count."
    End If

    ReDim resultArr(r1 To r2 + n, c1 To c2)
    
    insertAt = r1 + Index - 1
    ' Copy rows before insertion point
    For i = r1 To insertAt - 1
        For j = c1 To c2
            resultArr(i, j) = C_Arr(i, j)
        Next j
    Next i

    ' Insert new rows
    For i = 0 To n - 1
        For j = c1 To c2
            resultArr(insertAt + i, j) = NewRowArr(LBound(NewRowArr, 1) + i, j)
        Next j
    Next i

    ' Copy remaining rows
    For i = insertAt + n To r2 + n
        For j = c1 To c2
            resultArr(i, j) = C_Arr(i - n, j)
        Next j
    Next i

    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Sub InsertCol(ByVal Index As Long, ByVal newCol As Variant)
    Dim i As Long, j As Long
    Dim totalRows As Long, totalCols As Long, InsertColAt
    Dim resultArr() As Variant, NewColArr As Variant
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    Dim n As Long
    
    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2036, , "Target array is not initialized."
    End If

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)

    totalRows = r2 - r1 + 1
    totalCols = c2 - c1 + 1

    If Index < 1 Or Index > totalCols + 1 Then
        Err.Raise vbObjectError + 2037, , "Column index out of bounds."
    End If
    
    If Not IsArray(newCol) Then
        ReDim NewColArr(r1 To r2, 1 To 1)
        For i = LBound(C_Arr, 1) To UBound(C_Arr, 1)
            NewColArr(i, 1) = newCol
        Next i
    Else
        NewColArr = Convert2D(newCol, vbToCol)
    End If
    
    n = UBound(NewColArr, 2) - LBound(NewColArr, 2) + 1

    If UBound(NewColArr, 1) - LBound(NewColArr, 1) + 1 <> totalRows Then
        Err.Raise vbObjectError + 2038, , "New column must match row count."
    End If

    ReDim resultArr(r1 To r2, c1 To c2 + n)
    
    InsertColAt = c1 + Index - 1
    ' Copy columns before insertion point
    For i = r1 To r2
        For j = c1 To InsertColAt - 1
            resultArr(i, j) = C_Arr(i, j)
        Next j
    Next i

    ' Insert new columns
    For i = r1 To r2
        For j = 0 To n - 1
            resultArr(i, InsertColAt + j) = NewColArr(i, LBound(NewColArr, 2) + j)
        Next j
    Next i

    ' Copy remaining columns
    For i = r1 To r2
        For j = InsertColAt + n To c2 + n
            resultArr(i, j) = C_Arr(i, j - n)
        Next j
    Next i

    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = resultArr
End Sub

Public Sub InsertMultipleRows(Indexes As Variant, RowData As Variant)
    Dim IndexArr As Variant
    Dim i As Long, j As Long, k As Long
    Dim totalRows As Long, totalCols As Long
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    Dim insertCount As Long
    Dim resultArr() As Variant
    Dim insertMap As Object: Set insertMap = CreateObject("Scripting.Dictionary")
    Dim RowCount As Long
    
    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2050, , "Target array is not initialized."
    End If

    IndexArr = NormalizeToArray(Indexes)

    If UBound(IndexArr) <> UBound(RowData) Then
        Err.Raise vbObjectError + 2051, , "Indexes and RowData arrays must have the same length."
    End If

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)
    totalRows = r2 - r1 + 1
    totalCols = c2 - c1 + 1

    ' Normalize and validate all insert columns
    insertCount = UBound(IndexArr) - LBound(IndexArr) + 1
    ReDim IndexeArrays(LBound(IndexArr) To UBound(IndexArr))

    For i = LBound(IndexArr) To UBound(IndexArr)
        Dim idx As Long: idx = IndexArr(i)
        If idx < 1 Or idx > totalRows + 1 Then
            Err.Raise vbObjectError + 2052, , "Row index out of bounds at position " & i & "."
        End If

        Dim rowArr As Variant
        rowArr = Convert2D(IIf(IsArray(RowData(i)), RowData(i), ValToArray(RowData(i), totalRows, False)), vbToCol)

        If UBound(rowArr, 1) - LBound(rowArr, 1) + 1 <> totalCols Then
            Err.Raise vbObjectError + 2053, , "Rows " & i & " does not match column count."
        End If

        RowCount = UBound(rowArr, 1) - LBound(rowArr, 1) + 1
        insertCount = insertCount + RowCount - 1
        IndexeArrays = Array(idx, rowArr, RowCount)
        insertMap.Add i, IndexeArrays
    Next i

    ' Prepare result array
    ReDim resultArr(r1 To r2 + insertCount, c1 To c2)

    Dim insertData As Variant, insertAt As Long, InsertRow As Variant, rowOffset As Long
    For i = 0 To insertMap.Count - 1

        insertData = insertMap(i)
        insertAt = CLng(insertData(0))
        InsertRow = insertData(1)
        rowOffset = Get_Offset(insertAt, i + 1, IndexArr, insertMap)
        For k = LBound(InsertRow, 1) To UBound(InsertRow, 1)
            For j = c1 To c2
                resultArr(i + 1 + rowOffset, j) = InsertRow(k, j)
            Next j
        Next k
    Next i

    For i = r1 To r2
        rowOffset = Get_Offset(i, 0, IndexArr, insertMap)
        For j = c1 To c2
            resultArr(i + rowOffset, j) = C_Arr(i, j)
        Next j
    Next i

    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr

End Sub

Public Sub InsertMultipleCols(Indexes As Variant, ColData As Variant)
    Dim IndexArr As Variant
    Dim i As Long, j As Long, k As Long
    Dim totalRows As Long, totalCols As Long
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    Dim insertCount As Long
    Dim resultArr() As Variant
    Dim IndexeArrays() As Variant
    Dim insertMap As Object: Set insertMap = CreateObject("Scripting.Dictionary")
    Dim colCount As Long

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2050, , "Target array is not initialized."
    End If

    IndexArr = NormalizeToArray(Indexes)

    If UBound(IndexArr) <> UBound(ColData) Then
        Err.Raise vbObjectError + 2051, , "Indexes and ColData arrays must have the same length."
    End If

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)
    totalRows = r2 - r1 + 1
    totalCols = c2 - c1 + 1

    ' Normalize and validate all insert columns
    insertCount = UBound(IndexArr) - LBound(IndexArr) + 1
    ReDim IndexeArrays(LBound(IndexArr) To UBound(IndexArr))

    For i = LBound(IndexArr) To UBound(IndexArr)
        Dim idx As Long: idx = IndexArr(i)
        If idx < 1 Or idx > totalCols + 1 Then
            Err.Raise vbObjectError + 2052, , "Column index out of bounds at position " & i & "."
        End If

        Dim colArr As Variant
        colArr = Convert2D(IIf(IsArray(ColData(i)), ColData(i), ValToArray(ColData(i), totalRows, False)), vbToCol)

        If UBound(colArr, 1) - LBound(colArr, 1) + 1 <> totalRows Then
            Err.Raise vbObjectError + 2053, , "Column " & i & " does not match row count."
        End If

        colCount = UBound(colArr, 2) - LBound(colArr, 2) + 1
        insertCount = insertCount + colCount - 1
        IndexeArrays = Array(idx, colArr, colCount)
        insertMap.Add i, IndexeArrays
    Next i

    ' Prepare result array
    ReDim resultArr(r1 To r2, c1 To c2 + insertCount)

    Dim insertData As Variant, insertAt As Long, InsertCol As Variant, colOffset As Long
    For j = 0 To insertMap.Count - 1

        insertData = insertMap(j)
        insertAt = CLng(insertData(0))
        InsertCol = insertData(1)
        colOffset = Get_Offset(insertAt, j + 1, IndexArr, insertMap)
        For i = r1 To r2
            For k = LBound(InsertCol, 2) To UBound(InsertCol, 2)
                resultArr(i, j + 1 + colOffset) = InsertCol(i, k)
            Next k
        Next i
    Next j

    For j = c1 To c2
        colOffset = Get_Offset(j, 0, IndexArr, insertMap)
        For i = r1 To r2
            resultArr(i, j + colOffset) = C_Arr(i, j)
        Next i
    Next j

    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr

End Sub

Private Function Get_Offset(Index As Long, StartPos As Long, IndexArr As Variant, insertMap As Object)

    Dim i As Long, m As Long: m = 0
    For i = LBound(IndexArr) + StartPos To UBound(IndexArr)
        m = m + IIf(Index >= IndexArr(i), insertMap(i)(2), 0)
    Next i

    Get_Offset = m

End Function

'Public Sub InsertMultipleRows(Indexes As Variant, RowData As Variant)
'    Dim i As Long, j As Long, r As Long, c As Long
'    Dim totalRows As Long, totalCols As Long
'    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
'    Dim insertCount As Long
'    Dim resultArr() As Variant
'    Dim insertList() As Variant
'
'    If Not IsArray(C_Arr) Then
'        Err.Raise vbObjectError + 2060, , "Target array is not initialized."
'    End If
'
'    Dim IndexArr As Variant
'    IndexArr = NormalizeToArray(Indexes)
'
'    If UBound(IndexArr) <> UBound(RowData) Then
'        Err.Raise vbObjectError + 2061, , "Indexes and RowData arrays must have the same length."
'    End If
'
'    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
'    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)
'    totalRows = r2 - r1 + 1
'    totalCols = c2 - c1 + 1
'
'    insertCount = 0
'    ReDim insertList(1 To UBound(IndexArr) - LBound(IndexArr) + 1, 1 To 3)
'
'    ' Collect insert info
'    For i = LBound(IndexArr) To UBound(IndexArr)
'        Dim idx As Long: idx = IndexArr(i)
'        If idx < 1 Or idx > totalRows + 1 Then
'            Err.Raise vbObjectError + 2062, , "Row index out of bounds at position " & i & "."
'        End If
'
'        Dim rowArr As Variant
'        rowArr = Convert2D(IIf(IsArray(RowData(i)), RowData(i), ValToArray(RowData(i), totalCols, True)), vbToRow)
'
'        If UBound(rowArr, 2) - LBound(rowArr, 2) + 1 <> totalCols Then
'            Err.Raise vbObjectError + 2063, , "Row " & i & " does not match column count."
'        End If
'
'        Dim rowCount As Long
'        rowCount = UBound(rowArr, 1) - LBound(rowArr, 1) + 1
'        insertCount = insertCount + rowCount
'
'        insertList(i + 1, 1) = idx
'        insertList(i + 1, 2) = rowArr
'        insertList(i + 1, 3) = rowCount
'    Next i
'
'    ' Sort insertList by index (stable sort)
'    QuickSortInsertList insertList, 1, UBound(insertList, 1)
'
'    ' Prepare result array
'    ReDim resultArr(r1 To r2 + insertCount, c1 To c2)
'
'    Dim srcRow As Long, destRow As Long, insPtr As Long
'    srcRow = r1
'    destRow = r1
'    insPtr = 1
'
'    For j = 1 To totalRows + 1
'        ' Insert all rows scheduled before this original row
'        Do While insPtr <= UBound(insertList, 1) And insertList(insPtr, 1) = j
'            Dim newRow As Variant
'            newRow = insertList(insPtr, 2)
'            For r = LBound(newRow, 1) To UBound(newRow, 1)
'                For c = c1 To c2
'                    resultArr(destRow, c) = newRow(r, c)
'                Next c
'                destRow = destRow + 1
'            Next r
'            insPtr = insPtr + 1
'        Loop
'
'        ' Copy original row if within bounds
'        If j <= totalRows Then
'            For c = c1 To c2
'                resultArr(destRow, c) = C_Arr(srcRow, c)
'            Next c
'            srcRow = srcRow + 1
'            destRow = destRow + 1
'        End If
'    Next j
'
'    C_Arr = resultArr
'    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
'End Sub
'
'Public Sub InsertMultipleCols(Indexes As Variant, ColData As Variant)
'    Dim i As Long, j As Long, r As Long, c As Long
'    Dim totalRows As Long, totalCols As Long
'    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
'    Dim insertCount As Long
'    Dim resultArr() As Variant
'    Dim insertList() As Variant
'
'    If Not IsArray(C_Arr) Then
'        Err.Raise vbObjectError + 2050, , "Target array is not initialized."
'    End If
'
'    Dim IndexArr As Variant
'    IndexArr = NormalizeToArray(Indexes)
'
'    If UBound(IndexArr) <> UBound(ColData) Then
'        Err.Raise vbObjectError + 2051, , "Indexes and ColData arrays must have the same length."
'    End If
'
'    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
'    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)
'    totalRows = r2 - r1 + 1
'    totalCols = c2 - c1 + 1
'
'    insertCount = 0
'    ReDim insertList(1 To UBound(IndexArr) - LBound(IndexArr) + 1, 1 To 3)
'
'    ' Collect insert info
'    For i = LBound(IndexArr) To UBound(IndexArr)
'        Dim idx As Long: idx = IndexArr(i)
'        If idx < 1 Or idx > totalCols + 1 Then
'            Err.Raise vbObjectError + 2052, , "Column index out of bounds at position " & i & "."
'        End If
'
'        Dim colArr As Variant
'        colArr = Convert2D(IIf(IsArray(ColData(i)), ColData(i), ValToArray(ColData(i), totalRows, False)), vbToCol)
'
'        If UBound(colArr, 1) - LBound(colArr, 1) + 1 <> totalRows Then
'            Err.Raise vbObjectError + 2053, , "Column " & i & " does not match row count."
'        End If
'
'        Dim colCount As Long
'        colCount = UBound(colArr, 2) - LBound(colArr, 2) + 1
'        insertCount = insertCount + colCount
'
'        insertList(i + 1, 1) = idx
'        insertList(i + 1, 2) = colArr
'        insertList(i + 1, 3) = colCount
'    Next i
'
'    ' Sort insertList by index (stable sort)
'    QuickSortInsertList insertList, 1, UBound(insertList, 1)
'
'    ' Prepare result array
'    ReDim resultArr(r1 To r2, c1 To c2 + insertCount)
'
'    Dim srcCol As Long, destCol As Long, insPtr As Long
'    srcCol = c1
'    destCol = c1
'    insPtr = 1
'
'    For j = 1 To totalCols + 1
'        ' Insert all columns scheduled before this original column
'        Do While insPtr <= UBound(insertList, 1) And insertList(insPtr, 1) = j
'            Dim newCol As Variant
'            newCol = insertList(insPtr, 2)
'            For c = LBound(newCol, 2) To UBound(newCol, 2)
'                For r = r1 To r2
'                    resultArr(r, destCol) = newCol(r, c)
'                Next r
'                destCol = destCol + 1
'            Next c
'            insPtr = insPtr + 1
'        Loop
'
'        ' Copy original column if within bounds
'        If j <= totalCols Then
'            For r = r1 To r2
'                resultArr(r, destCol) = C_Arr(r, srcCol)
'            Next r
'            srcCol = srcCol + 1
'            destCol = destCol + 1
'        End If
'    Next j
'
'    C_Arr = resultArr
'    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
'End Sub
'
'Private Sub QuickSortInsertList(ByRef arr() As Variant, low As Long, high As Long)
'    Dim i As Long, j As Long, pivot As Long
'    Dim temp1 As Variant, temp2 As Variant, temp3 As Variant
'    If low < high Then
'        pivot = arr((low + high) \ 2, 1)
'        i = low: j = high
'        Do While i <= j
'            Do While arr(i, 1) < pivot: i = i + 1: Loop
'            Do While arr(j, 1) > pivot: j = j - 1: Loop
'            If i <= j Then
'                temp1 = arr(i, 1): temp2 = arr(i, 2): temp3 = arr(i, 3)
'                arr(i, 1) = arr(j, 1): arr(i, 2) = arr(j, 2): arr(i, 3) = arr(j, 3)
'                arr(j, 1) = temp1: arr(j, 2) = temp2: arr(j, 3) = temp3
'                i = i + 1: j = j - 1
'            End If
'        Loop
'        If low < j Then QuickSortInsertList arr, low, j
'        If i < high Then QuickSortInsertList arr, i, high
'    End If
'End Sub

Public Sub DeleteRows(ByVal Index As Variant)
    Dim i As Long, j As Long, n As Long
    Dim totalRows As Long, totalCols As Long
    Dim resultArr() As Variant, IndexArr As Variant
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    Dim dict As Object

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2039, , "Target array is not initialized."
    End If

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)

    totalRows = r2 - r1 + 1
    totalCols = c2 - c1 + 1

    IndexArr = NormalizeToArray(Index)

    If DimensionOfArray(IndexArr) > 1 Then
        Err.Raise vbObjectError + 2101, , "Expected an array of or Long/Integer Variable"
    End If

    Set dict = VBA.CreateObject("Scripting.Dictionary")
    For i = LBound(IndexArr) To UBound(IndexArr)
        If Not IsNumeric(IndexArr(i)) Then
            Err.Raise vbObjectError + 2101, , "Expected an array of or Long/Integer Variable"
        ElseIf IndexArr(i) < 1 Or IndexArr(i) > totalRows Then
            Err.Raise vbObjectError + 2040, , "Row index out of bounds."
        Else
            If Not dict.Exists(IndexArr(i)) Then dict.Add CLng(IndexArr(i)), True
        End If
    Next i

    ReDim resultArr(r1 To r2 - dict.Count, c1 To c2)

    For i = r1 To r2
        If Not dict.Exists(i - r1 + 1) Then
            For j = c1 To c2
                resultArr(i - n, j) = C_Arr(i, j)
            Next j
        Else
            n = n + 1
        End If
    Next i

    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Sub DeleteCols(ByVal Index As Variant)
    Dim i As Long, j As Long, n As Long
    Dim totalRows As Long, totalCols As Long
    Dim resultArr() As Variant, IndexArr As Variant
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    Dim dict As Object

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2039, , "Target array is not initialized."
    End If

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)

    totalRows = r2 - r1 + 1
    totalCols = c2 - c1 + 1

    IndexArr = NormalizeToArray(Index)

    If DimensionOfArray(IndexArr) > 1 Then
        Err.Raise vbObjectError + 2101, , "Expected an array of or Long/Integer Variable"
    End If

    Set dict = VBA.CreateObject("Scripting.Dictionary")
    For i = LBound(IndexArr) To UBound(IndexArr)
        If Not IsNumeric(IndexArr(i)) Then
            Err.Raise vbObjectError + 2101, , "Expected an array of or Long/Integer Variable"
        ElseIf IndexArr(i) < 1 Or IndexArr(i) > totalCols Then
            Err.Raise vbObjectError + 2040, , "Column index out of bounds."
        Else
            If Not dict.Exists(IndexArr(i)) Then dict.Add CLng(IndexArr(i)), True
        End If
    Next i

    ReDim resultArr(r1 To r2, c1 To c2 - dict.Count)

    For i = r1 To r2
        n = 0
        For j = c1 To c2
            If Not dict.Exists(j - c1 + 1) Then
                resultArr(i, j - n) = C_Arr(i, j)
            Else
                n = n + 1
            End If
        Next j
    Next i

    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Function ChooseRows(ByVal RowIndices As Variant) As Arrays
    Dim i As Long, j As Long, r As Long
    Dim totalCols As Long
    Dim resultArr() As Variant
    Dim IndexArr As Variant
    Dim arrObj As New Arrays
    Dim dict As Object
    
    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2043, , "Array not initialized."
    End If

    IndexArr = NormalizeToArray(RowIndices)

    If DimensionOfArray(IndexArr) > 1 Then
        Err.Raise vbObjectError + 2101, , "Expected a 1D array of row indices."
    End If
    totalCols = UBound(C_Arr, 2) - LBound(C_Arr, 2) + 1
    ReDim resultArr(1 To UBound(IndexArr) - LBound(IndexArr) + 1, 1 To totalCols)

    For i = LBound(IndexArr) To UBound(IndexArr)
        r = IndexArr(i)
        If r < LBound(C_Arr, 1) Or r > UBound(C_Arr, 1) Then
            Err.Raise vbObjectError + 2045, , "Row index out of bounds: " & r
        End If
        For j = LBound(C_Arr, 2) To UBound(C_Arr, 2)
            resultArr(i - LBound(IndexArr) + 1, j - LBound(C_Arr, 2) + 1) = C_Arr(r, j)
        Next j
    Next i
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add C_Ref, resultArr
    
    Set arrObj.Arrays = dict
    arrObj.UseRef C_Ref
    Set ChooseRows = arrObj

End Function

Public Function ChooseCols(ByVal ColIndices As Variant) As Arrays
    Dim i As Long, j As Long, c As Long
    Dim totalRows As Long
    Dim resultArr() As Variant
    Dim IndexArr As Variant
    Dim arrObj As New Arrays
    Dim dict As Object
    
    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2046, , "Array not initialized."
    End If

    IndexArr = NormalizeToArray(ColIndices)

    If DimensionOfArray(IndexArr) > 1 Then
        Err.Raise vbObjectError + 2101, , "Expected a 1D array of column indices."
    End If

    totalRows = UBound(C_Arr, 1) - LBound(C_Arr, 1) + 1
    ReDim resultArr(1 To totalRows, 1 To UBound(IndexArr) - LBound(IndexArr) + 1)

    For j = LBound(IndexArr) To UBound(IndexArr)
        c = IndexArr(j)
        If c < LBound(C_Arr, 2) Or c > UBound(C_Arr, 2) Then
            Err.Raise vbObjectError + 2048, , "Column index out of bounds: " & c
        End If
        For i = LBound(C_Arr, 1) To UBound(C_Arr, 1)
            resultArr(i - LBound(C_Arr, 1) + 1, j - LBound(IndexArr) + 1) = C_Arr(i, c)
        Next i
    Next j
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add C_Ref, resultArr
    
    Set arrObj.Arrays = dict
    arrObj.UseRef C_Ref
    Set ChooseCols = arrObj

End Function

Private Function NormalizeToArray(Index As Variant) As Variant
    Dim i As Long, Var As Variant
    Dim VarComma() As String, VarColon() As String
    Dim startVal As Long, endVal As Long, NewLength As Long
    Dim returnArr() As Variant, oldUBound As Long

    Select Case VarType(Index)
        Case vbArray + vbInteger, vbArray + vbLong
            NormalizeToArray = Index

        Case vbInteger, vbLong
            NormalizeToArray = Array(Index)

        Case vbString
            If IsNumeric(Index) Then
                NormalizeToArray = Array(CLng(Index))
            ElseIf InStr(Index, ":") > 0 Or InStr(Index, ",") > 0 Then
                VarComma = Split(Index, ",")
                Erase returnArr ' Start with empty array

                For Each Var In VarComma
                    VarColon = Split(Var, ":")
                    If UBound(VarColon) <> 1 Or Not (IsNumeric(VarColon(0)) And IsNumeric(VarColon(1))) Then
                        Err.Raise vbObjectError + 1017, , "Reference Index not valid."
                    End If

                    startVal = CLng(VarColon(0)): endVal = CLng(VarColon(1))
                    If startVal > endVal Then Swap startVal, endVal

                    oldUBound = 0
                    If Not IsArray(returnArr) Then oldUBound = UBound(returnArr)
                    NewLength = endVal - startVal + 1
                    ReDim Preserve returnArr(0 To oldUBound + NewLength - 1)

                    For i = 0 To NewLength - 1
                        returnArr(oldUBound + i) = startVal + i
                    Next i
                Next Var

                NormalizeToArray = returnArr
            Else
                Err.Raise vbObjectError + 1017, , "Reference Index not valid."
            End If

        Case vbVariant, vbArray + vbVariant
            If IsArray(Index) Then
                For i = LBound(Index) To UBound(Index)
                    If Not IsNumeric(Index(i)) Then Err.Raise vbObjectError + 2102, , "Unsupported Index type"
                Next i
                NormalizeToArray = Index
            ElseIf IsNumeric(Index) Then
                NormalizeToArray = Array(Index)
            Else
                Err.Raise vbObjectError + 2102, , "Unsupported Index type"
            End If

        Case Else
            Err.Raise vbObjectError + 2102, , "Unsupported Index type"
    End Select
End Function

'Public Function NormalizeCriteria(ByVal colIndexArr As Variant, ByVal criteria As Variant, ByVal totalRows As Long) As Variant
'    Dim criteriaArr As Variant
'    Dim i As Long, totalIndexes As Long, totalCriteria As Long
'
'    totalIndexes = UBound(colIndexArr) - LBound(colIndexArr) + 1
'
'    If Not IsArray(criteria) Then
'        criteriaArr = Array(criteria)
'
'    ElseIf DimensionOfArray(criteria) = 1 Then
'        totalCriteria = UBound(criteria) - LBound(criteria) + 1
'        If totalIndexes <> totalCriteria Then
'            Err.Raise vbObjectError + 2061, , "colIndex and criteria must be of equal length."
'        End If
'        criteriaArr = criteria
'
'    ElseIf DimensionOfArray(criteria) = 2 Then
'        If totalIndexes <> 1 Then
'            Err.Raise vbObjectError + 2063, , "Multi-row criteria only allowed for single colIndex."
'        End If
'        If UBound(criteria, 1) - LBound(criteria, 1) + 1 <> totalRows Then
'            Err.Raise vbObjectError + 2063, , "Criteria array must match row count."
'        End If
'        criteriaArr = criteria
'
'    Else
'        Err.Raise vbObjectError + 2063, , "Invalid criteria format."
'    End If
'
'    NormalizeCriteria = criteriaArr
'End Function

Private Function NormalizeMatchMode(ByVal colIndexArr As Variant, ByVal matchmode As Variant) As Variant
    Dim MatchModeArr() As Variant
    Dim j As Long

    If IsArray(matchmode) Then
        If UBound(matchmode) - LBound(matchmode) <> UBound(colIndexArr) - LBound(colIndexArr) Then
            Err.Raise vbObjectError + 2090, , "MatchMode array must be same length as colIndex and Criteria."
        End If
        MatchModeArr = matchmode
    Else
        ReDim MatchModeArr(LBound(colIndexArr) To UBound(colIndexArr))
        For j = LBound(colIndexArr) To UBound(colIndexArr)
            MatchModeArr(j) = matchmode
        Next j
    End If

    NormalizeMatchMode = MatchModeArr
End Function

Public Function CountIfs(ByVal ColIndex As Variant, _
                            ByVal criteria As Variant, _
                            Optional ByVal matchmode As Variant = vbTextEquals, _
                            Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Long
                         
    Dim i As Long, j As Long
    Dim matchCount As Long
    Dim totalRows As Long, totalCols As Long
    Dim colIndexArr As Variant, criteriaArr, MatchModeArr As Variant

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2049, , "Array not initialized."
    End If
    
    totalRows = UBound(C_Arr, 1) - LBound(C_Arr, 2) + 1
    totalCols = UBound(C_Arr, 2) - LBound(C_Arr, 2) + 1
    
    colIndexArr = NormalizeToArray(ColIndex)
    criteriaArr = IIf(IsArray(criteria), criteria, Array(criteria))
    MatchModeArr = NormalizeMatchMode(colIndexArr, matchmode)
    
    If DimensionOfArray(colIndexArr) > 1 Or DimensionOfArray(criteriaArr) > 1 Then
        Err.Raise vbObjectError + 2050, , "Both colIndex and criteria must be 1D arrays or a pair of Single Variables."
    End If

    If UBound(colIndexArr) - LBound(colIndexArr) <> UBound(criteriaArr) - LBound(criteriaArr) Then
        Err.Raise vbObjectError + 2051, , "colIndex and criteria must be of equal length."
    End If

    For i = 1 To totalRows
        Dim Bool As Boolean: Bool = True
        For j = LBound(colIndexArr) To UBound(colIndexArr)
            Dim Col As Long: Col = colIndexArr(j)
            If Col < 1 Or Col > totalCols Then
                Err.Raise vbObjectError + 2052, , "Column index out of bounds: " & Col
            End If
            If Not IsMatch(C_Arr(i, Col), criteriaArr(j), MatchModeArr(j), CompareMethod) Then
                Bool = False
                Exit For
            End If
        Next j
        If Bool Then matchCount = matchCount + 1
    Next i

    CountIfs = matchCount
End Function

Public Function SumIfs(ByVal ColIndex As Variant, _
                       ByVal criteria As Variant, _
                       ByVal SumCol As Long, _
                       Optional ByVal matchmode As Variant = vbTextEquals, _
                       Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Double
    Dim i As Long, j As Long
    Dim totalRows As Long, totalCols As Long
    Dim colIndexArr As Variant, criteriaArr As Variant, MatchModeArr As Variant
    Dim sumTotal As Double

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2053, , "Array not initialized."

    colIndexArr = NormalizeToArray(ColIndex)
    criteriaArr = IIf(IsArray(criteria), criteria, Array(criteria))
    MatchModeArr = NormalizeMatchMode(colIndexArr, matchmode)
    
    If DimensionOfArray(colIndexArr) > 1 Or DimensionOfArray(criteriaArr) > 1 Then
        Err.Raise vbObjectError + 2054, , "Both colIndex and criteria must be 1D arrays or a pair of Single Variables."
    End If

    If UBound(colIndexArr) - LBound(colIndexArr) <> UBound(criteriaArr) - LBound(criteriaArr) Then
        Err.Raise vbObjectError + 2055, , "colIndex and criteria arrays must be of equal length."
    End If

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)

    If SumCol < 1 Or SumCol > totalCols Then
        Err.Raise vbObjectError + 2056, , "Sum column index out of bounds: " & SumCol
    End If

    For i = 1 To totalRows
        Dim Bool As Boolean: Bool = True
        For j = LBound(colIndexArr) To UBound(colIndexArr)
            If Not IsMatch(C_Arr(i, colIndexArr(j)), criteriaArr(j), MatchModeArr(j), CompareMethod) Then
                Bool = False
                Exit For
            End If
        Next j
        If Bool Then sumTotal = sumTotal + CDbl(C_Arr(i, SumCol))
    Next i

    SumIfs = sumTotal
End Function

Public Function MaxIfs(ByVal ColIndex As Variant, _
                       ByVal criteria As Variant, _
                       ByVal targetCol As Long, _
                       Optional ByVal matchmode As Variant = vbTextEquals, _
                       Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Double
    Dim i As Long, j As Long
    Dim totalRows As Long, totalCols As Long
    Dim colIndexArr As Variant, criteriaArr As Variant, MatchModeArr As Variant
    Dim maxVal As Double
    Dim firstMatch As Boolean

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2057, , "Array not initialized."

    colIndexArr = NormalizeToArray(ColIndex)
    criteriaArr = IIf(IsArray(criteria), criteria, Array(criteria))
    MatchModeArr = NormalizeMatchMode(colIndexArr, matchmode)
    
    If DimensionOfArray(colIndexArr) > 1 Or DimensionOfArray(criteriaArr) > 1 Then
        Err.Raise vbObjectError + 2058, , "Both colIndex and criteria must be 1D arrays or a pair of Single Variables."
    End If

    If UBound(colIndexArr) - LBound(colIndexArr) <> UBound(criteriaArr) - LBound(criteriaArr) Then
        Err.Raise vbObjectError + 2059, , "colIndex and criteria arrays must be of equal length."
    End If

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)

    If targetCol < 1 Or targetCol > totalCols Then
        Err.Raise vbObjectError + 2060, , "Target column index out of bounds: " & targetCol
    End If

    firstMatch = False
    For i = 1 To totalRows
        Dim Bool As Boolean: Bool = True
        For j = LBound(colIndexArr) To UBound(colIndexArr)
            If Not IsMatch(C_Arr(i, colIndexArr(j)), criteriaArr(j), matchmode, CompareMethod) Then
                Bool = False
                Exit For
            End If
        Next j
        If Bool Then
            If Not firstMatch Then
                maxVal = CDbl(C_Arr(i, targetCol))
                firstMatch = True
            ElseIf CDbl(C_Arr(i, targetCol)) > maxVal Then
                maxVal = CDbl(C_Arr(i, targetCol))
            End If
        End If
    Next i

    MaxIfs = maxVal
End Function

Public Function MinIfs(ByVal ColIndex As Variant, _
                       ByVal criteria As Variant, _
                       ByVal targetCol As Long, _
                       Optional ByVal matchmode As Variant = vbTextEquals, _
                       Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Double
    Dim i As Long, j As Long
    Dim totalRows As Long, totalCols As Long
    Dim colIndexArr As Variant, criteriaArr As Variant, MatchModeArr As Variant
    Dim minVal As Double
    Dim firstMatch As Boolean

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2061, , "Array not initialized."

    colIndexArr = NormalizeToArray(ColIndex)
    criteriaArr = IIf(IsArray(criteria), criteria, Array(criteria))
    MatchModeArr = NormalizeMatchMode(colIndexArr, matchmode)
    
    If DimensionOfArray(colIndexArr) > 1 Or DimensionOfArray(criteriaArr) > 1 Then
        Err.Raise vbObjectError + 2062, , "Both colIndex and criteria must be 1D arrays or a pair of Single Variables."
    End If

    If UBound(colIndexArr) - LBound(colIndexArr) <> UBound(criteriaArr) - LBound(criteriaArr) Then
        Err.Raise vbObjectError + 2063, , "colIndex and criteria arrays must be of equal length."
    End If
    
    ' Normalize MatchMode
    If IsArray(matchmode) Then
        MatchModeArr = matchmode
        If UBound(MatchModeArr) - LBound(MatchModeArr) <> UBound(colIndexArr) - LBound(colIndexArr) Then
            Err.Raise vbObjectError + 2090, , "MatchMode array must be same length as colIndex and Criteria."
        End If
    Else
        ReDim MatchModeArr(LBound(colIndexArr) To UBound(colIndexArr))
        For j = LBound(MatchModeArr) To UBound(MatchModeArr)
            MatchModeArr(j) = matchmode
        Next j
    End If
    
    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)

    If targetCol < 1 Or targetCol > totalCols Then
        Err.Raise vbObjectError + 2064, , "Target column index out of bounds: " & targetCol
    End If

    firstMatch = False
    For i = 1 To totalRows
        Dim Bool As Boolean: Bool = True
        For j = LBound(colIndexArr) To UBound(colIndexArr)
            If Not IsMatch(C_Arr(i, colIndexArr(j)), criteriaArr(j), MatchModeArr(j), CompareMethod) Then
                Bool = False
                Exit For
            End If
        Next j
        If Bool Then
            If Not firstMatch Then
                minVal = CDbl(C_Arr(i, targetCol))
                firstMatch = True
            ElseIf CDbl(C_Arr(i, targetCol)) < minVal Then
                minVal = CDbl(C_Arr(i, targetCol))
            End If
        End If
    Next i

    MinIfs = minVal
End Function

Public Function AverageIfs(ByVal ColIndex As Variant, _
                           ByVal criteria As Variant, _
                           ByVal targetCol As Long, _
                           Optional matchmode As Variant = vbTextEquals, _
                           Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Double
    Dim i As Long, j As Long, Count As Long
    Dim totalRows As Long, totalCols As Long
    Dim colIndexArr As Variant, criteriaArr As Variant, MatchModeArr As Variant
    Dim sumTotal As Double

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2086, , "Array not initialized."

    colIndexArr = NormalizeToArray(ColIndex)
    criteriaArr = IIf(IsArray(criteria), criteria, Array(criteria))
    MatchModeArr = NormalizeMatchMode(colIndexArr, matchmode)
    
    If DimensionOfArray(colIndexArr) > 1 Or DimensionOfArray(criteriaArr) > 1 Then
        Err.Raise vbObjectError + 2087, , "colIndex and criteria must be 1D arrays."
    End If

    If UBound(colIndexArr) - LBound(colIndexArr) <> UBound(criteriaArr) - LBound(criteriaArr) Then
        Err.Raise vbObjectError + 2088, , "colIndex and criteria arrays must be of equal length."
    End If

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)

    If targetCol < 1 Or targetCol > totalCols Then
        Err.Raise vbObjectError + 2089, , "Target column index out of bounds: " & targetCol
    End If

    For i = 1 To totalRows
        Dim Bool As Boolean: Bool = True
        For j = LBound(colIndexArr) To UBound(colIndexArr)
            If Not IsMatch(C_Arr(i, colIndexArr(j)), criteriaArr(j), MatchModeArr(j), CompareMethod) Then
                Bool = False
                Exit For
            End If
        Next j
        If Bool Then
            sumTotal = sumTotal + CDbl(C_Arr(i, targetCol))
            Count = Count + 1
        End If
    Next i

    If Count = 0 Then
        AverageIfs = CVErr(xlErrDiv0)
    Else
        AverageIfs = sumTotal / Count
    End If
End Function

Public Function FilterIndexes(ByVal ColIndex As Variant, _
                              ByVal criteria As Variant, _
                              Optional ByVal logic As vbFilterLogic = vbLogicAnd, _
                              Optional ByVal matchmode As Variant = vbTextEquals, _
                              Optional CompareMethod As VbCompareMethod = vbBinaryCompare, _
                              Optional InvertMatch As Boolean = False) As Variant
    Dim i As Long, j As Long
    Dim totalRows As Long
    Dim colIndexArr As Variant, criteriaArr As Variant, MatchModeArr As Variant
    Dim resultArr() As Long, matchCount As Long

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2079, , "Array not initialized."

    colIndexArr = NormalizeToArray(ColIndex)
    criteriaArr = IIf(IsArray(criteria), criteria, Array(criteria))
    MatchModeArr = NormalizeMatchMode(colIndexArr, matchmode)

    If DimensionOfArray(colIndexArr) > 1 Or DimensionOfArray(criteriaArr) > 1 Then
        Err.Raise vbObjectError + 2080, , "Both colIndex and criteria must be 1D arrays."
    End If

    If UBound(colIndexArr) - LBound(colIndexArr) <> UBound(criteriaArr) - LBound(criteriaArr) Then
        Err.Raise vbObjectError + 2081, , "colIndex and criteria arrays must be of equal length."
    End If

    totalRows = UBound(C_Arr, 1)
    ReDim resultArr(1 To totalRows)

    For i = 1 To totalRows
        Dim Bool As Boolean
        Select Case logic
            Case vbLogicAnd
                Bool = True
                For j = LBound(colIndexArr) To UBound(colIndexArr)
                    If Not IsMatch(C_Arr(i, colIndexArr(j)), criteriaArr(j), MatchModeArr(j), CompareMethod) Then
                        Bool = False: Exit For
                    End If
                Next j
            Case vbLogicOr
                Bool = False
                For j = LBound(colIndexArr) To UBound(colIndexArr)
                    If IsMatch(C_Arr(i, colIndexArr(j)), criteriaArr(j), MatchModeArr(j), CompareMethod) Then
                        Bool = True: Exit For
                    End If
                Next j
            Case Else
                Err.Raise vbObjectError + 2082, , "Invalid logic type."
        End Select

        If InvertMatch Then Bool = Not Bool

        If Bool Then
            matchCount = matchCount + 1
            resultArr(matchCount) = i
        End If
    Next i

    If matchCount = 0 Then
        FilterIndexes = Array()
    Else
        ReDim Preserve resultArr(1 To matchCount)
        FilterIndexes = resultArr
    End If
End Function

Public Function FilterIfs(ByVal ColIndex As Variant, _
                          ByVal criteria As Variant, _
                          Optional ByVal logic As vbFilterLogic = vbLogicAnd, _
                          Optional ByVal matchmode As Variant = vbTextEquals, _
                          Optional CompareMethod As VbCompareMethod = vbBinaryCompare, _
                          Optional InvertMatch As Boolean = False) As Arrays
    Dim i As Long, j As Long
    Dim totalCols As Long
    Dim resultArr() As Variant
    Dim matchIndexes As Variant
    Dim dict As Object
    Dim arrObj As Arrays

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2065, , "Array not initialized."

    matchIndexes = FilterIndexes(ColIndex, criteria, logic, matchmode, CompareMethod, InvertMatch)

    totalCols = UBound(C_Arr, 2)

    If UBound(matchIndexes) >= 0 Then
        ReDim resultArr(1 To UBound(matchIndexes), 1 To totalCols)
        For i = 1 To UBound(matchIndexes)
            For j = 1 To totalCols
                resultArr(i, j) = C_Arr(matchIndexes(i), j)
            Next j
        Next i
    Else
        ReDim resultArr(1 To 1, 1 To totalCols)
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add C_Ref, resultArr

    Set arrObj = New Arrays
    Set arrObj.Arrays = dict
    arrObj.UseRef C_Ref

    Set FilterIfs = arrObj
    
End Function

Public Sub ConvertColumns(ByVal ColIndex As Variant, ByVal conversionType As vbConversionType)
    Dim i As Long, j As Long
    Dim totalRows As Long, totalCols As Long
    Dim colIndexArr As Variant
    Dim ConversionErrors As Collection
    Dim Col As Long, originalValue As Variant

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2069, , "Array not initialized."
    End If
    colIndexArr = NormalizeToArray(ColIndex)

    If DimensionOfArray(colIndexArr) > 1 Then
        Err.Raise vbObjectError + 2070, , "colIndex must be a 1D array or a Long/Integer variable."
    End If

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)

    Set ConversionErrors = New Collection

    For j = LBound(colIndexArr) To UBound(colIndexArr)
        Col = colIndexArr(j)
        If Col < 1 Or Col > totalCols Then
            Err.Raise vbObjectError + 2071, , "Column index out of bounds: " & Col
        End If

        For i = 1 To totalRows
            originalValue = C_Arr(i, Col)
            On Error GoTo ConversionFailed
            Select Case conversionType
                Case vbToDate:     C_Arr(i, Col) = CDate(originalValue)
                Case vbToCurrency: C_Arr(i, Col) = FormatCurrency(originalValue)
                Case vbToInteger:  C_Arr(i, Col) = CInt(originalValue)
                Case vbToLong:     C_Arr(i, Col) = CLng(originalValue)
                Case vbToDouble:   C_Arr(i, Col) = CDbl(originalValue)
                Case vbToString:   C_Arr(i, Col) = CStr(originalValue)
                Case Else
                    Err.Raise vbObjectError + 2072, , "Invalid conversion type."
            End Select
            On Error GoTo 0
            GoTo ContinueLoop

ConversionFailed:
            ConversionErrors.Add "Row " & i & ", Col " & Col & ": '" & originalValue & "'"
            Resume ContinueLoop

ContinueLoop:
        Next i
    Next j
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr

    ' Optional: Display or handle errors
    If ConversionErrors.Count > 0 Then
        Dim msg As String: msg = "Conversion failed for the following cells:" & vbCrLf
        For Each originalValue In ConversionErrors
            msg = msg & originalValue & vbCrLf
        Next
        MsgBox msg, vbExclamation, "Conversion Errors"
    End If
End Sub

Public Sub FormatColumns(ByVal ColIndex As Variant, ByVal formatString As String)
    Dim i As Long, j As Long
    Dim totalRows As Long, totalCols As Long
    Dim colIndexArr As Variant
    Dim FormatErrors As Collection
    Dim Col As Long, originalValue As Variant

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2073, , "Array not initialized."
    End If

    colIndexArr = NormalizeToArray(ColIndex)

    If DimensionOfArray(colIndexArr) > 1 Then
        Err.Raise vbObjectError + 2074, , "colIndex must be a 1D array or a Long/Integer variable."
    End If

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)

    Set FormatErrors = New Collection

    For j = LBound(colIndexArr) To UBound(colIndexArr)
        Col = colIndexArr(j)
        If Col < 1 Or Col > totalCols Then
            Err.Raise vbObjectError + 2075, , "Column index out of bounds: " & Col
        End If

        For i = 1 To totalRows
            originalValue = C_Arr(i, Col)
            On Error GoTo FormatFailed
            C_Arr(i, Col) = Format(originalValue, formatString)
            On Error GoTo 0
            GoTo ContinueLoop

FormatFailed:
            FormatErrors.Add "Row " & i & ", Col " & Col & ": '" & originalValue & "'"
            Resume ContinueLoop

ContinueLoop:
        Next i
    Next j

    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr

    ' Optional: Display or handle errors
    If FormatErrors.Count > 0 Then
        Dim msg As String: msg = "Formatting failed for the following cells:" & vbCrLf
        For Each originalValue In FormatErrors
            msg = msg & originalValue & vbCrLf
        Next
        MsgBox msg, vbExclamation, "Formatting Errors"
    End If
End Sub

Private Function IsReplacable(FindWithin As String, findValue As String, operatorType As vbReplaceOperator, compareMode As VbCompareMethod) As Boolean

    Select Case operatorType
        Case vbEquals
            IsReplacable = (StrComp(FindWithin, findValue, compareMode) = 0)
        Case vbContains
            IsReplacable = (InStr(1, FindWithin, findValue, compareMode) > 0)
        Case vbStartsWith
            IsReplacable = (StrComp(left(FindWithin, Len(findValue)), findValue, compareMode) = 0)
        Case vbEndsWith
            IsReplacable = (StrComp(right(FindWithin, Len(findValue)), findValue, compareMode) = 0)
        Case Else
            Err.Raise vbObjectError + 2079, , "Invalid replace operator."
    End Select
    
End Function

Public Sub ReplaceInColumns(ByVal ColIndex As Variant, _
                             ByVal findValues As Variant, _
                             ByVal replaceValue As Variant, _
                             TrimBeforeReplace As Boolean, _
                             Optional ReplaceWithin As Boolean = False, _
                             Optional ByVal operatorType As vbReplaceOperator = vbEquals, _
                             Optional compareMode As VbCompareMethod = vbBinaryCompare)

    Dim i As Long, j As Long, k As Long
    Dim totalRows As Long, totalCols As Long
    Dim colIndexArr As Variant, findArr As Variant
    Dim ReplaceErrors As Collection
    Dim Col As Long, CellValue As String, searchValue As String

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2076, , "Array not initialized."
    colIndexArr = NormalizeToArray(ColIndex)
    findArr = IIf(IsArray(findValues), findValues, Array(findValues))

    If DimensionOfArray(colIndexArr) > 1 Then
        Err.Raise vbObjectError + 2077, , "colIndex must be a 1D array or a Long/Integer variable."
    End If

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)

    Set ReplaceErrors = New Collection

    For j = LBound(colIndexArr) To UBound(colIndexArr)
        Col = colIndexArr(j)
        If Col < 1 Or Col > totalCols Then
            Err.Raise vbObjectError + 2078, , "Column index out of bounds: " & Col
        End If

        For i = 1 To totalRows
            On Error GoTo ReplaceFailed
            CellValue = CStr(C_Arr(i, Col))
            If TrimBeforeReplace Then CellValue = Trim(CellValue)

            ' Loop through all findValues
            For k = LBound(findArr) To UBound(findArr)
                searchValue = CStr(findArr(k))

                If IsReplacable(CellValue, searchValue, operatorType, compareMode) Then
                    If ReplaceWithin Then
                        C_Arr(i, Col) = Replace(CellValue, searchValue, CStr(replaceValue), , , compareMode)
                    Else
                        C_Arr(i, Col) = replaceValue
                    End If
                    Exit For ' Stop after first match
                End If
            Next k

            On Error GoTo 0
ContinueLoop:
        Next i
    Next j

    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr

    ' Optional: Display or handle errors
    If ReplaceErrors.Count > 0 Then
        Dim msg As String: msg = "Replacement failed for the following cells:" & vbCrLf
        Dim FailedCells As Variant
        For Each FailedCells In ReplaceErrors
            msg = msg & FailedCells & vbCrLf
        Next
        MsgBox msg, vbExclamation, "Replacement Errors"
    End If
    Exit Sub

ReplaceFailed:
    ReplaceErrors.Add "Row " & i & ", Col " & Col & ": '" & C_Arr(i, Col) & "'"
    Resume ContinueLoop
End Sub

Public Sub ReplaceInRows(ByVal RowIndex As Variant, _
                       ByVal findValues As Variant, _
                       ByVal replaceValue As Variant, _
                       Optional TrimBeforeReplace As Boolean = False, _
                       Optional ReplaceWithin As Boolean = False, _
                       Optional ByVal operatorType As vbReplaceOperator = vbEquals, _
                       Optional compareMode As VbCompareMethod = vbTextCompare)

    Dim i As Long, j As Long, k As Long
    Dim totalRows As Long, totalCols As Long
    Dim rowIndexArr As Variant, findArr As Variant
    Dim ReplaceErrors As Collection
    Dim CellValue As String, searchValue As String

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2076, , "Array not initialized."
    rowIndexArr = NormalizeToArray(RowIndex)
    findArr = IIf(IsArray(findValues), findValues, Array(findValues))

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)

    Set ReplaceErrors = New Collection

    For i = LBound(rowIndexArr) To UBound(rowIndexArr)
        Dim rowNum As Long
        rowNum = rowIndexArr(i)
        If rowNum < 1 Or rowNum > totalRows Then
            Err.Raise vbObjectError + 2080, , "Row index out of bounds: " & rowNum
        End If

        For j = 1 To totalCols
            On Error GoTo ReplaceFailed
            CellValue = CStr(C_Arr(rowNum, j))
            If TrimBeforeReplace Then CellValue = Trim(CellValue)

            For k = LBound(findArr) To UBound(findArr)
                searchValue = CStr(findArr(k))

                If IsReplacable(CellValue, searchValue, operatorType, compareMode) Then
                    If ReplaceWithin Then
                        C_Arr(rowNum, j) = Replace(CellValue, searchValue, CStr(replaceValue), , , compareMode)
                    Else
                        C_Arr(rowNum, j) = replaceValue
                    End If
                    Exit For
                End If
            Next k

            On Error GoTo 0
ContinueLoop:
        Next j
    Next i

    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr

    If ReplaceErrors.Count > 0 Then
        Dim msg As String: msg = "Replacement failed for the following cells:" & vbCrLf
        Dim FailedCells As Variant
        For Each FailedCells In ReplaceErrors
            msg = msg & FailedCells & vbCrLf
        Next
        MsgBox msg, vbExclamation, "Replacement Errors"
    End If
    Exit Sub

ReplaceFailed:
    ReplaceErrors.Add "Row " & rowNum & ", Col " & j & ": '" & C_Arr(rowNum, j) & "'"
    Resume ContinueLoop
End Sub

Public Sub ReplaceAll(ByVal findValues As Variant, _
                      ByVal replaceValue As Variant, _
                      Optional TrimBeforeReplace As Boolean = False, _
                      Optional ReplaceWithin As Boolean = False, _
                      Optional ByVal operatorType As vbReplaceOperator = vbEquals, _
                      Optional compareMode As VbCompareMethod = vbTextCompare)

    Dim i As Long, j As Long, k As Long
    Dim totalRows As Long, totalCols As Long
    Dim findArr As Variant
    Dim ReplaceErrors As Collection
    Dim CellValue As String, searchValue As String

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2076, , "Array not initialized."
    findArr = IIf(IsArray(findValues), findValues, Array(findValues))

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)

    Set ReplaceErrors = New Collection

    For i = 1 To totalRows
        For j = 1 To totalCols
            On Error GoTo ReplaceFailed
            CellValue = CStr(C_Arr(i, j))
            If TrimBeforeReplace Then CellValue = Trim(CellValue)

            For k = LBound(findArr) To UBound(findArr)
                searchValue = CStr(findArr(k))

                If IsReplacable(CellValue, searchValue, operatorType, compareMode) Then
                    If ReplaceWithin Then
                        C_Arr(i, j) = Replace(CellValue, searchValue, CStr(replaceValue), , , compareMode)
                    Else
                        C_Arr(i, j) = replaceValue
                    End If
                    Exit For
                End If
            Next k

            On Error GoTo 0
ContinueLoop:
        Next j
    Next i

    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr

    If ReplaceErrors.Count > 0 Then
        Dim msg As String: msg = "Replacement failed for the following cells:" & vbCrLf
        Dim FailedCells As Variant
        For Each FailedCells In ReplaceErrors
            msg = msg & FailedCells & vbCrLf
        Next
        MsgBox msg, vbExclamation, "Replacement Errors"
    End If
    Exit Sub

ReplaceFailed:
    ReplaceErrors.Add "Row " & i & ", Col " & j & ": '" & C_Arr(i, j) & "'"
    Resume ContinueLoop
End Sub

Public Sub AddCalculatedColumn(Field As Long, _
                               ByRef colIndexes() As Variant, _
                               ByVal operatorType As vbMathOperator)

    Dim i As Long, j As Long, n As Long
    Dim totalRows As Long, totalCols As Long
    Dim resultArr() As Variant
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long, InsertColAt As Long

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2080, , "Array not initialized."
    End If

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)

    totalRows = r2 - r1 + 1
    totalCols = c2 - c1 + 1

    If Field < 1 Or Field > totalCols + 1 Then
        Err.Raise vbObjectError + 2063, , "Field is out of Column Bound: " & Field
    End If

    InsertColAt = c1 + Field - 1
    ReDim resultArr(r1 To r2, c1 To c2 + 1)

    For i = r1 To r2
        ' Copy columns before the insert Fieldition
        For j = c1 To InsertColAt - 1
            resultArr(i, j) = C_Arr(i, j)
        Next j

        ' Calculate new column value
        Dim Value As Double
        Value = Val(C_Arr(i, colIndexes(LBound(colIndexes))))

        For j = LBound(colIndexes) + 1 To UBound(colIndexes)
            Dim nextVal As Double: nextVal = Val(C_Arr(i, colIndexes(j)))
            Select Case operatorType
                Case vbAdd:       Value = Value + nextVal
                Case vbSubtract:  Value = Value - nextVal
                Case vbMultiply:  Value = Value * nextVal
                Case vbDivide:    Value = IIf(nextVal <> 0, Value / nextVal, CVErr(xlErrDiv0))
                Case vbPower:     Value = Value ^ nextVal
                Case vbMod:       Value = IIf(nextVal <> 0, Value Mod nextVal, CVErr(xlErrDiv0))
                Case Else:        Err.Raise vbObjectError + 2082, , "Invalid operator type."
            End Select
        Next j

        resultArr(i, InsertColAt) = Value

        ' Copy columns after the insert Fieldition
        For j = InsertColAt To c2
            resultArr(i, j + 1) = C_Arr(i, j)
        Next j
    Next i

    ' Update internal array and dictionary
    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Private Function ValToArray(ArrValue As Variant, LengthOfArray As Long, Optional ZeroBased As Boolean = True) As Variant
    Dim startIndex As Long, endIndex As Long
    Dim returnArr() As Variant

    If ZeroBased Then
        startIndex = 0
        endIndex = LengthOfArray - 1
    Else
        startIndex = 1
        endIndex = LengthOfArray
    End If

    ReDim returnArr(startIndex To endIndex)

    Dim i As Long
    For i = startIndex To endIndex
        returnArr(i) = ArrValue
    Next i

    ValToArray = returnArr
End Function

Public Sub AddConditionalColumn(ByVal Field As Long, _
                                ByVal ColIndex As Variant, _
                                ByVal criteria As Variant, _
                                ByVal TrueValue As Variant, _
                                Optional ByVal FalseValue As Variant, _
                                Optional ByVal matchmode As Variant = vbTextEquals, _
                                Optional CompareMethod As VbCompareMethod = vbBinaryCompare, _
                                Optional InvertMatch As Boolean = False, _
                                Optional UserLogic As vbFilterLogic = vbLogicAnd)

    Dim i As Long, j As Long, n As Long
    Dim totalRows As Long, totalCols As Long
    Dim resultArr() As Variant, colIndexArr As Variant, criteriaArr As Variant, MatchModeArr As Variant
    Dim TrueValueArr As Variant, FalseValueArr As Variant
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long, InsertColAt As Long
    Dim Bool As Boolean

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2065, , "Array not initialized."

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)

    totalRows = r2 - r1 + 1
    totalCols = c2 - c1 + 1
    If Field < 1 Or Field > totalCols + 1 Then Err.Raise vbObjectError + 2062, , "Field is out of Column Bound: " & Field

    colIndexArr = NormalizeToArray(ColIndex)
    criteriaArr = IIf(IsArray(criteria), criteria, Array(criteria))
    MatchModeArr = NormalizeMatchMode(colIndexArr, matchmode)
    
    If DimensionOfArray(colIndexArr) > 1 Or DimensionOfArray(criteriaArr) > 1 Then
        Err.Raise vbObjectError + 2050, , "Both colIndex and criteria must be 1D arrays or a pair of Single Variables."
    End If
    
    If UBound(colIndexArr) - LBound(colIndexArr) <> UBound(criteriaArr) - LBound(criteriaArr) Then
        Err.Raise vbObjectError + 2051, , "colIndex and criteria must be of equal length."
    End If
    
    ' Prepare TrueValueArr
    If IsArray(TrueValue) Then
        If DimensionOfArray(TrueValue) = 2 Then
            TrueValueArr = TrueValue
        Else
            TrueValueArr = Convert2D(ValToArray(TrueValue, totalRows, False), vbToCol)
        End If
    Else
        TrueValueArr = Convert2D(ValToArray(TrueValue, totalRows, False), vbToCol)
    End If

    ' Prepare FalseValueArr
    If Not IsMissing(FalseValue) Then
        If IsArray(FalseValue) Then
            If DimensionOfArray(FalseValue) = 2 Then
                FalseValueArr = FalseValue
            Else
                FalseValueArr = Convert2D(ValToArray(FalseValue, totalRows, False), vbToCol)
            End If
        Else
            FalseValueArr = Convert2D(ValToArray(FalseValue, totalRows, False), vbToCol)
        End If
    Else
        FalseValueArr = Convert2D(ValToArray(Empty, totalRows, False), vbToCol)
    End If

    InsertColAt = c1 + Field - 1
    ReDim resultArr(r1 To r2, c1 To c2 + 1)

    For i = r1 To r2
        ' Copy columns before the insert Fieldition
        For j = c1 To InsertColAt - 1
            resultArr(i, j) = C_Arr(i, j)
        Next j
    
        ' Evaluate condition with AND or OR logic
        Select Case UserLogic
            Case vbLogicAnd
                Bool = True
                For n = LBound(colIndexArr) To UBound(colIndexArr)
                    If Not IsMatch(C_Arr(i, colIndexArr(n)), criteriaArr(n), MatchModeArr(n), CompareMethod) Then
                        Bool = False
                        Exit For
                    End If
                Next n
            Case vbLogicOr
                Bool = False
                For n = LBound(colIndexArr) To UBound(colIndexArr)
                    If IsMatch(C_Arr(i, colIndexArr(n)), criteriaArr(n), MatchModeArr(n), CompareMethod) Then
                        Bool = True
                        Exit For
                    End If
                Next n
            Case Else
                Err.Raise vbObjectError + 2082, , "Invalid logic type."
        End Select
        
        If InvertMatch Then Bool = Not Bool

        ' Insert conditional value
        resultArr(i, InsertColAt) = IIf(Bool, TrueValueArr(i, 1), FalseValueArr(i, 1))

        ' Copy columns after the insert Fieldition
        For j = InsertColAt To c2
            resultArr(i, j + 1) = C_Arr(i, j)
        Next j
    Next i

    ' Update internal array and dictionary
    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub
    
Public Sub AddComparitiveColumn(ByVal Field As Long, _
                                ByVal ColIndex As Variant, _
                                ByVal CompareCol As Variant, _
                                ByVal TrueValue As Variant, _
                                Optional ByVal FalseValue As Variant, _
                                Optional ByVal matchmode As Variant = vbTextEquals, _
                                Optional CompareMethod As VbCompareMethod = vbBinaryCompare, _
                                Optional InvertMatch As Boolean = False, _
                                Optional UserLogic As vbFilterLogic = vbLogicAnd)

    Dim i As Long, j As Long, n As Long
    Dim totalRows As Long, totalCols As Long
    Dim resultArr() As Variant, colIndexArr As Variant, CompareColArr As Variant, MatchModeArr As Variant
    Dim TrueValueArr As Variant, FalseValueArr As Variant
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long, InsertColAt As Long
    Dim Bool As Boolean

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2065, , "Array not initialized."

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)

    totalRows = r2 - r1 + 1
    totalCols = c2 - c1 + 1
    If Field < 1 Or Field > totalCols + 1 Then Err.Raise vbObjectError + 2062, , "Field is out of Column Bound: " & Field

    colIndexArr = NormalizeToArray(ColIndex)
    CompareColArr = NormalizeToArray(CompareCol)
    MatchModeArr = NormalizeMatchMode(colIndexArr, matchmode)

    If DimensionOfArray(colIndexArr) > 1 Or DimensionOfArray(CompareColArr) > 1 Then
        Err.Raise vbObjectError + 2050, , "Both colIndex and CompareCol must be 1D arrays or single values."
    End If

    If UBound(colIndexArr) - LBound(colIndexArr) <> UBound(CompareColArr) - LBound(CompareColArr) Then
        Err.Raise vbObjectError + 2051, , "colIndex and CompareCol must be of equal length."
    End If

    ' Prepare TrueValueArr
    If IsArray(TrueValue) Then
        If DimensionOfArray(TrueValue) = 2 Then
            TrueValueArr = TrueValue
        Else
            TrueValueArr = Convert2D(ValToArray(TrueValue, totalRows, False), vbToCol)
        End If
    Else
        TrueValueArr = Convert2D(ValToArray(TrueValue, totalRows, False), vbToCol)
    End If

    ' Prepare FalseValueArr
    If Not IsMissing(FalseValue) Then
        If IsArray(FalseValue) Then
            If DimensionOfArray(FalseValue) = 2 Then
                FalseValueArr = FalseValue
            Else
                FalseValueArr = Convert2D(ValToArray(FalseValue, totalRows, False), vbToCol)
            End If
        Else
            FalseValueArr = Convert2D(ValToArray(FalseValue, totalRows, False), vbToCol)
        End If
    Else
        FalseValueArr = Convert2D(ValToArray(Empty, totalRows, False), vbToCol)
    End If

    InsertColAt = c1 + Field - 1
    ReDim resultArr(r1 To r2, c1 To c2 + 1)

    For i = r1 To r2
        ' Copy columns before the insert Fieldition
        For j = c1 To InsertColAt - 1
            resultArr(i, j) = C_Arr(i, j)
        Next j

        ' Evaluate condition with AND or OR logic
        Select Case UserLogic
            Case vbLogicAnd
                Bool = True
                For n = LBound(colIndexArr) To UBound(colIndexArr)
                    If Not IsMatch(C_Arr(i, colIndexArr(n)), CompareColArr(n), MatchModeArr(n), CompareMethod) Then
                        Bool = False
                        Exit For
                    End If
                Next n
            Case vbLogicOr
                Bool = False
                For n = LBound(colIndexArr) To UBound(colIndexArr)
                    If IsMatch(C_Arr(i, colIndexArr(n)), CompareColArr(n), MatchModeArr(n), CompareMethod) Then
                        Bool = True
                        Exit For
                    End If
                Next n
            Case Else
                Err.Raise vbObjectError + 2082, , "Invalid logic type."
        End Select

        If InvertMatch Then Bool = Not Bool

        ' Insert conditional value
        resultArr(i, InsertColAt) = IIf(Bool, TrueValueArr(i, 1), FalseValueArr(i, 1))

        ' Copy columns after the insert Fieldition
        For j = InsertColAt To c2
            resultArr(i, j + 1) = C_Arr(i, j)
        Next j
    Next i

    ' Update internal array and dictionary
    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Sub AddConditionalComparitiveColumn(ByVal Field As Long, _
                                ByVal ColIndexCriteria As Variant, _
                                ByVal criteria As Variant, _
                                ByVal ColIndexCompare As Variant, _
                                ByVal CompareCol As Variant, _
                                ByVal TrueValue As Variant, _
                                Optional ByVal FalseValue As Variant, _
                                Optional ByVal matchmode As Variant = vbTextEquals, _
                                Optional CompareMethod As VbCompareMethod = vbBinaryCompare, _
                                Optional InvertMatch As Boolean = False, _
                                Optional UserLogic As vbFilterLogic = vbLogicAnd)

    Dim i As Long, j As Long, n As Long
    Dim totalRows As Long, totalCols As Long
    Dim resultArr() As Variant, ColIndexCriteriaArr As Variant, ColIndexCompareArr As Variant
    Dim criteriaArr As Variant, CompareColArr As Variant, MatchModeArr As Variant
    Dim TrueValueArr As Variant, FalseValueArr As Variant
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long, InsertColAt As Long
    Dim Bool As Boolean

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2065, , "Array not initialized."

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)

    totalRows = r2 - r1 + 1
    totalCols = c2 - c1 + 1
    If Field < 1 Or Field > totalCols + 1 Then Err.Raise vbObjectError + 2062, , "Field is out of Column Bound: " & Field

    ColIndexCriteriaArr = NormalizeToArray(ColIndexCriteria)
    criteriaArr = IIf(IsArray(criteria), criteria, Array(criteria))
    ColIndexCompareArr = NormalizeToArray(ColIndexCompare)
    CompareColArr = NormalizeToArray(CompareCol)
    MatchModeArr = NormalizeMatchMode(ColIndexCriteriaArr, matchmode)

    If DimensionOfArray(ColIndexCriteriaArr) > 1 Or DimensionOfArray(criteriaArr) > 1 Then
        Err.Raise vbObjectError + 2050, , "Both colIndex and criteria must be 1D arrays or a pair of Single Variables."
    End If
    
    If UBound(ColIndexCriteriaArr) - LBound(ColIndexCriteriaArr) <> UBound(criteriaArr) - LBound(criteriaArr) Then
        Err.Raise vbObjectError + 2051, , "colIndex and criteria must be of equal length."
    End If
    
    If DimensionOfArray(ColIndexCompareArr) > 1 Or DimensionOfArray(CompareColArr) > 1 Then
        Err.Raise vbObjectError + 2050, , "Both colIndex and CompareCol must be 1D arrays or single values."
    End If

    If UBound(ColIndexCompareArr) - LBound(ColIndexCompareArr) <> UBound(CompareColArr) - LBound(CompareColArr) Then
        Err.Raise vbObjectError + 2051, , "colIndex and CompareCol must be of equal length."
    End If

    ' Prepare TrueValueArr
    TrueValueArr = Convert2D(ValToArray(TrueValue, totalRows, False), vbToCol)

    ' Prepare FalseValueArr
    If Not IsMissing(FalseValue) Then
        FalseValueArr = Convert2D(ValToArray(FalseValue, totalRows, False), vbToCol)
    Else
        FalseValueArr = Convert2D(ValToArray(Empty, totalRows, False), vbToCol)
    End If

    InsertColAt = c1 + Field - 1
    ReDim resultArr(r1 To r2, c1 To c2 + 1)

    For i = r1 To r2
        ' Copy columns before the insert position
        For j = c1 To InsertColAt - 1
            resultArr(i, j) = C_Arr(i, j)
        Next j

        ' Evaluate condition with AND or OR logic
        Select Case UserLogic
            Case vbLogicAnd
                Bool = True
                ' Criteria checks
                For n = LBound(ColIndexCriteriaArr) To UBound(ColIndexCriteriaArr)
                    If Not IsMatch(C_Arr(i, ColIndexCriteriaArr(n)), criteriaArr(n), MatchModeArr(n), CompareMethod) Then
                        Bool = False: Exit For
                    End If
                Next n
                ' Comparative checks
                If Bool Then
                    For n = LBound(ColIndexCompareArr) To UBound(ColIndexCompareArr)
                        If Not IsMatch(C_Arr(i, ColIndexCompareArr(n)), C_Arr(i, CompareColArr(n)), MatchModeArr(n), CompareMethod) Then
                            Bool = False: Exit For
                        End If
                    Next n
                End If
            Case vbLogicOr
                Bool = False
                ' Criteria checks
                For n = LBound(ColIndexCriteriaArr) To UBound(ColIndexCriteriaArr)
                    If IsMatch(C_Arr(i, ColIndexCriteriaArr(n)), criteriaArr(n), MatchModeArr(n), CompareMethod) Then
                        Bool = True: Exit For
                    End If
                Next n
                ' Comparative checks
                If Not Bool Then
                    For n = LBound(ColIndexCompareArr) To UBound(ColIndexCompareArr)
                        If IsMatch(C_Arr(i, ColIndexCompareArr(n)), C_Arr(i, CompareColArr(n)), MatchModeArr(n), CompareMethod) Then
                            Bool = True: Exit For
                        End If
                    Next n
                End If
            Case Else
                Err.Raise vbObjectError + 2082, , "Invalid logic type."
        End Select

        If InvertMatch Then Bool = Not Bool

        ' Insert conditional value
        resultArr(i, InsertColAt) = IIf(Bool, TrueValueArr(i, 1), FalseValueArr(i, 1))

        ' Copy columns after the insert position
        For j = InsertColAt To c2
            resultArr(i, j + 1) = C_Arr(i, j)
        Next j
    Next i

    ' Update internal array and dictionary
    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Sub CopyCols(ByVal Field As Long, ByVal ColumnsToCopy As Variant)
    Dim i As Long, j As Long, k As Long
    Dim totalRows As Long, totalCols As Long
    Dim resultArr() As Variant
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    Dim InsertColAt As Long
    Dim CopyCols() As Variant

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2065, , "Array not initialized."

    CopyCols = NormalizeToArray(ColumnsToCopy)

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)

    totalRows = r2 - r1 + 1
    totalCols = c2 - c1 + 1

    If Field < 1 Or Field > totalCols + 1 Then
        Err.Raise vbObjectError + 2062, , "Insert Fieldition out of bounds: " & Field
    End If

    ' Validate all columns to copy
    For k = LBound(CopyCols) To UBound(CopyCols)
        If CopyCols(k) < 1 Or CopyCols(k) > totalCols Then
            Err.Raise vbObjectError + 2063, , "ColumnToCopy out of bounds: " & CopyCols(k)
        End If
    Next k

    InsertColAt = c1 + Field - 1
    ReDim resultArr(r1 To r2, c1 To c2 + UBound(CopyCols) - LBound(CopyCols) + 1)

    For i = r1 To r2
        ' Copy columns before the insert Fieldition
        For j = c1 To InsertColAt - 1
            resultArr(i, j) = C_Arr(i, j)
        Next j

        ' Insert copied columns
        For k = LBound(CopyCols) To UBound(CopyCols)
            resultArr(i, InsertColAt + k - LBound(CopyCols)) = C_Arr(i, c1 + CopyCols(k) - 1)
        Next k

        ' Copy columns after the insert Fieldition
        For j = InsertColAt To c2
            resultArr(i, j + UBound(CopyCols) - LBound(CopyCols) + 1) = C_Arr(i, j)
        Next j
    Next i

    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Sub CopyRows(ByVal Field As Long, ByVal RowsToCopy As Variant)
    Dim i As Long, j As Long, k As Long
    Dim totalRows As Long, totalCols As Long
    Dim resultArr() As Variant
    Dim r1 As Long, r2 As Long, c1 As Long, c2 As Long
    Dim InsertRowAt As Long
    Dim CopyRows() As Variant

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2070, , "Array not initialized."

    CopyRows = NormalizeToArray(RowsToCopy)

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)

    totalRows = r2 - r1 + 1
    totalCols = c2 - c1 + 1

    If Field < 1 Or Field > totalRows + 1 Then
        Err.Raise vbObjectError + 2071, , "Insert Fieldition out of bounds: " & Field
    End If

    ' Validate all rows to copy
    For k = LBound(CopyRows) To UBound(CopyRows)
        If CopyRows(k) < 1 Or CopyRows(k) > totalRows Then
            Err.Raise vbObjectError + 2072, , "RowToCopy out of bounds: " & CopyRows(k)
        End If
    Next k

    InsertRowAt = r1 + Field - 1
    ReDim resultArr(r1 To r2 + UBound(CopyRows) - LBound(CopyRows) + 1, c1 To c2)

    For j = c1 To c2
        ' Copy rows before the insert Fieldition
        For i = r1 To InsertRowAt - 1
            resultArr(i, j) = C_Arr(i, j)
        Next i

        ' Insert copied rows
        For k = LBound(CopyRows) To UBound(CopyRows)
            resultArr(InsertRowAt + k - LBound(CopyRows), j) = C_Arr(r1 + CopyRows(k) - 1, j)
        Next k

        ' Copy rows after the insert Fieldition
        For i = InsertRowAt To r2
            resultArr(i + UBound(CopyRows) - LBound(CopyRows) + 1, j) = C_Arr(i, j)
        Next i
    Next j

    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Function ConcatenateRows(ByRef RowIndexes() As Long, ByVal delimiter As String, Optional ByVal SkipEmpty As Boolean = False) As Arrays
    Dim i As Long, j As Long
    Dim totalRows As Long, totalCols As Long
    Dim resultArr() As Variant
    Dim arrObj As New Arrays
    Dim dict As Object

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2086, , "Array not initialized."
    End If

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)
    ReDim resultArr(1 To totalCols)

    For j = 1 To totalCols
        Dim concatStr As String: concatStr = ""
        For i = LBound(RowIndexes) To UBound(RowIndexes)
            Dim row As Long: row = RowIndexes(i)
            If row < 1 Or row > totalRows Then
                Err.Raise vbObjectError + 2088, , "Row index out of bounds: " & row
            End If

            Dim cellVal As String: cellVal = Trim(CStr(C_Arr(row, j)))
            If SkipEmpty And Len(cellVal) = 0 Then
                ' Skip empty cells
            Else
                concatStr = concatStr & cellVal & delimiter
            End If
        Next i

        If Len(concatStr) > 0 Then
            resultArr(j) = left(concatStr, Len(concatStr) - Len(delimiter)) ' remove trailing delimiter
        Else
            resultArr(j) = ""
        End If
    Next j

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add C_Ref, resultArr

    Set arrObj.Arrays = dict
    arrObj.UseRef C_Ref
    Set ConcatenateRows = arrObj
End Function

Public Function ConcatenateColumns(colIndexes As Variant, ByVal delimiter As String, Optional ByVal SkipEmpty As Boolean = False) As Arrays
    Dim i As Long, j As Long
    Dim totalRows As Long, totalCols As Long
    Dim resultArr() As Variant
    Dim arrObj As New Arrays
    Dim dict As Object

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2083, , "Array not initialized."
    End If

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)
    ReDim resultArr(1 To totalRows)

    For i = 1 To totalRows
        Dim concatStr As String: concatStr = ""
        For j = LBound(colIndexes) To UBound(colIndexes)
            Dim Col As Long: Col = colIndexes(j)
            If Col < 1 Or Col > totalCols Then
                Err.Raise vbObjectError + 2085, , "Column index out of bounds: " & Col
            End If

            Dim cellVal As String: cellVal = Trim(CStr(C_Arr(i, Col)))
            If SkipEmpty And Len(cellVal) = 0 Then
                ' Skip empty cells
            Else
                concatStr = concatStr & cellVal & delimiter
            End If
        Next j

        If Len(concatStr) > 0 Then
            resultArr(i) = left(concatStr, Len(concatStr) - Len(delimiter)) ' remove trailing delimiter
        Else
            resultArr(i) = ""
        End If
    Next i

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add C_Ref, resultArr

    Set arrObj.Arrays = dict
    arrObj.UseRef C_Ref
    Set ConcatenateColumns = arrObj
    
End Function

Public Sub RearrangeRows(ByVal originalRows As Variant, ByVal targetRows As Variant)
    Dim i As Long, j As Long
    Dim totalRows As Long, totalCols As Long
    Dim resultArr() As Variant
    Dim rowMap As Object
    Dim originalRowsArr() As Variant, targetRowsArr() As Variant
    Dim r1 As Long, r2 As Long

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2092, , "Array not initialized."

    originalRowsArr = NormalizeToArray(originalRows)
    targetRowsArr = NormalizeToArray(targetRows)

    If UBound(originalRowsArr) - LBound(originalRowsArr) <> UBound(targetRowsArr) - LBound(targetRowsArr) Then
        Err.Raise vbObjectError + 2094, , "originalRows and targetRows must be of equal length."
    End If

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    totalRows = r2 - r1 + 1
    totalCols = UBound(C_Arr, 2)

    Set rowMap = CreateObject("Scripting.Dictionary")

    ' Build mapping and validate
    For i = LBound(originalRowsArr) To UBound(originalRowsArr)
        Dim origRow As Long: origRow = originalRowsArr(i)
        Dim targetRow As Long: targetRow = targetRowsArr(i)

        If rowMap.Exists(targetRow) Then
            Err.Raise vbObjectError + 2096, , "Duplicate target row: " & targetRow
        End If
        If origRow < 1 Or origRow > totalRows Or targetRow < 1 Or targetRow > totalRows Then
            Err.Raise vbObjectError + 2097, , "Row index out of bounds at Fieldition " & i
        End If

        rowMap(targetRow) = origRow
    Next i

    ' Rearrange rows
    ReDim resultArr(r1 To r2, 1 To totalCols)

    For i = r1 To r2
        Dim sourceRow As Long
        If rowMap.Exists(i) Then
            sourceRow = rowMap(i)
        Else
            sourceRow = i
        End If

        For j = 1 To totalCols
            resultArr(i, j) = C_Arr(sourceRow, j)
        Next j
    Next i

    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Sub RearrangeCols(ByVal originalCols As Variant, ByVal targetCols As Variant)
    Dim i As Long, j As Long
    Dim totalRows As Long, totalCols As Long
    Dim resultArr() As Variant
    Dim colMap As Object
    Dim originalColsArr() As Variant, targetColsArr() As Variant
    Dim c1 As Long, c2 As Long

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2107, , "Array not initialized."

    originalColsArr = NormalizeToArray(originalCols)
    targetColsArr = NormalizeToArray(targetCols)

    If UBound(originalColsArr) - LBound(originalColsArr) <> UBound(targetColsArr) - LBound(targetColsArr) Then
        Err.Raise vbObjectError + 2108, , "originalCols and targetCols must be of equal length."
    End If

    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)
    totalCols = c2 - c1 + 1
    totalRows = UBound(C_Arr, 1)

    Set colMap = CreateObject("Scripting.Dictionary")

    ' Build mapping and validate
    For i = LBound(originalColsArr) To UBound(originalColsArr)
        Dim origCol As Long: origCol = originalColsArr(i)
        Dim targetCol As Long: targetCol = targetColsArr(i)

        If colMap.Exists(targetCol) Then
            Err.Raise vbObjectError + 2109, , "Duplicate target column: " & targetCol
        End If
        If origCol < 1 Or origCol > totalCols Or targetCol < 1 Or targetCol > totalCols Then
            Err.Raise vbObjectError + 2110, , "Column index out of bounds at Fieldition " & i
        End If

        colMap(targetCol) = origCol
    Next i

    ' Rearrange columns
    ReDim resultArr(1 To totalRows, c1 To c2)

    For j = c1 To c2
        Dim sourceCol As Long
        If colMap.Exists(j) Then
            sourceCol = colMap(j)
        Else
            sourceCol = j
        End If

        For i = 1 To totalRows
            resultArr(i, j) = C_Arr(i, sourceCol)
        Next i
    Next j

    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Sub SwapRows(ByVal row1 As Variant, ByVal row2 As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim row1Arr() As Variant, row2Arr() As Variant
    Dim totalRows As Long, totalCols As Long

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2096, , "Array not initialized."
    End If

    row1Arr = NormalizeToArray(row1)
    row2Arr = NormalizeToArray(row2)

    If UBound(row1Arr) - LBound(row1Arr) <> UBound(row2Arr) - LBound(row2Arr) Then
        Err.Raise vbObjectError + 2097, , "Expected arrays of same length."
    End If

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)

    ' Validate bounds
    For i = LBound(row1Arr) To UBound(row1Arr)
        If row1Arr(i) < 1 Or row1Arr(i) > totalRows Or _
           row2Arr(i) < 1 Or row2Arr(i) > totalRows Then
            Err.Raise vbObjectError + 2098, , "Row index out of bounds at Fieldition " & i & "."
        End If
    Next i

    ' Swap rows
    For i = LBound(row1Arr) To UBound(row1Arr)
        For j = 1 To totalCols
            temp = C_Arr(row1Arr(i), j)
            C_Arr(row1Arr(i), j) = C_Arr(row2Arr(i), j)
            C_Arr(row2Arr(i), j) = temp
        Next j
    Next i

    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Sub SwapCols(ByVal col1 As Variant, ByVal col2 As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim col1Arr() As Variant, col2Arr() As Variant
    Dim totalRows As Long, totalCols As Long

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2100, , "Array not initialized."
    End If

    col1Arr = NormalizeToArray(col1)
    col2Arr = NormalizeToArray(col2)

    If UBound(col1Arr) - LBound(col1Arr) <> UBound(col2Arr) - LBound(col2Arr) Then
        Err.Raise vbObjectError + 2101, , "Expected arrays of same length."
    End If

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)

    ' Validate bounds
    For i = LBound(col1Arr) To UBound(col1Arr)
        If col1Arr(i) < 1 Or col1Arr(i) > totalCols Or _
           col2Arr(i) < 1 Or col2Arr(i) > totalCols Then
            Err.Raise vbObjectError + 2102, , "Column index out of bounds at Fieldition " & i & "."
        End If
    Next i

    ' Swap columns
    For i = LBound(col1Arr) To UBound(col1Arr)
        For j = 1 To totalRows
            temp = C_Arr(j, col1Arr(i))
            C_Arr(j, col1Arr(i)) = C_Arr(j, col2Arr(i))
            C_Arr(j, col2Arr(i)) = temp
        Next j
    Next i

    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Private Function NormalizeDescFlags(descendingFlags As Variant, colArr() As Variant) As Variant
    Dim i As Long
    
    If IsArray(descendingFlags) Then
        Dim flag As Variant
        For Each flag In descendingFlags
            If VarType(flag) <> vbBoolean Then
                Err.Raise vbObjectError + 2102, , "Expected a Boolean: " & CStr(flag)
            End If
        Next flag
        
        If UBound(descendingFlags) - LBound(descendingFlags) <> UBound(colArr) - LBound(colArr) Then
            Err.Raise vbObjectError + 2101, , "Descending flags must match column count."
        End If
    
        NormalizeDescFlags = descendingFlags
    Else
        If VarType(descendingFlags) <> vbBoolean Then
            Err.Raise vbObjectError + 2102, , "Expected a Boolean: " & CStr(descendingFlags)
        End If
        
        Dim descending() As Variant
        ReDim descending(LBound(colArr) To UBound(colArr))
        For i = LBound(descending) To UBound(descending)
            descending(i) = descendingFlags
        Next i
        NormalizeDescFlags = descending
    End If
End Function

Public Sub SortByColumns(ByVal colIndexes As Variant, Optional ByVal descendingFlags As Variant, Optional SortMethod As vbUserSortMethod = vbUserMergeSort)
    Dim totalRows As Long, totalCols As Long
    Dim colArr() As Variant, descending() As Variant
    Dim i As Long

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2100, , "Array not initialized."

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)

    ' Normalize column indexes
    colArr = NormalizeToArray(colIndexes)

    ' Normalize descending flags
    If IsMissing(descendingFlags) Then
        ReDim descending(LBound(colArr) To UBound(colArr))
        For i = LBound(descending) To UBound(descending)
            descending(i) = False
        Next i
    Else
        descending = NormalizeDescFlags(descendingFlags, colArr)
    End If
        
    ' Build index array for sorting
    Dim idx() As Long
    ReDim idx(1 To totalRows)
    For i = 1 To totalRows
        idx(i) = i
    Next i

    Select Case SortMethod
        Case vbUserQuickSort
            ' Perform QuickSort on index array using custom comparison
            QuickSort idx, LBound(idx), UBound(idx), colArr, descending
        Case vbUserMergeSort
            ' Perform MergeSort on index array using custom comparison
             MergeSort idx, LBound(idx), UBound(idx), colArr, descending
        Case Else
            Err.Raise vbObjectError + 2051, "Sorting method: " & SortMethod & " not supported."
    End Select

    ' Rebuild sorted array
    Dim tempArr() As Variant
    ReDim tempArr(1 To totalRows, 1 To totalCols)
    Dim r As Long, c As Long
    For r = 1 To totalRows
        For c = 1 To totalCols
            tempArr(r, c) = C_Arr(idx(r), c)
        Next c
    Next r

    C_Arr = tempArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

'==========================
' QuickSort Implementation
'==========================

Private Sub QuickSort(ByRef arr() As Long, low As Long, high As Long, colArr() As Variant, descending() As Variant)
    Dim i As Long, j As Long, pivot As Long, temp As Long
    Dim mid As Long
    
    Do While low < high
        ' Median-of-three pivot selection
        mid = (low + high) \ 2
        If CompareRows(arr(low), arr(mid), colArr, descending) > 0 Then Swap arr(low), arr(mid)
        If CompareRows(arr(low), arr(high), colArr, descending) > 0 Then Swap arr(low), arr(high)
        If CompareRows(arr(mid), arr(high), colArr, descending) > 0 Then Swap arr(mid), arr(high)
        pivot = arr(mid)
        
        i = low
        j = high
        Do While i <= j
            Do While CompareRows(arr(i), pivot, colArr, descending) < 0
                i = i + 1
            Loop
            Do While CompareRows(arr(j), pivot, colArr, descending) > 0
                j = j - 1
            Loop
            If i <= j Then
                Swap arr(i), arr(j)
                i = i + 1
                j = j - 1
            End If
        Loop
        
        ' Tail recursion elimination: sort smaller partition first
        If (j - low) < (high - i) Then
            If low < j Then QuickSort arr, low, j, colArr, descending
            low = i
        Else
            If i < high Then QuickSort arr, i, high, colArr, descending
            high = j
        End If
    Loop
End Sub

'==========================
' MergeSort Implementation
'==========================
Private Sub MergeSort(ByRef arr() As Long, ByVal left As Long, ByVal right As Long, _
                      ByRef colArr() As Variant, ByRef descending() As Variant)
    If left >= right Then Exit Sub

    Dim mid As Long
    mid = (left + right) \ 2

    MergeSort arr, left, mid, colArr, descending
    MergeSort arr, mid + 1, right, colArr, descending

    Merge arr, left, mid, right, colArr, descending
End Sub

Private Sub Merge(ByRef arr() As Long, ByVal left As Long, ByVal mid As Long, ByVal right As Long, _
                  ByRef colArr() As Variant, ByRef descending() As Variant)
    Dim temp() As Long
    Dim i As Long, j As Long, k As Long

    ReDim temp(left To right)

    i = left
    j = mid + 1
    k = left

    Do While i <= mid And j <= right
        If CompareRows(arr(i), arr(j), colArr, descending) <= 0 Then
            temp(k) = arr(i)
            i = i + 1
        Else
            temp(k) = arr(j)
            j = j + 1
        End If
        k = k + 1
    Loop

    Do While i <= mid
        temp(k) = arr(i)
        i = i + 1
        k = k + 1
    Loop

    Do While j <= right
        temp(k) = arr(j)
        j = j + 1
        k = k + 1
    Loop

    For k = left To right
        arr(k) = temp(k)
    Next k
End Sub

'==========================
' Row Comparison Logic
'==========================

Private Function CompareRows(row1 As Long, row2 As Long, colArr() As Variant, descending() As Variant) As Long
    Dim k As Long, val1 As Variant, val2 As Variant
    For k = LBound(colArr) To UBound(colArr)
        val1 = C_Arr(row1, colArr(k))
        val2 = C_Arr(row2, colArr(k))
        If val1 <> val2 Then
            If descending(k) Then
                CompareRows = IIf(val1 < val2, 1, -1)
            Else
                CompareRows = IIf(val1 < val2, -1, 1)
            End If
            Exit Function
        End If
    Next k
    CompareRows = 0
End Function

Public Sub MoveRow(ByVal fromIndex As Long, ByVal toIndex As Long)
    Dim j As Long, temp As Variant
    Dim totalCols As Long, totalRows As Long
    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2105, , "Array not initialized."

    totalRows = UBound(C_Arr, 1)
    totalCols = UBound(C_Arr, 2)

    If fromIndex < 1 Or fromIndex > totalRows Or toIndex < 1 Or toIndex > totalRows Then
        Err.Raise vbObjectError + 2106, , "Row index out of bounds."
    End If

    If fromIndex = toIndex Then Exit Sub

    Dim resultArr() As Variant
    ReDim resultArr(1 To totalRows, 1 To totalCols)

    Dim rowMap() As Long
    ReDim rowMap(1 To totalRows)

    ' Build new row order
    Dim r As Long, Field As Long: Field = 1
    For r = 1 To totalRows
        If r <> fromIndex Then
            If Field = toIndex Then Field = Field + 1
            rowMap(Field) = r
            Field = Field + 1
        End If
    Next r
    rowMap(toIndex) = fromIndex

    ' Rearrange rows
    For r = 1 To totalRows
        For j = 1 To totalCols
            resultArr(r, j) = C_Arr(rowMap(r), j)
        Next j
    Next r

    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Sub MoveCol(ByVal fromIndex As Long, ByVal toIndex As Long)
    Dim i As Long, temp As Variant
    Dim totalCols As Long, totalRows As Long

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2103, , "Array not initialized."

    totalCols = UBound(C_Arr, 2)
    totalRows = UBound(C_Arr, 1)

    If fromIndex < 1 Or fromIndex > totalCols Or toIndex < 1 Or toIndex > totalCols Then
        Err.Raise vbObjectError + 2104, , "Column index out of bounds."
    End If

    If fromIndex = toIndex Then Exit Sub

    Dim resultArr() As Variant
    ReDim resultArr(1 To totalRows, 1 To totalCols)

    Dim colMap() As Long
    ReDim colMap(1 To totalCols)

    ' Build new column order
    Dim c As Long, Field As Long: Field = 1
    For c = 1 To totalCols
        If c <> fromIndex Then
            If Field = toIndex Then Field = Field + 1
            colMap(Field) = c
            Field = Field + 1
        End If
    Next c
    colMap(toIndex) = fromIndex

    ' Rearrange columns
    For i = 1 To totalRows
        For c = 1 To totalCols
            resultArr(i, c) = C_Arr(i, colMap(c))
        Next c
    Next i

    C_Arr = resultArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Function CreateMap(ByVal keyCol As Long, ItemCol As Variant) As Object
    Dim retDict As Object
    Dim itemArr() As Variant
    Dim i As Long, j As Long
    Dim r1 As Long, r2 As Long
    Dim c1 As Long, c2 As Long

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2080, , "Array not initialized."
    End If

    r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
    c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)

    If keyCol > c2 Or keyCol < c1 Then
        Err.Raise vbObjectError + 2081, , "Key column index out of bounds: " & keyCol
    End If

    If IsArray(ItemCol) Then
        For j = LBound(ItemCol) To UBound(ItemCol)
            If ItemCol(j) > c2 Or ItemCol(j) < c1 Then
                Err.Raise vbObjectError + 2082, , "Item column index out of bounds: " & ItemCol(j)
            End If
        Next j
    Else
        If ItemCol > c2 Or ItemCol < c1 Then
            Err.Raise vbObjectError + 2082, , "Item column index out of bounds: " & ItemCol
        End If
    End If

    Set retDict = VBA.CreateObject("Scripting.Dictionary")

    For i = r1 To r2
        If Not retDict.Exists(C_Arr(i, keyCol)) Then
            If IsArray(ItemCol) Then
                ReDim itemArr(LBound(ItemCol) To UBound(ItemCol))
                For j = LBound(ItemCol) To UBound(ItemCol)
                    itemArr(j) = C_Arr(i, ItemCol(j))
                Next j
                retDict.Add C_Arr(i, keyCol), itemArr
            Else
                retDict.Add C_Arr(i, keyCol), C_Arr(i, ItemCol)
            End If
        End If
    Next i

    Set CreateMap = retDict
    
End Function

Public Sub Map(ByVal ColIndex As Variant, ByVal valueMap As Object, Optional ByVal DefaultValue As Variant = Empty)
    Dim i As Long, j As Long
    Dim colIndexArr As Variant, MapArr() As Variant
    Dim key As Variant
    Dim Col As Long
    Dim mappedValue As Variant

    ' Validate source array
    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2080, , "Array not initialized."
    End If

    ' Validate valueMap
    If TypeName(valueMap) <> "Dictionary" Then
        Err.Raise vbObjectError + 2083, , "valueMap must be a Scripting.Dictionary object."
    End If

    ' Normalize column index to array
    colIndexArr = NormalizeToArray(ColIndex)

    ' Ensure colIndexArr is 1D
    If DimensionOfArray(colIndexArr) > 1 Then
        Err.Raise vbObjectError + 2081, , "colIndex must be a 1D array or a Long/Integer variable."
    End If

    ' Validate column indices
    For j = LBound(colIndexArr) To UBound(colIndexArr)
        Col = colIndexArr(j)
        If Col < LBound(C_Arr, 2) Or Col > UBound(C_Arr, 2) Then
            Err.Raise vbObjectError + 2082, , "Column index out of bounds: " & Col
        End If
    Next j
    
    MapArr = C_Arr
    ' Apply mapping in-place
    For i = LBound(MapArr, 1) To UBound(MapArr, 1)
        For j = LBound(colIndexArr) To UBound(colIndexArr)
            Col = colIndexArr(j)
            key = MapArr(i, Col)
            If valueMap.Exists(key) Then
                mappedValue = valueMap(key)
                If IsArray(mappedValue) Then
                    Err.Raise vbObjectError + 2084, , "Mapped value for key '" & key & "' is an array. Arrays are not supported."
                End If
                MapArr(i, Col) = mappedValue
            Else
                MapArr(i, Col) = DefaultValue
            End If
        Next j
    Next i

    ' Update reference dictionary if applicable
    C_Arr = MapArr
    If C_Ref <> "" Then C_Dict(C_Ref) = C_Arr
End Sub

Public Function LookUpRows(ByVal RowIndex As Variant, ByVal valueMap As Object, Optional ByVal DefaultValue As Variant = Empty) As Arrays
    Dim i As Long, j As Long, n As Long
    Dim totalCols As Long, retRows As Long
    Dim rowIndexArr As Variant, retArr As Variant
    Dim key As Variant
    Dim row As Long
    Dim arrObj As New Arrays
    Dim dict As Object
    
    ' Validate source array
    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2080, , "Array not initialized."
    End If

    ' Validate valueMap
    If TypeName(valueMap) <> "Dictionary" Then
        Err.Raise vbObjectError + 2083, , "valueMap must be a Scripting.Dictionary object."
    End If

    ' Normalize row index to array
    rowIndexArr = NormalizeToArray(RowIndex)

    ' Ensure rowIndexArr is 1D
    If DimensionOfArray(rowIndexArr) > 1 Then
        Err.Raise vbObjectError + 2081, , "RowIndex must be a 1D array or a Long/Integer variable."
    End If

    ' Get array dimensions
    totalCols = UBound(C_Arr, 2) - LBound(C_Arr, 2) + 1
    retRows = UBound(rowIndexArr) - LBound(rowIndexArr) + 1

    ' Validate row indices
    For i = LBound(rowIndexArr) To UBound(rowIndexArr)
        row = rowIndexArr(i)
        If row < LBound(C_Arr, 1) Or row > UBound(C_Arr, 1) Then
            Err.Raise vbObjectError + 2082, , "Row index out of bounds: " & row
        End If
    Next i

    ' Initialize return array with only mapped rows
    ReDim retArr(1 To retRows, 1 To totalCols)

    ' Apply mapping
    n = 1
    For i = LBound(rowIndexArr) To UBound(rowIndexArr)
        row = rowIndexArr(i)
        For j = 1 To totalCols
            key = C_Arr(row, j)
            If valueMap.Exists(key) Then
                If IsArray(valueMap(key)) Then
                    retArr(n, j) = "|" & Join(valueMap(key), "|, ") & "|"
                Else
                    retArr(n, j) = valueMap(key)
                End If
            Else
                retArr(n, j) = DefaultValue
            End If
        Next j
        n = n + 1
    Next i
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add C_Ref, retArr
    
    Set arrObj.Arrays = dict
    arrObj.UseRef C_Ref
    Set LookUpRows = arrObj
    
End Function

Public Function LookUpCols(ByVal ColIndex As Variant, ByVal valueMap As Object, Optional ByVal DefaultValue As Variant = Empty) As Arrays
    
    Dim i As Long, j As Long, n As Long
    Dim totalRows As Long, retCols As Long
    Dim colIndexArr As Variant, retArr As Variant
    Dim key As Variant
    Dim Col As Long
    Dim arrObj As New Arrays
    Dim dict As Object
    
    ' Validate source array
    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2080, , "Array not initialized."
    End If

    ' Validate valueMap
    If TypeName(valueMap) <> "Dictionary" Then
        Err.Raise vbObjectError + 2083, , "valueMap must be a Scripting.Dictionary object."
    End If

    ' Normalize column index to array
    colIndexArr = NormalizeToArray(ColIndex)

    ' Ensure colIndexArr is 1D
    If DimensionOfArray(colIndexArr) > 1 Then
        Err.Raise vbObjectError + 2081, , "colIndex must be a 1D array or a Long/Integer variable."
    End If

    ' Get array dimensions
    totalRows = UBound(C_Arr, 1) - LBound(C_Arr, 1) + 1
    retCols = UBound(colIndexArr) - LBound(colIndexArr) + 1

    ' Validate column indices
    For j = LBound(colIndexArr) To UBound(colIndexArr)
        Col = colIndexArr(j)
        If Col < LBound(C_Arr, 2) Or Col > UBound(C_Arr, 2) Then
            Err.Raise vbObjectError + 2082, , "Column index out of bounds: " & Col
        End If
    Next j

    ' Initialize return array with only mapped columns
    ReDim retArr(1 To totalRows, 1 To retCols)

    ' Apply mapping
    For i = 1 To totalRows
        n = 1
        For j = LBound(colIndexArr) To UBound(colIndexArr)
            Col = colIndexArr(j)
            key = C_Arr(i, Col)
            If valueMap.Exists(key) Then
                If IsArray(valueMap(key)) Then
                    retArr(i, n) = "|" & Join(valueMap(key), "|, ") & "|"
                Else
                    retArr(i, n) = valueMap(key)
                End If
            Else
                retArr(i, n) = DefaultValue
            End If
            n = n + 1
        Next j
    Next i
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add C_Ref, retArr
    
    Set arrObj.Arrays = dict
    arrObj.UseRef C_Ref
    Set LookUpCols = arrObj
    
End Function

Public Function GroupBy(ByVal colIndexes As Variant, _
                        ByVal aggFunc As String, _
                        ByVal targetCol As Long, _
                        Optional matchmode As Variant = vbTextEquals, _
                        Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Arrays
    Dim i As Long, j As Long
    Dim totalRows As Long
    Dim colIndexArr As Variant, MatchModeArr As Variant
    Dim keyDict As Object, resultArr() As Variant
    Dim key As String, groupKeys As Variant
    Dim groupCount As Long
    Dim dict As Object
    Dim arrObj As Arrays

    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2091, , "Array not initialized."
    If targetCol < LBound(C_Arr, 2) Or targetCol > UBound(C_Arr, 2) Then Err.Raise vbObjectError + 2093, , "Target column out of bounds."

    colIndexArr = NormalizeToArray(colIndexes)
    MatchModeArr = NormalizeMatchMode(colIndexArr, matchmode)

    totalRows = UBound(C_Arr, 1)
    Set keyDict = CreateObject("Scripting.Dictionary")

    ' Step 1: Build unique group keys
    For i = 1 To totalRows
        key = ""
        For j = LBound(colIndexArr) To UBound(colIndexArr)
            key = key & "|" & CStr(C_Arr(i, colIndexArr(j)))
        Next j
        key = mid(key, 2) ' remove leading "|"
        If Not keyDict.Exists(key) Then keyDict.Add key, True
    Next i

    groupKeys = keyDict.keys
    groupCount = keyDict.Count
    ReDim resultArr(1 To groupCount, 1 To UBound(colIndexArr) + 2)

    ' Step 2: Aggregate per group
    For i = 0 To groupCount - 1
        Dim criteria() As Variant
        ReDim criteria(LBound(colIndexArr) To UBound(colIndexArr))
        Dim parts() As String: parts = Split(groupKeys(i), "|")

        For j = LBound(colIndexArr) To UBound(colIndexArr)
            criteria(j) = parts(j)
            resultArr(i + 1, j + 1) = criteria(j)
        Next j

        resultArr(i + 1, UBound(colIndexArr) + 2) = AggregateIfs(aggFunc, colIndexArr, criteria, targetCol, MatchModeArr, CompareMethod)
    Next i

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add C_Ref, resultArr

    Set arrObj = New Arrays
    Set arrObj.Arrays = dict
    arrObj.UseRef C_Ref

    Set GroupBy = arrObj
End Function

Private Function AggregateIfs(aggFunc As String, colIndexes As Variant, criteria As Variant, targetCol As Long, matchModes As Variant, CompareMethod As VbCompareMethod) As Variant
    Select Case LCase(aggFunc)
        Case "sum": AggregateIfs = SumIfs(colIndexes, criteria, targetCol, matchModes, CompareMethod)
        Case "count": AggregateIfs = CountIfs(colIndexes, criteria, matchModes, CompareMethod)
        Case "max": AggregateIfs = MaxIfs(colIndexes, criteria, targetCol, matchModes, CompareMethod)
        Case "min": AggregateIfs = MinIfs(colIndexes, criteria, targetCol, matchModes, CompareMethod)
        Case "average": AggregateIfs = AverageIfs(colIndexes, criteria, targetCol, matchModes, CompareMethod)
        Case Else: Err.Raise vbObjectError + 2092, , "Unsupported aggregation function: " & aggFunc
    End Select
End Function

Public Function pivot(rowFieldIndex As Long, columnFieldIndex As Long, dataFieldIndex As Long, Optional aggFunc As String = "Sum") As Variant
    Dim grouped As Arrays
    Dim rowVals As Collection, colVals As Collection
    Dim result() As Variant
    Dim i As Long, r As Long, c As Long
    Dim tempArr() As Variant
    
    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2091, , "Array not initialized."
    
    Set grouped = Me.GroupBy(Array(rowFieldIndex, columnFieldIndex), aggFunc, dataFieldIndex)
    tempArr = grouped.RefArray

    Set rowVals = Me.DistinctValues(rowFieldIndex)
    Set colVals = Me.DistinctValues(columnFieldIndex)

    ReDim result(0 To rowVals.Count, 0 To colVals.Count)

    ' Fill headers
    For i = 1 To rowVals.Count
        result(i, 0) = rowVals(i)
    Next i
    For i = 1 To colVals.Count
        result(0, i) = colVals(i)
    Next i

    ' Fill data
    For i = 1 To UBound(tempArr, 1)
        r = FindIndex(rowVals, tempArr(i, 1))
        c = FindIndex(colVals, tempArr(i, 2))
        result(r, c) = tempArr(i, 3)
    Next i

    pivot = result
End Function

Public Function DistinctValues(Index As Long, Optional ByCol As Boolean = True) As Variant
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, Val As Variant
    
    If ByCol Then
        ' Distinct values from a column
        For i = LBound(C_Arr, 1) To UBound(C_Arr, 1)
            Val = C_Arr(i, Index)
            If Not dict.Exists(Val) Then dict.Add Val, Nothing
        Next i
    Else
        ' Distinct values from a row
        For i = LBound(C_Arr, 2) To UBound(C_Arr, 2)
            Val = C_Arr(Index, i)
            If Not dict.Exists(Val) Then dict.Add Val, Nothing
        Next i
    End If
    
    DistinctValues = dict.keys()
End Function

Private Function FindIndex(Coll As Collection, Val As Variant) As Long
    Dim i As Long
    For i = 1 To Coll.Count
        If Coll(i) = Val Then
            FindIndex = i
            Exit Function
        End If
    Next i
    FindIndex = 0 ' Not found
End Function

Public Function SplitCells(ByVal delimiter As String, Optional ByVal TransFieldeResult As Boolean = False) As Arrays
    Dim i As Long, j As Long
    Dim parts() As String, maxParts As Long
    Dim tempArr() As Variant, finalArr() As Variant
    Dim isRow As Boolean
    Dim arrObj As New Arrays
    Dim dict As Object
    
    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2091, , "Array not initialized."
    
    ' Determine if it's a single row or single column
    If UBound(C_Arr, 1) = 1 Then
        isRow = True
    ElseIf UBound(C_Arr, 2) = 1 Then
        isRow = False
    Else
        Err.Raise vbObjectError + 2020, , "SplitCells only supports a single row or single column."
    End If

    ' First pass: determine max number of splits
    maxParts = 0
    If isRow Then
        For i = LBound(C_Arr, 2) To UBound(C_Arr, 2)
            parts = Split(C_Arr(1, i), delimiter)
            If UBound(parts) > maxParts Then maxParts = UBound(parts)
        Next i
        ReDim tempArr(1 To 1, 0 To maxParts)
        For i = LBound(C_Arr, 2) To UBound(C_Arr, 2)
            parts = Split(C_Arr(1, i), delimiter)
            For j = 0 To UBound(parts)
                tempArr(1, j + 1) = parts(j)
            Next j
        Next i
    Else
        For i = LBound(C_Arr, 1) To UBound(C_Arr, 1)
            parts = Split(C_Arr(i, 1), delimiter)
            If UBound(parts) > maxParts Then maxParts = UBound(parts)
        Next i
        ReDim tempArr(1 To UBound(C_Arr, 1), 0 To maxParts)
        For i = LBound(C_Arr, 1) To UBound(C_Arr, 1)
            parts = Split(C_Arr(i, 1), delimiter)
            For j = 0 To UBound(parts)
                tempArr(i, j + 1) = parts(j)
            Next j
        Next i
    End If

    ' Optional transFielde
    If TransFieldeResult Then
        ReDim finalArr(LBound(tempArr, 2) To UBound(tempArr, 2), LBound(tempArr, 1) To UBound(tempArr, 1))
        For i = LBound(tempArr, 1) To UBound(tempArr, 1)
            For j = LBound(tempArr, 2) To UBound(tempArr, 2)
                finalArr(j, i) = tempArr(i, j)
            Next j
        Next i
    Else
        finalArr = tempArr
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add C_Ref, finalArr

    Set arrObj.Arrays = dict
    arrObj.UseRef C_Ref
    Set SplitCells = arrObj
End Function

Public Function SplitColumns(ColIndex As Long, delimiter As String) As Variant
    Dim maxParts As Long, i As Long, parts() As String
    Dim tempArray() As Variant
    Dim RowCount As Long, colCount As Long
    
    If Not IsArray(C_Arr) Then Err.Raise vbObjectError + 2091, , "Array not initialized."
    
    RowCount = UBound(C_Arr, 1)
    
    ' Determine max number of splits
    For i = 1 To RowCount
        parts = Split(C_Arr(i, ColIndex), delimiter)
        If UBound(parts) + 1 > maxParts Then maxParts = UBound(parts) + 1
    Next i
    
    colCount = UBound(C_Arr, 2)
    ReDim tempArray(1 To RowCount, 0 To colCount - 1 + maxParts)
    
    ' Copy original data and split column
    For i = 1 To RowCount
        Dim j As Long
        For j = 0 To colCount - 1
            If j <> ColIndex Then
                tempArray(i, j) = C_Arr(i, j)
            End If
        Next j
        
        parts = Split(C_Arr(i, ColIndex), delimiter)
        For j = 0 To UBound(parts)
            tempArray(i, colCount - 1 + j) = parts(j)
        Next j
    Next i
    
    SplitColumns = tempArray
End Function

Public Function ToString() As String
    Dim output As String
    Dim i As Long, j As Long
    
    If Not IsArray(C_Arr) Then
        ToString = "Array not initialized."
        Exit Function
    End If
    
    output = "Reference: " & C_Ref & vbCrLf
    output = output & "Dimensions: " & Dimension & vbCrLf
    output = output & "Bounds: " & BoundsInfo & vbCrLf
    output = output & "Data:" & vbCrLf
    
    For i = LBound(C_Arr, 1) To UBound(C_Arr, 1)
        For j = LBound(C_Arr, 2) To UBound(C_Arr, 2)
            output = output & C_Arr(i, j) & vbTab
        Next j
        output = output & vbCrLf
    Next i
    
    ToString = output
End Function

Public Sub ToRange(ByVal targetRange As Range)
    Dim totalRows As Long, totalCols As Long

    If Not IsArray(C_Arr) Then
        Err.Raise vbObjectError + 2110, , "Array not initialized."
    End If

    totalRows = UBound(C_Arr, 1) - LBound(C_Arr, 1) + 1
    totalCols = UBound(C_Arr, 2) - LBound(C_Arr, 2) + 1

    With targetRange.Resize(totalRows, totalCols)
        .Value = C_Arr
    End With
End Sub

Public Sub ToCSV(ByVal FilePath As String, Optional ByVal delimiter As String = ",")
    Dim fNum As Integer
    Dim i As Long, j As Long
    Dim line As String
    Dim arr As Variant

    If Not C_Dict.Exists(C_Ref) Then
        Err.Raise vbObjectError + 2115, , "Reference key not found in C_Dict."
    End If

    arr = C_Dict(C_Ref)
    fNum = FreeFile

    Open FilePath For Output As #fNum
    For i = LBound(arr, 1) To UBound(arr, 1)
        line = ""
        For j = LBound(arr, 2) To UBound(arr, 2)
            line = line & """" & Replace(CStr(arr(i, j)), """", """""") & """" & delimiter
        Next j
        line = left(line, Len(line) - Len(delimiter)) ' Remove trailing delimiter
        Print #fNum, line
    Next i
    Close #fNum
End Sub

Public Sub ToWorkbook(ByVal FilePath As String, Optional ByVal sheetName As String = "ExportedData")
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim arr As Variant
    Dim fso As Object
    Dim fileExists As Boolean

    If Not C_Dict.Exists(C_Ref) Then
        Err.Raise vbObjectError + 2096, , "Array not found in C_Dict."
    End If

    arr = C_Dict(C_Ref)
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileExists = fso.fileExists(FilePath)

    If fileExists Then
        Set wb = Workbooks.Open(FilePath)
    Else
        Set wb = Workbooks.Add
    End If

    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = sheetName

    With ws.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2))
        .Value = arr
    End With

    If Not fileExists Then
        Application.DisplayAlerts = False
        wb.SaveAs FilePath
        Application.DisplayAlerts = True
    Else
        wb.Save
    End If

    wb.Close SaveChanges:=False
End Sub

