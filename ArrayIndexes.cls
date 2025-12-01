VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayIndexes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private C_Arr() As Variant
Private C_rowIndex() As Variant
Private C_colIndex() As Variant
Public HasRowIndex As Boolean, HasColIndex As Boolean
Private HasRun As Boolean
Private DimArr As Long
Private pParent As Arrays

Private Sub Class_Initialize()
HasRowIndex = False: HasColIndex = False
HasRun = False
End Sub

Public Property Let ParentArrays(ByRef parentObj As Arrays)
    Set pParent = parentObj
End Property

Public Property Get ParentArrays() As Arrays
    Set ParentArrays = pParent
End Property

Public Property Get InitialArray() As Variant
    If IsArray(C_Arr) Then
        InitialArray = C_Arr
        Call Default
    End If
End Property

Public Property Let InitialArray(ByRef ArrayArg As Variant)
    On Error Resume Next
    If IsArray(C_Arr) And Not IsEmpty(C_Arr) Then Erase C_Arr
    C_Arr = ArrayArg
    Me.Reset
    On Error GoTo 0
End Property

Private Sub Default()
    HasRowIndex = False
    HasColIndex = False
    HasRun = False
End Sub

Public Property Get HasInitialArray() As Variant
    HasInitialArray = Not IsEmpty(C_Arr)
End Property

Public Property Get HasRowIndexes() As Boolean
    HasRowIndexes = IsArray(C_rowIndex) And Not IsEmpty(C_rowIndex)
End Property

Public Property Get HasColumnIndexes() As Boolean
    HasColumnIndexes = IsArray(C_colIndex) And Not IsEmpty(C_colIndex)
End Property

Public Property Get Dimension() As Long

    If HasRun Then Dimension = DimArr: Exit Sub
    Dim i As Long, n As Long

    If Me.HasInitialArray Then
        On Error Resume Next
        i = 1
        Do
            n = UBound(C_Arr, i)
            i = i + 1
        Loop While Err.Number = 0
        Err.Clear
        On Error GoTo 0

        Dimension = i - 1
    Else
        Dimension = 0
    End If
    
    HasRun = True
End Property

Public Property Get RowIndexes() As Variant
    If HasRowIndex Then
        RowIndexes = C_rowIndex
    ElseIf Me.HasInitialArray Then
        Dim r1 As Long, r2 As Long
        r1 = LBound(C_Arr, 1): r2 = UBound(C_Arr, 1)
        
        Dim retArr  As Variant
        ReDim retArr(r1 To r2)

        Dim fi As Long
        For fi = r1 To r2
            retArr(fi) = fi
        Next fi
        RowIndexes = retArr
    Else
        RowIndexes = Array()
    End If
End Property

Public Property Let RowIndexes(Index As Variant)
    C_rowIndex = Index
    HasRowIndex = True
End Property

Public Property Get ColumnIndexes() As Variant

    If HasColIndex Then
        ColumnIndexes = C_colIndex
    ElseIf Me.HasInitialArray And Me.Dimension > 1 Then
            
        Dim c1 As Long, c2 As Long
        c1 = LBound(C_Arr, 2): c2 = UBound(C_Arr, 2)
        
        Dim retArr() As Variant
        ReDim retArr(c1 To c2)
        
        Dim fi As Long
        For fi = c1 To c2
            retArr(fi) = fi
        Next fi
        ColumnIndexes = retArr
    Else
        ColumnIndexes = Array()
    End If
End Property

Public Property Let ColumnIndexes(Index As Variant)
    C_colIndex = Index
    HasColIndex = True
End Property

Public Property Get RowIndexCount() As Long
    On Error Resume Next
    If IsArray(C_rowIndex) And Not IsEmpty(C_rowIndex) Then
        RowIndexCount = UBound(C_rowIndex) - LBound(C_rowIndex) + 1
    Else
        RowIndexCount = UBound(C_Arr, 1) - LBound(C_Arr, 1) + 1
    End If
    On Error GoTo 0
End Property

Public Property Get ColIndexCount() As Long
    On Error Resume Next
    If IsArray(C_colIndex) And Not IsEmpty(C_colIndex) Then
        ColIndexCount = UBound(C_colIndex) - LBound(C_colIndex) + 1
    Else
        ColIndexCount = UBound(C_Arr, 2) - LBound(C_Arr, 2) + 1
    End If
    On Error GoTo 0
End Property

Public Property Get Value()

    Dim fi As Long, fj As Long, m As Long, n As Long
    Dim RowIndexes() As Long, colIndexes() As Long
    Dim retArr As Variant

    Select Case Me.Dimension
        Case 0
            retArr = Array()
        Case 1
            If Me.HasColIndex Then Err.Raise vbObjectError + 3001, , "1-D array can't have column indexes."
                        
            RowIndexes = Me.RowIndexes
            ReDim retArr(1 To Me.RowIndexCount)
            For fi = LBound(RowIndexes) To UBound(RowIndexes)
                m = m + 1
                retArr(m) = C_Arr(RowIndexes(fi))
            Next fi
        Case Else
            RowIndexes = Me.RowIndexes
            colIndexes = Me.ColumnIndexes
            ReDim retArr(1 To Me.RowIndexCount, 1 To Me.ColIndexCount)
            For fi = LBound(RowIndexes) To UBound(RowIndexes)
                m = m + 1
                n = 0
                For fj = LBound(colIndexes) To UBound(colIndexes)
                    n = n + 1
                    retArr(m, n) = C_Arr(RowIndexes(fi), colIndexes(fj))
                Next fj
            Next fi
    End Select
    
    Value = retArr
    
End Property

Public Property Let Value(NewValue As Variant)
    Dim ValArray() As Variant
    Dim fi As Long, fj As Long, n As Long
    Dim RowIndexes() As Long, colIndexes() As Long

    If IsArray(NewValue) Then
        ValArray = FlattenArray(NewValue)
    Else
        ValArray = Array(NewValue)
    End If
    
    If UBound(ValArray) - LBound(ValArray) + 1 < Me.RowIndexCount * Me.ColIndexCount Then
        Err.Raise vbObjectError + 3002, , "Value array is too small for target range."
    End If

    Select Case Me.Dimension
        Case 0
            Me.InitialArray = ValArray

        Case 1
            If Me.HasColIndex Then Err.Raise vbObjectError + 3001, , "1-D array can't have column indexes."
            n = LBound(ValArray)
            RowIndexes = Me.RowIndexes
            For fi = LBound(RowIndexes) To UBound(RowIndexes)
                C_Arr(RowIndexes(fi)) = ValArray(n)
                n = n + 1
            Next fi

        Case Else
            RowIndexes = Me.RowIndexes
            colIndexes = Me.ColumnIndexes
            
            n = LBound(ValArray)
            For fi = LBound(RowIndexes) To UBound(RowIndexes)
                For fj = LBound(colIndexes) To UBound(colIndexes)
                    C_Arr(RowIndexes(fi), colIndexes(fj)) = ValArray(n)
                    n = n + 1
                Next fj
            Next fi
    End Select
    
    Call ApplyChangesToParent
    
End Property

Public Property Get cell(Index As Long) As Variant
    Dim RowIndexes() As Long, colIndexes() As Long
    Dim quotient As Long, remainder As Long

    Select Case Me.Dimension
        Case 0
            ' Return an empty array for scalar/uninitialized case
            cell = Array()

        Case 1
            If Me.HasColIndex Then Err.Raise vbObjectError + 3001, , "1-D array can't have column indexes."
            RowIndexes = Me.RowIndexes
            If Index < 0 Or Index > UBound(RowIndexes) Then Err.Raise vbObjectError + 3002, , "Index out of bounds."
            cell = C_Arr(RowIndexes(Index))

        Case Else
            RowIndexes = Me.RowIndexes
            colIndexes = Me.ColumnIndexes
            If Index < 0 Or Index >= Me.RowIndexCount * Me.ColIndexCount Then Err.Raise vbObjectError + 3003, , "Index out of bounds."

            quotient = Index \ Me.ColIndexCount
            remainder = Index Mod Me.ColIndexCount
            cell = C_Arr(RowIndexes(quotient), colIndexes(remainder))
    End Select
End Property

Public Property Let cell(Index As Long, NewValue As Variant)
    Dim RowIndexes() As Long, colIndexes() As Long
    Dim quotient As Long, remainder As Long

    Select Case Me.Dimension
        Case 0
            ReDim C_Arr(Index To Index)
            C_Arr(Index) = NewValue

        Case 1
            If Me.HasColIndex Then Err.Raise vbObjectError + 3001, , "1-D array can't have column indexes."
            RowIndexes = Me.RowIndexes
            If Index < 0 Or Index > UBound(RowIndexes) Then Err.Raise vbObjectError + 3002, , "Index out of bounds."
            C_Arr(RowIndexes(Index)) = NewValue

        Case Else
            RowIndexes = Me.RowIndexes
            colIndexes = Me.ColumnIndexes
            If Index < 0 Or Index >= Me.RowIndexCount * Me.ColIndexCount Then Err.Raise vbObjectError + 3003, , "Index out of bounds."

            quotient = Index \ Me.ColIndexCount
            remainder = Index Mod Me.ColIndexCount
            C_Arr(RowIndexes(quotient), colIndexes(remainder)) = NewValue
    End Select
    
    Call ApplyChangesToParent
    
End Property

Public Sub Reset()

If Not IsEmpty(C_rowIndex) Then Erase C_rowIndex
If Not IsEmpty(C_colIndex) Then Erase C_colIndex

End Sub

Public Sub Add(RowIndex As Variant, Optional ColIndex As Variant)

If Not IsEmpty(C_rowIndex) Then Erase C_rowIndex
C_rowIndex = RowIndex

If IsMissing(ColIndex) Then Exit Sub
If Not IsEmpty(C_colIndex) Then Erase C_colIndex
C_colIndex = ColIndex

End Sub

Public Sub Clear()
    Dim fi As Long, fj As Long
    Dim RowIndexes() As Long
    Dim colIndexes() As Long

    Select Case Me.Dimension
        Case 0
            Exit Sub
        Case 1
            If Me.HasColIndex Then Err.Raise vbObjectError + 3001, , "1-D array can't have column indexes."

            RowIndexes = Me.RowIndexes
            For fi = LBound(RowIndexes) To UBound(RowIndexes)
                C_Arr(RowIndexes(fi)) = Empty
            Next fi

        Case Else
            RowIndexes = Me.RowIndexes
            colIndexes = Me.ColumnIndexes

            For fi = LBound(RowIndexes) To UBound(RowIndexes)
                For fj = LBound(colIndexes) To UBound(colIndexes)
                    C_Arr(RowIndexes(fi), colIndexes(fj)) = Empty
                Next fj
            Next fi
    End Select
    
    Call ApplyChangesToParent
    
End Sub

Public Function Count() As Long
    On Error Resume Next
    Dim rCount As Long, cCount As Long
    rCount = Me.RowIndexCount
    cCount = Me.ColIndexCount
    Count = rCount * cCount
    On Error GoTo 0
End Function

Public Function FilterdArray() As Variant
    Dim fi As Long, fj As Long
    Dim tempArr As Variant
    Dim r_Fil As Long, c_Fil As Long

    r_Fil = Me.RowIndexCount
    c_Fil = Me.ColIndexCount

    Select Case Me.Dimension
        Case 0
            FilterdArray = Array()
        Case 1
            If HasColIndex Then Err.Raise vbObjectError + 3001, , "1-D array can't have column indexes."

            If HasRowIndex Then
                ReDim tempArr(1 To r_Fil)
                For fi = 1 To r_Fil
                    tempArr(fi) = C_Arr(C_rowIndex(fi))
                Next fi
                FilterdArray = tempArr
            End If
        Case Else
            If Not (HasRowIndex And HasColIndex) Then
                FilterdArray = C_Arr
                Exit Function
            End If
            
            Dim r1_Arr As Long, c1_Arr As Long, r2_Arr As Long, c2_Arr As Long
            r1_Arr = LBound(C_Arr, 1): r2_Arr = UBound(C_Arr, 1)
            c1_Arr = LBound(C_Arr, 2): c2_Arr = UBound(C_Arr, 2)

            Dim totalRows As Long, totalCols As Long
            totalRows = IIf(HasRowIndex, r_Fil, r2_Arr - r1_Arr + 1)
            totalCols = IIf(HasColIndex, c_Fil, c2_Arr - c1_Arr + 1)

            ReDim tempArr(1 To totalRows, 1 To totalCols)
            For fi = 1 To totalRows
                For fj = 1 To totalCols
                    tempArr(fi, fj) = C_Arr(C_rowIndex(fi), C_colIndex(fj))
                Next fj
            Next fi
            FilterdArray = tempArr
    End Select
    
End Function

Public Function Filter()
    Me.InitialArray = Me.FilterdArray()
    ApplyChangesToParent
End Function

Private Function FlattenArray(ByVal inputArr As Variant) As Variant
    Dim i As Long, j As Long, k As Long
    Dim flatArr() As Variant

    If Not IsArray(inputArr) Then
        FlattenArray = Array()
        Exit Function
    End If

    On Error Resume Next
    Dim lb2 As Long: lb2 = LBound(inputArr, 2)
    If Err.Number <> 0 Then
        ' 1D array
        FlattenArray = inputArr
        Exit Function
    End If
    On Error GoTo 0

    ' 2D array: flatten it
    ReDim flatArr(1 To (UBound(inputArr, 1) - LBound(inputArr, 1) + 1) * _
                      (UBound(inputArr, 2) - LBound(inputArr, 2) + 1))

    k = 1
    For i = LBound(inputArr, 1) To UBound(inputArr, 1)
        For j = LBound(inputArr, 2) To UBound(inputArr, 2)
            flatArr(k) = inputArr(i, j)
            k = k + 1
        Next j
    Next i

    FlattenArray = flatArr
End Function

Private Sub ApplyChangesToParent()
    If Not pParent Is Nothing Then pParent.Update C_Arr
End Sub


