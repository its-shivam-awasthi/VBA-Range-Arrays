VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private C_Dict As Object
Private C_Dict_Wb As Object
Private C_Dict_Ws As Object
Private C_Dict_Rng As Object
Private arrObj As Arrays

Public Enum vbExtendToRange
    vbAbsolute
    vbCurrentregion
    vbEndxlDownToxlRight
    vbUsedRegion
    vbLastRowLastColumn
End Enum

Public Enum vbHeaderOrientation
    vbNone = 0
    vbRowHeader = 1
    vbColumnHeader = 2
    vbBoth = 3
End Enum

Private Sub Class_Initialize()
    Set C_Dict = VBA.CreateObject("Scripting.Dictionary")
    Set C_Dict_Wb = VBA.CreateObject("Scripting.Dictionary")
    Set C_Dict_Ws = VBA.CreateObject("Scripting.Dictionary")
    Set C_Dict_Rng = VBA.CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    If Not C_Dict Is Nothing Then Set C_Dict = Nothing
    If Not C_Dict_Wb Is Nothing Then Set C_Dict_Wb = Nothing
    If Not C_Dict_Ws Is Nothing Then Set C_Dict_Ws = Nothing
    If Not C_Dict_Rng Is Nothing Then Set C_Dict_Rng = Nothing
End Sub

Public Sub ArraysFromRanges(Refrence_Name As String, _
                     Range_Reference As Variant, _
                     Optional Worksheet_Arg As Variant, _
                     Optional Workbook_Arg As Variant, _
                     Optional extendedRange As vbExtendToRange = vbAbsolute, _
                     Optional Has_Headers As Boolean = False, _
                     Optional HeaderOrientation As vbHeaderOrientation = vbRowHeader, _
                     Optional TrimEmptyRows As Boolean = False)

Dim c_wb As Workbook, c_ws As Worksheet, c_range As Range

Call Refresh_C_Dict

' Resolve Workbook
If Not IsMissing(Workbook_Arg) Then
    Select Case TypeName(Workbook_Arg)
        Case "Workbook"
            Set c_wb = Workbook_Arg
        Case "String"
            On Error Resume Next
            Set c_wb = Workbooks(Workbook_Arg)
            On Error GoTo 0
            If c_wb Is Nothing Then Err.Raise vbObjectError + 1000, , "Workbook not found: " & Workbook_Arg
        Case Else
            Err.Raise vbObjectError + 1001, , "Invalid Workbook_Arg type: " & TypeName(Workbook_Arg)
    End Select
Else
    Set c_wb = ActiveWorkbook
End If

' Resolve Worksheet
If Not IsMissing(Worksheet_Arg) Then
    Select Case TypeName(Worksheet_Arg)
        Case "Worksheet"
            Set c_ws = Worksheet_Arg
        Case "String", "Integer"
            On Error Resume Next
            Set c_ws = c_wb.Worksheets(Worksheet_Arg)
            On Error GoTo 0
            If c_ws Is Nothing Then Err.Raise vbObjectError + 1003, , "Worksheet not found: " & Worksheet_Arg
        Case Else
            Err.Raise vbObjectError + 1004, , "Invalid Worksheet_Arg type: " & TypeName(Worksheet_Arg)
    End Select
    
    If Not (c_ws.Parent Is c_wb Or IsMissing(Workbook_Arg)) Then
        Err.Raise vbObjectError + 1002, , "Worksheet does not belong to specified workbook."
    End If
    If IsMissing(Workbook_Arg) Then Set c_wb = c_ws.Parent
Else
    Set c_ws = c_wb.ActiveSheet
End If

' Resolve Range
Select Case TypeName(Range_Reference)
    Case "ListObject"
        Set c_range = Range_Reference.Range
    Case "Range"
        Set c_range = Range_Reference
    Case "String"
        On Error Resume Next
        Set c_range = c_ws.Range(Range_Reference)
        If c_range Is Nothing Then Set c_range = c_wb.Names(Range_Reference).RefersToRange
        On Error GoTo 0
        If c_range Is Nothing Then Err.Raise vbObjectError + 1007, , "Invalid range reference: " & Range_Reference
    Case Else
        Err.Raise vbObjectError + 1008, , "Invalid Range_Reference type: " & TypeName(Range_Reference)
End Select

If Not (c_range.Parent.Parent Is c_wb Or IsMissing(Workbook_Arg)) Then
    Err.Raise vbObjectError + 1005, , "Range does not belong to specified workbook."
End If
If IsMissing(Worksheet_Arg) Then Set c_ws = c_range.Parent
If Not (c_range.Parent Is c_ws Or IsMissing(Worksheet_Arg)) Then
    Err.Raise vbObjectError + 1006, , "Range does not belong to specified worksheet."
End If
If IsMissing(Workbook_Arg) Then Set c_wb = c_range.Parent.Parent

' Handle Extended Range
Dim extended_Range As Range
Select Case extendedRange
    Case vbAbsolute
        Set extended_Range = c_range
        
    Case vbCurrentregion
        Set extended_Range = c_range.CurrentRegion
        
    Case vbEndxlDownToxlRight
        Set extended_Range = c_range.End(xlDown).End(xlToRight)
    
    Case vbUsedRegion
        Dim UsedRng As Range
        Set UsedRng = c_ws.UsedRange
        
        Dim ur As Long, uc As Long
        With UsedRng
            ur = .row + .Rows.Count - 1
            uc = .Column + .Columns.Count - 1
        End With
        
        Dim cr As Long, cc As Long
        With c_range
            cr = .row + .Rows.Count - 1
            cc = .Column + .Columns.Count - 1
        End With
        
        If ur < c_range.row Or uc < c_range.Column Then
            Err.Raise vbObjectError + 1009, , "Cannot extend range, range outside extended parameters."
        End If
        
        Dim newRow As Long: newRow = IIf(ur > cr, ur, cr) - c_range.row + 1
        Dim newCol As Long: newCol = IIf(uc > cc, uc, cc) - c_range.Column + 1
        
        Set extended_Range = c_range.Resize(newRow, newCol)
        
    Case vbLastRowLastColumn
        
        Dim lr As Long, lc As Long
        
        On Error Resume Next
        lr = c_ws.Cells.Find("*", , , , xlByRows, xlPrevious).row
        lc = c_ws.Cells.Find("*", , , , xlByColumns, xlPrevious).Column
        On Error GoTo 0

        If lr < c_range.row Or lc < c_range.Column Then
            Err.Raise vbObjectError + 1009, , "Cannot extend range, range outside extended parameters."
        End If

        Set extended_Range = c_ws.Range(c_ws.Cells(c_range.row, c_range.Column), c_ws.Cells(lr, lc))
 End Select
        
If Has_Headers Then Set extended_Range = Exclude_Header(extended_Range, HeaderOrientation, TrimEmptyRows)

C_Dict_Rng.Add Refrence_Name, extended_Range.Address
C_Dict.Add Refrence_Name, extended_Range.Value
C_Dict_Wb.Add Refrence_Name, c_wb.FullName
C_Dict_Ws.Add Refrence_Name, c_ws.Name

Call Refresh_arrObj

Set c_wb = Nothing
Set c_ws = Nothing
Set c_range = Nothing
Set extended_Range = Nothing

End Sub

Private Function Exclude_Header(Arg As Range, _
                                Optional HeaderOrientation As vbHeaderOrientation = vbRowHeader, _
                                Optional TrimEmptyRows As Boolean = False) As Range
    Dim r As Long, c As Long
    Dim RowCount As Long, colCount As Long
    Dim ws As Worksheet
    Set ws = Arg.Worksheet

    RowCount = Arg.Rows.Count
    colCount = Arg.Columns.Count

    ' Trim trailing empty rows if requested
    If TrimEmptyRows Then
        Do While RowCount > 1 And WorksheetFunction.CountA(Arg.Rows(RowCount)) = 0
            RowCount = RowCount - 1
        Loop
    End If

    ' Validate header exclusion
    If HeaderOrientation = vbColumnHeader Or HeaderOrientation = vbBoth Then
        If RowCount <= 1 Then Err.Raise vbObjectError + 1012, , "Cannot exclude row header from a single-row range."
        r = Arg.row + 1
        RowCount = RowCount - 1
    Else
        r = Arg.row
    End If

    If HeaderOrientation = vbRowHeader Or HeaderOrientation = vbBoth Then
        If colCount <= 1 Then Err.Raise vbObjectError + 1013, , "Cannot exclude column header from a single-column range."
        c = Arg.Column + 1
        colCount = colCount - 1
    Else
        c = Arg.Column
    End If

    Set Exclude_Header = ws.Range(ws.Cells(r, c), ws.Cells(r + RowCount - 1, c + colCount - 1))
End Function

' Property Get to retrieve array by reference name
Public Property Get RangeArray(ByVal ReferenceName As String) As Variant
Call Refresh_C_Dict
If C_Dict.Exists(ReferenceName) Then
    RangeArray = C_Dict(ReferenceName)
Else
    Err.Raise vbObjectError + 1010, , "Reference name not found: " & ReferenceName
End If
End Property

' Property Let to update or insert array by reference name
Public Property Let RangeArray(ByVal ReferenceName As String, ByVal NewArray As Variant)

Call Refresh_C_Dict
If C_Dict.Exists(ReferenceName) Then
    C_Dict(ReferenceName) = NewArray
Else
    C_Dict.Add ReferenceName, NewArray
End If
Call Refresh_arrObj

End Property

Public Function Exists(ByVal ReferenceName As String) As Boolean
Call Refresh_C_Dict
Exists = C_Dict.Exists(ReferenceName)
End Function

' Method Add to insert array by reference name
Public Sub Add(ByVal ReferenceName As String, ByVal NewArray As Variant, _
               Optional Overwrite As Boolean = False, _
               Optional WorkbookName As String = "", _
               Optional WorksheetName As String = "", _
               Optional RangeAddress As String = "")

    Call Refresh_C_Dict

    If C_Dict.Exists(ReferenceName) Then
        If Overwrite Then
            C_Dict(ReferenceName) = NewArray
            If WorkbookName <> "" Then C_Dict_Wb(ReferenceName) = WorkbookName
            If WorksheetName <> "" Then C_Dict_Ws(ReferenceName) = WorksheetName
            If RangeAddress <> "" Then C_Dict_Rng(ReferenceName) = RangeAddress
        Else
            Err.Raise vbObjectError + 1011, , "Reference name already exists: " & ReferenceName
        End If
    Else
        C_Dict.Add ReferenceName, NewArray
        If WorkbookName <> "" Then C_Dict_Wb.Add ReferenceName, WorkbookName
        If WorksheetName <> "" Then C_Dict_Ws.Add ReferenceName, WorksheetName
        If RangeAddress <> "" Then C_Dict_Rng.Add ReferenceName, RangeAddress
    End If

    Call Refresh_arrObj
End Sub

' Method Remove to Remove array by reference name
Public Sub Remove(ByVal ReferenceName As String)
    Call Refresh_C_Dict

    If C_Dict.Exists(ReferenceName) Then
        C_Dict.Remove ReferenceName

        If C_Dict_Wb.Exists(ReferenceName) Then C_Dict_Wb.Remove ReferenceName
        If C_Dict_Ws.Exists(ReferenceName) Then C_Dict_Ws.Remove ReferenceName
        If C_Dict_Rng.Exists(ReferenceName) Then C_Dict_Rng.Remove ReferenceName

        Call Refresh_arrObj
    Else
        Err.Raise vbObjectError + 1012, , "Reference doesn't exist or was already removed: " & ReferenceName
    End If
End Sub

Public Function Count() As Long
Call Refresh_C_Dict
Count = C_Dict.Count
End Function

Public Sub ClearAll()
    Call Refresh_C_Dict

    If Not C_Dict Is Nothing Then
        C_Dict.RemoveAll
    Else
        Err.Raise vbObjectError + 1013, , "Main dictionary not initialized."
    End If

    If Not C_Dict_Wb Is Nothing Then C_Dict_Wb.RemoveAll
    If Not C_Dict_Ws Is Nothing Then C_Dict_Ws.RemoveAll
    If Not C_Dict_Rng Is Nothing Then C_Dict_Rng.RemoveAll

    Call Refresh_arrObj
End Sub

Public Property Get Arrays(ByVal ReferenceName As String) As Arrays

' Refresh the dictionary before accessing
Call Refresh_C_Dict

' Initialize the private arrObj variable
If Not arrObj Is Nothing Then Set arrObj = Nothing
Set arrObj = New Arrays

' Check if the reference exists in the dictionary
If C_Dict.Exists(ReferenceName) Then
    Set arrObj.Arrays = C_Dict  ' Assign the full dictionary
    arrObj.UseRef ReferenceName   ' Assign the key to be used
    Set Arrays = arrObj         ' Return the object
Else
    Err.Raise vbObjectError + 1015, , "Reference name not found: " & ReferenceName
End If

End Property

Public Function ArraySummary(ByVal ReferenceName As String) As String

If C_Dict.Exists(ReferenceName) Then
    Dim tempArr As Arrays
    Set tempArr = Me.Arrays(ReferenceName)
    ArraySummary = tempArr.ToString
Else
    Err.Raise vbObjectError + 1010, , "Reference name not found: " & ReferenceName
End If

End Function

Private Sub Refresh_C_Dict()

If Not arrObj Is Nothing Then Set C_Dict = arrObj.Arrays

End Sub

Private Sub Refresh_arrObj()

If Not arrObj Is Nothing Then Set arrObj.Arrays = C_Dict

End Sub

Public Sub ExportToRange(ByVal ReferenceName As String, _
                         ByVal targetRange As Range, _
                         Optional ClearBeforeWrite As Boolean = True)

Call Refresh_C_Dict

If Not C_Dict.Exists(ReferenceName) Then
    Err.Raise vbObjectError + 1014, , "Reference not found: " & ReferenceName
End If

Dim arr As Variant
arr = C_Dict(ReferenceName)

If Not IsArray(arr) Then
    Err.Raise vbObjectError + 1015, , "Referenced item is not an array."
End If

If ClearBeforeWrite Then
    targetRange.Resize(UBound(arr, 1), UBound(arr, 2)).ClearContents
End If

targetRange.Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr

End Sub

Public Sub RenameReference(ByVal OldName As String, ByVal NewName As String)
Call Refresh_C_Dict

If Not C_Dict.Exists(OldName) Then
    Err.Raise vbObjectError + 1016, , "Old reference name not found: " & OldName
End If

If C_Dict.Exists(NewName) Then
    Err.Raise vbObjectError + 1017, , "New reference name already exists: " & NewName
End If

' Rename in main dictionary
C_Dict.Add NewName, C_Dict(OldName)
C_Dict.Remove OldName

' Rename in metadata dictionaries
If C_Dict_Wb.Exists(OldName) Then
    C_Dict_Wb.Add NewName, C_Dict_Wb(OldName)
    C_Dict_Wb.Remove OldName
End If

If C_Dict_Ws.Exists(OldName) Then
    C_Dict_Ws.Add NewName, C_Dict_Ws(OldName)
    C_Dict_Ws.Remove OldName
End If

If C_Dict_Rng.Exists(OldName) Then
    C_Dict_Rng.Add NewName, C_Dict_Rng(OldName)
    C_Dict_Rng.Remove OldName
End If

Call Refresh_arrObj
End Sub

Public Function ListReferences() As Variant
Call Refresh_C_Dict

If C_Dict.Count = 0 Then
    ListReferences = Array()
    Exit Function
End If

Dim keys() As Variant
keys = C_Dict.keys
ListReferences = keys
End Function

Public Sub CloneReference(ByVal SourceName As String, ByVal TargetName As String)
Call Refresh_C_Dict

If Not C_Dict.Exists(SourceName) Then
    Err.Raise vbObjectError + 1018, , "Source reference not found: " & SourceName
End If

If C_Dict.Exists(TargetName) Then
    Err.Raise vbObjectError + 1019, , "Target reference already exists: " & TargetName
End If

C_Dict.Add TargetName, C_Dict(SourceName)
C_Dict_Wb.Add TargetName, C_Dict_Wb(SourceName)
C_Dict_Ws.Add TargetName, C_Dict_Ws(SourceName)
C_Dict_Rng.Add TargetName, C_Dict_Rng(SourceName)

Call Refresh_arrObj
End Sub

Public Function GetWorkbook(ByVal ReferenceName As String) As Workbook
If Not C_Dict_Wb.Exists(ReferenceName) Then
    Err.Raise vbObjectError + 1020, , "Workbook info not found for: " & ReferenceName
End If

Dim wbName As String
wbName = C_Dict_Wb(ReferenceName)

On Error Resume Next
Set GetWorkbook = Workbooks(wbName)
On Error GoTo 0

If GetWorkbook Is Nothing Then
    Err.Raise vbObjectError + 1021, , "Workbook is not open: " & wbName
End If
End Function

Public Function GetWorksheet(ByVal ReferenceName As String) As Worksheet
Dim wb As Workbook
Set wb = GetWorkbook(ReferenceName)

If Not C_Dict_Ws.Exists(ReferenceName) Then
    Err.Raise vbObjectError + 1022, , "Worksheet info not found for: " & ReferenceName
End If

Dim wsName As String
wsName = C_Dict_Ws(ReferenceName)

On Error Resume Next
Set GetWorksheet = wb.Worksheets(wsName)
On Error GoTo 0

If GetWorksheet Is Nothing Then
    Err.Raise vbObjectError + 1023, , "Worksheet not found in workbook: " & wsName
End If
End Function

Public Function GetRange(ByVal ReferenceName As String) As Range
Dim ws As Worksheet
Set ws = GetWorksheet(ReferenceName)

If Not C_Dict_Rng.Exists(ReferenceName) Then
    Err.Raise vbObjectError + 1024, , "Range info not found for: " & ReferenceName
End If

Dim rngAddress As String
rngAddress = C_Dict_Rng(ReferenceName)

On Error Resume Next
Set GetRange = ws.Range(rngAddress)
On Error GoTo 0

If GetRange Is Nothing Then
    Err.Raise vbObjectError + 1025, , "Unable to resolve range: " & rngAddress
End If
End Function
