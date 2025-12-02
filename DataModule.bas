Attribute VB_Name = "DataModule"
Option Explicit

Public Type StudentColIndices
    Neptun As Long
    Subject As Long
    Valid As Long
    EntryDate As Long
    EntryValue As Long
    Recognized As Long
    EntryType As Long
End Type

Public Function LoadStudentData(filePath As String) As Variant
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim data As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim wasOpen As Boolean
    Dim fileName As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = fso.GetFileName(filePath)
    
    ' Check if workbook is already open in this instance
    wasOpen = False
    On Error Resume Next
    Set wbSource = Workbooks(fileName)
    On Error GoTo 0
    
    If wbSource Is Nothing Then
        Set wbSource = Workbooks.Open(filePath, ReadOnly:=True)
    Else
        wasOpen = True
    End If
    
    Set wsSource = wbSource.Sheets(1) ' Assume data is on first sheet
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
    ' Read all data into array
    If lastRow > 1 Then
        data = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, lastCol)).Value
    End If
    
    If Not wasOpen Then
        wbSource.Close SaveChanges:=False
    End If
    
    ' Map columns dynamically based on headers in row 1
    ' Note: For simplicity in this macro, we will assume standard column order or map them here.
    ' To make it robust, we should find column indices by name.
    ' For now, we'll assume the user's file matches the Python script's expected columns.
    ' If needed, we can add a header mapping function.
    
    LoadStudentData = data
End Function

Public Function GetUniqueStudents(data As Variant) As Object
    Dim dict As Object
    Dim i As Long
    Dim key As String
    Dim studentInfo As Variant
    Dim result As Collection
    Dim colNeptun As Long, colModule As Long, colSemester As Long, colName As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    colNeptun = GetColIndex("Neptun kód", data)
    colModule = GetColIndex("Modulkód", data)
    colSemester = GetColIndex("Felvétel féléve", data)
    colName = GetColIndex("Nyomtatási név", data)
    
    ' Data is 1-based 2D array. Row 1 is header.
    For i = 2 To UBound(data, 1)
        ' Key: Neptun Code
        key = CStr(data(i, colNeptun))
        
        If Not dict.Exists(key) Then
            ' Store basic info: Modul, Semester, Neptun, Name
            ReDim studentInfo(1 To 4)
            studentInfo(1) = data(i, colModule)
            studentInfo(2) = data(i, colSemester)
            studentInfo(3) = key
            studentInfo(4) = data(i, colName)
            dict.Add key, studentInfo
        End If
    Next i
    
    Set GetUniqueStudents = dict
End Function

Public Function GetStudentColIndices(data As Variant) As StudentColIndices
    Dim idx As StudentColIndices
    idx.Neptun = GetColIndex("Neptun kód", data)
    idx.Subject = GetColIndex("Tárgykód", data)
    idx.Valid = GetColIndex("Érvényes", data)
    idx.EntryDate = GetColIndex("Bejegyzés dátuma", data)
    idx.EntryValue = GetColIndex("Bejegyzés értéke", data)
    idx.Recognized = GetColIndex("Elismert", data)
    idx.EntryType = GetColIndex("Bejegyzés típusa", data)
    GetStudentColIndices = idx
End Function

Public Function ValidateColumns(data As Variant) As Boolean
    Dim missing As String
    missing = ""
    
    If GetColIndex("Neptun kód", data) = 0 Then missing = missing & "Neptun kód, "
    If GetColIndex("Modulkód", data) = 0 Then missing = missing & "Modulkód, "
    If GetColIndex("Felvétel féléve", data) = 0 Then missing = missing & "Felvétel féléve, "
    If GetColIndex("Nyomtatási név", data) = 0 Then missing = missing & "Nyomtatási név, "
    If GetColIndex("Tárgykód", data) = 0 Then missing = missing & "Tárgykód, "
    If GetColIndex("Érvényes", data) = 0 Then missing = missing & "Érvényes, "
    If GetColIndex("Bejegyzés dátuma", data) = 0 Then missing = missing & "Bejegyzés dátuma, "
    If GetColIndex("Bejegyzés értéke", data) = 0 Then missing = missing & "Bejegyzés értéke, "
    If GetColIndex("Elismert", data) = 0 Then missing = missing & "Elismert, "
    If GetColIndex("Bejegyzés típusa", data) = 0 Then missing = missing & "Bejegyzés típusa, "
    
    If missing <> "" Then
        MsgBox "A következő oszlopok hiányoznak a fájlból: " & vbCrLf & Left(missing, Len(missing) - 2), vbCritical
        ValidateColumns = False
    Else
        ValidateColumns = True
    End If
End Function

Public Function GetColIndex(headerName As String, data As Variant) As Long
    Dim j As Long
    For j = 1 To UBound(data, 2)
        If LCase(Trim(data(1, j))) = LCase(Trim(headerName)) Then
            GetColIndex = j
            Exit Function
        End If
    Next j
    ' Fallback or Error
    GetColIndex = 0
End Function
