Attribute VB_Name = "MainModule"
Option Explicit

Public Sub BrowseFile()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Válassza ki a hallgatói adatokat tartalmazó Excel fájlt"
        .Filters.Clear
        .Filters.Add "Excel fájlok", "*.xlsx; *.xls; *.xlsm"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            Range("StudentFilePath").Value = .SelectedItems(1)
        End If
    End With
End Sub

Public Sub GenerateReport()
    Dim filePath As String
    Dim wsControl As Worksheet
    Dim tblCourses As ListObject
    Dim courses As Variant
    Dim studentData As Variant
    Dim uniqueStudents As Object
    
    Set wsControl = ThisWorkbook.Sheets("Vezérlőpult")
    filePath = wsControl.Range("StudentFilePath").Value
    
    ' Validation
    If filePath = "" Or Dir(filePath) = "" Then
        MsgBox "Kérem válasszon ki egy érvényes forrásfájlt!", vbExclamation
        Exit Sub
    End If
    
    Set tblCourses = wsControl.ListObjects("KurzusLista")
    If tblCourses.ListRows.Count = 0 Then
        MsgBox "Nincsenek megadva kurzusok!", vbExclamation
        Exit Sub
    End If
    
    ' Read courses
    courses = tblCourses.DataBodyRange.Value
    
    ' Check for empty table content
    Dim hasContent As Boolean
    Dim i As Long
    hasContent = False
    
    For i = LBound(courses, 1) To UBound(courses, 1)
        If Trim(CStr(courses(i, 1))) <> "" Then
            hasContent = True
            Exit For
        End If
    Next i
    
    If Not hasContent Then
        MsgBox "Nincsenek megadva kurzusok!", vbExclamation
        Exit Sub
    End If
    
    ' Check for missing grading types
    Dim missingGrading As String
    missingGrading = ""
    
    For i = LBound(courses, 1) To UBound(courses, 1)
        If Trim(CStr(courses(i, 1))) <> "" Then
            If Trim(CStr(courses(i, 2))) = "" Then
                missingGrading = missingGrading & courses(i, 1) & ", "
            End If
        End If
    Next i
    
    If missingGrading <> "" Then
        If MsgBox("A következő tárgyaknál nincs megadva bejegyzés típus: " & vbCrLf & _
                  Left(missingGrading, Len(missingGrading) - 2) & vbCrLf & vbCrLf & _
                  "Ezeknél a program 'Évközi jegy' módban fog futni." & vbCrLf & _
                  "Szeretné folytatni?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Adatok betöltése..."
    
    ' Load Data
    On Error GoTo ErrorHandler
    studentData = DataModule.LoadStudentData(filePath)
    
    ' Validate Columns
    If Not DataModule.ValidateColumns(studentData) Then Exit Sub
    
    Application.StatusBar = "Hallgatók feldolgozása..."
    Set uniqueStudents = DataModule.GetUniqueStudents(studentData)
    
    Application.StatusBar = "Kimutatás készítése..."
    ReportModule.CreateReport uniqueStudents, courses, studentData, filePath
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Hiba történt: " & Err.Description, vbCritical
End Sub
