Attribute VB_Name = "ReportModule"
Option Explicit

Public Sub CreateReport(uniqueStudents As Object, courses As Variant, studentData As Variant, sourceFilePath As String)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim r As Long, c As Long
    Dim studentKey As Variant
    Dim studentInfo As Variant
    Dim courseCode As String
    Dim gradingType As String
    Dim result As Variant
    Dim i As Long
    Dim sheetName As String
    Dim fso As Object
    Dim fileName As String
    Dim cols As StudentColIndices
    
    Set wb = ThisWorkbook
    
    ' Get Column Indices
    cols = DataModule.GetStudentColIndices(studentData)
    
    ' Generate Sheet Name
    Set fso = CreateObject("Scripting.FileSystemObject")
    fileName = fso.GetBaseName(sourceFilePath)
    
    ' Format: HHMM_FileName (Max 31 chars)
    sheetName = Format(Now, "hhmm") & "_" & fileName
    If Len(sheetName) > 31 Then
        sheetName = Left(sheetName, 31)
    End If
    
    ' Create new sheet
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    
    On Error Resume Next
    ws.Name = sheetName
    If Err.Number <> 0 Then
        ' Fallback if name exists or invalid
        ws.Name = Left("Rep_" & sheetName, 31)
    End If
    On Error GoTo 0
    
    ' Headers
    Dim headers As Variant
    headers = Array("Modulkód", "Felvétel féléve", "Neptun kód", "Nyomtatási név", "Felvételi összes pontszám", "Státusz")
    
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
        ws.Cells(1, i + 1).Font.Bold = True
        ws.Cells(1, i + 1).Borders.LineStyle = xlContinuous
    Next i
    
    ' Write Student Data
    r = 2
    For Each studentKey In uniqueStudents.Keys
        studentInfo = uniqueStudents(studentKey)
        ws.Cells(r, 1).Value = studentInfo(1)
        ws.Cells(r, 2).Value = studentInfo(2)
        ws.Cells(r, 3).Value = studentInfo(3) ' Neptun
        ws.Cells(r, 4).Value = studentInfo(4)
        
        ' Borders for student info
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)).Borders.LineStyle = xlContinuous
        r = r + 1
    Next studentKey
    
    ' Process Courses
    Dim startCol As Long
    startCol = 7
    Dim colIdx As Long
    colIdx = startCol
    
    Dim color1 As Long, color2 As Long
    color1 = RGB(217, 225, 242) ' Light Blue 1
    color2 = RGB(180, 198, 231) ' Light Blue 2
    Dim currentColor As Long
    Dim courseIdx As Long
    
    For courseIdx = 1 To UBound(courses, 1)
        courseCode = courses(courseIdx, 1)
        gradingType = Trim(courses(courseIdx, 2))
        
        If courseIdx Mod 2 = 1 Then currentColor = color1 Else currentColor = color2
        
        If gradingType = "Aláírás és Vizsgajegy" Then
            ' Merge Header
            With ws.Range(ws.Cells(1, colIdx), ws.Cells(1, colIdx + 1))
                .Merge
                .Value = courseCode
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = currentColor
                .Borders.LineStyle = xlContinuous
            End With
            
            ' Process Data
            r = 2
            For Each studentKey In uniqueStudents.Keys
                result = LogicModule.GetStudentSignatureAndExam(studentData, CStr(studentKey), courseCode, cols)
                ' result: 0=Sig, 1=Exam, 2=SigRecog, 3=ExamRecog
                
                ' Signature Cell
                With ws.Cells(r, colIdx)
                    .Value = result(0)
                    .Borders.LineStyle = xlContinuous
                    .Interior.Color = currentColor
                    
                    If result(0) = "" Then
                        .Interior.Color = RGB(255, 255, 0) ' Yellow if missing
                    ElseIf result(2) And result(3) Then
                         ' Green if both recognized? Or just if Sig recognized?
                         ' Python logic: if sig_recognized and exam_recognized: cell1.fill = green_fill
                         .Interior.Color = RGB(146, 208, 80)
                    End If
                End With
                
                ' Exam Cell
                With ws.Cells(r, colIdx + 1)
                    .Value = result(1)
                    .Borders.LineStyle = xlContinuous
                    .Interior.Color = currentColor
                    
                    If result(0) = "" Then
                         .Interior.Color = RGB(255, 255, 0)
                    ElseIf result(2) And result(3) Then
                         .Interior.Color = RGB(146, 208, 80)
                    End If
                End With
                
                r = r + 1
            Next studentKey
            
            ws.Columns(colIdx).ColumnWidth = 11.33
            ws.Columns(colIdx + 1).ColumnWidth = 11.33
            
            colIdx = colIdx + 2
        Else
            ' Single Column
            With ws.Cells(1, colIdx)
                .Value = courseCode
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Bold = True
                .Interior.Color = currentColor
                .Borders.LineStyle = xlContinuous
            End With
            
            ' Process Data
            r = 2
            For Each studentKey In uniqueStudents.Keys
                result = LogicModule.GetStudentGrade(studentData, CStr(studentKey), courseCode, cols)
                ' result: 0=Grade, 1=Recognized
                
                With ws.Cells(r, colIdx)
                    .Value = result(0)
                    .Borders.LineStyle = xlContinuous
                    .Interior.Color = currentColor
                    
                    If result(0) = "" Then
                        .Interior.Color = RGB(255, 255, 0)
                    ElseIf result(1) Then
                        .Interior.Color = RGB(146, 208, 80)
                    End If
                End With
                
                r = r + 1
            Next studentKey
            
            ws.Columns(colIdx).ColumnWidth = 13.44
            
            colIdx = colIdx + 1
        End If
    Next courseIdx
    
    ' AutoFit Student Data Columns
    ws.Range(ws.Columns(1), ws.Columns(6)).AutoFit
    
    ws.Activate
End Sub
