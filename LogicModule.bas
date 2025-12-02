Attribute VB_Name = "LogicModule"
Option Explicit

' Returns Array(Grade, IsRecognized)
Public Function GetStudentGrade(studentData As Variant, neptun As String, courseCode As String, cols As StudentColIndices) As Variant
    Dim i As Long
    Dim bestDate As Date
    Dim bestGrade As String
    Dim bestRecog As Boolean
    Dim found As Boolean
    Dim currentDate As Date
    
    found = False
    bestDate = CDate(0)
    
    For i = 2 To UBound(studentData, 1)
        If studentData(i, cols.Neptun) = neptun And studentData(i, cols.Subject) = courseCode Then
            ' Check Validity (Must be True/Igaz)
            If IsValid(studentData(i, cols.Valid)) Then
                currentDate = CDate(studentData(i, cols.EntryDate))
                If Not found Or currentDate > bestDate Then
                    bestDate = currentDate
                    bestGrade = CStr(studentData(i, cols.EntryValue))
                    bestRecog = IsRecognized(studentData(i, cols.Recognized))
                    found = True
                End If
            End If
        End If
    Next i
    
    If found Then
        GetStudentGrade = Array(bestGrade, bestRecog)
    Else
        GetStudentGrade = Array("", False)
    End If
End Function

' Returns Array(Signature, Exam, SigRecognized, ExamRecognized)
Public Function GetStudentSignatureAndExam(studentData As Variant, neptun As String, courseCode As String, cols As StudentColIndices) As Variant
    Dim i As Long
    
    ' 1. Find Latest Signature (Aláírás) - IGNORE VALIDITY
    Dim sigFound As Boolean
    Dim sigDate As Date, bestSigDate As Date
    Dim sigValue As String
    Dim sigRecog As Boolean
    
    sigFound = False
    bestSigDate = CDate(0)
    
    For i = 2 To UBound(studentData, 1)
        If studentData(i, cols.Neptun) = neptun And studentData(i, cols.Subject) = courseCode Then
            If studentData(i, cols.EntryType) = "Aláírás" Then
                ' CRITICAL: Do NOT check IsValid for Aláírás
                sigDate = CDate(studentData(i, cols.EntryDate))
                If Not sigFound Or sigDate > bestSigDate Then
                    bestSigDate = sigDate
                    sigValue = CStr(studentData(i, cols.EntryValue))
                    sigRecog = IsRecognized(studentData(i, cols.Recognized))
                    sigFound = True
                End If
            End If
        End If
    Next i
    
    If Not sigFound Then
        GetStudentSignatureAndExam = Array("", "", False, False)
        Exit Function
    End If
    
    ' If Signature is "Megtagadva", return immediately
    If sigValue = "Megtagadva" Then
        GetStudentSignatureAndExam = Array(sigValue, "", sigRecog, False)
        Exit Function
    End If
    
    ' 2. If "Aláírva", find Latest Valid Exam (Vizsgajegy)
    Dim examFound As Boolean
    Dim examDate As Date, bestExamDate As Date
    Dim examValue As String
    Dim examRecog As Boolean
    
    examFound = False
    bestExamDate = CDate(0)
    
    If sigValue = "Aláírva" Then
        For i = 2 To UBound(studentData, 1)
            If studentData(i, cols.Neptun) = neptun And studentData(i, cols.Subject) = courseCode Then
                If studentData(i, cols.EntryType) = "Vizsgajegy" Then
                    ' Check Validity for Exam
                    If IsValid(studentData(i, cols.Valid)) Then
                        examDate = CDate(studentData(i, cols.EntryDate))
                        If Not examFound Or examDate > bestExamDate Then
                            bestExamDate = examDate
                            examValue = CStr(studentData(i, cols.EntryValue))
                            examRecog = IsRecognized(studentData(i, cols.Recognized))
                            examFound = True
                        End If
                    End If
                End If
            End If
        Next i
    End If
    
    If examFound Then
        GetStudentSignatureAndExam = Array(sigValue, examValue, sigRecog, examRecog)
    Else
        GetStudentSignatureAndExam = Array(sigValue, "", sigRecog, False)
    End If

End Function

Private Function IsValid(val As Variant) As Boolean
    Dim s As String
    s = LCase(CStr(val))
    IsValid = (s = "igaz" Or s = "true")
End Function

Private Function IsRecognized(val As Variant) As Boolean
    Dim s As String
    s = LCase(CStr(val))
    IsRecognized = (s = "igaz" Or s = "true")
End Function
