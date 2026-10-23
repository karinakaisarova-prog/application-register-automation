Attribute VB_Name = "modWorkbookContext"
Option Explicit

' Книга с макросами, справочником data и перечнем — одна и та же книга.
' Запускатель работает в ThisWorkbook и не создаёт отдельную книгу результата.
Public gRegisterWorkbook As Workbook

Public Sub StartRegisterAutomation()
    Dim sourceWb As Workbook
    Dim sourceData As Worksheet
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed

    Set sourceWb = ThisWorkbook
    On Error Resume Next
    Set sourceData = sourceWb.Worksheets("data")
    On Error GoTo Failed

    If sourceData Is Nothing Then
        MsgBox "В этой книге не найден лист с точным именем 'data'." & vbCrLf & _
               "Откройте подготовленную книгу, содержащую справочник data и макросы. " & _
               "Перечень будет создан в этой же книге.", _
               vbExclamation, "Нет справочника data"
        Exit Sub
    End If

    Set gRegisterWorkbook = sourceWb
    sourceWb.Activate

    MsgBox "Перечень будет создан в этой же книге, где находится лист data." & vbCrLf & _
           "Исходные документы останутся в выбранных папках; в Excel будут записаны ссылки на них.", _
           vbInformation, "Автоматизация перечня"

    CreateMultiFolderApplicationsListCorrected
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error GoTo 0
    MsgBox "Не удалось запустить формирование перечня." & vbCrLf & _
           "Ошибка " & failureNumber & ": " & failureDescription, _
           vbExclamation, "Автоматизация перечня"
End Sub

Public Function GetRegisterWorkbook() As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet

    ' Prefer the workbook explicitly selected by the startup procedure.
    On Error Resume Next
    If Not gRegisterWorkbook Is Nothing Then
        Set ws = gRegisterWorkbook.Worksheets(1)
        If Err.Number = 0 Then
            Set GetRegisterWorkbook = gRegisterWorkbook
            Exit Function
        End If
    End If
    Err.Clear
    On Error GoTo 0

    ' The macro workbook is the intended home for the data sheet and register.
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("data")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set gRegisterWorkbook = ThisWorkbook
        Set GetRegisterWorkbook = ThisWorkbook
        Exit Function
    End If

    ' Fallback after a VBA reset: identify a workbook by the register title/header.
    For Each wb In Application.Workbooks
        For Each ws In wb.Worksheets
            If CStr(ws.Range("A1").Value) = "Перечень приложений" Or _
               CStr(ws.Range("A2").Value) = "№ п/п" Then
                Set gRegisterWorkbook = wb
                Set GetRegisterWorkbook = wb
                Exit Function
            End If
        Next ws
    Next wb
End Function

Public Function GetActiveRegisterSheet() As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = GetRegisterWorkbook()
    If wb Is Nothing Then Exit Function

    If Not ActiveWorkbook Is Nothing Then
        If ActiveWorkbook Is wb Then
            If TypeName(ActiveSheet) = "Worksheet" Then
                If Trim$(CStr(ActiveSheet.Cells(2, 1).Value)) = "№ п/п" And _
                   Trim$(CStr(ActiveSheet.Cells(2, 6).Value)) = "Ответственное лицо" And _
                   Trim$(CStr(ActiveSheet.Cells(2, 8).Value)) = "Раздел" Then
                    Set GetActiveRegisterSheet = ActiveSheet
                    Exit Function
                End If
            End If
        End If
    End If

    For Each ws In wb.Worksheets
        If Trim$(CStr(ws.Cells(2, 1).Value)) = "№ п/п" And _
           Trim$(CStr(ws.Cells(2, 6).Value)) = "Ответственное лицо" And _
           Trim$(CStr(ws.Cells(2, 8).Value)) = "Раздел" Then
            Set GetActiveRegisterSheet = ws
            Exit Function
        End If
    Next ws
End Function
