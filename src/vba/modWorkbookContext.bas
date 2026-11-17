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
    Set sourceData = TryGetWorksheet(sourceWb, "data")

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
    If failureNumber = 18 Then
        MsgBox "Операция отменена.", vbInformation, "Автоматизация перечня"
    Else
        MsgBox "Не удалось запустить формирование перечня." & vbCrLf & _
               "Ошибка " & failureNumber & ": " & failureDescription, _
               vbExclamation, "Автоматизация перечня"
    End If
End Sub

Public Function GetRegisterWorkbook() As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet

    If Not gRegisterWorkbook Is Nothing Then
        Set ws = TryGetWorksheet(gRegisterWorkbook, "data")
        If Not ws Is Nothing Then
            Set GetRegisterWorkbook = gRegisterWorkbook
            Exit Function
        End If
        Set gRegisterWorkbook = Nothing
    End If

    Set ws = TryGetWorksheet(ThisWorkbook, "data")
    If Not ws Is Nothing Then
        Set gRegisterWorkbook = ThisWorkbook
        Set GetRegisterWorkbook = ThisWorkbook
        Exit Function
    End If

    If Not ActiveWorkbook Is Nothing Then
        Set ws = TryGetWorksheet(ActiveWorkbook, "data")
        If Not ws Is Nothing Then
            Set gRegisterWorkbook = ActiveWorkbook
            Set GetRegisterWorkbook = ActiveWorkbook
            Exit Function
        End If
    End If

    For Each wb In Application.Workbooks
        Set ws = TryGetWorksheet(wb, "data")
        If Not ws Is Nothing Then
            Set gRegisterWorkbook = wb
            Set GetRegisterWorkbook = wb
            Exit Function
        End If
    Next wb

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

    Err.Raise vbObjectError + 2101, "GetRegisterWorkbook", _
              "Не удалось определить книгу для перечня. Откройте книгу с листом 'data' и запустите макрос из неё."
End Function

Private Function TryGetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim lookupError As Long

    If wb Is Nothing Then Exit Function

    On Error Resume Next
    Err.Clear
    Set TryGetWorksheet = wb.Worksheets(sheetName)
    lookupError = Err.Number
    On Error GoTo 0

    If lookupError = 18 Then Err.Raise 18
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
