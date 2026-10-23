Attribute VB_Name = "modWorkbookContext"
Option Explicit

' Книга с макросами и книга с результатом — разные книги.
' Запускатель создаёт книгу результата; остальные процедуры работают с ней.
Public gRegisterWorkbook As Workbook

Public Sub StartRegisterAutomation()
    Dim sourceWb As Workbook
    Dim resultWb As Workbook
    Dim sourceData As Worksheet
    Dim oldAlerts As Boolean
    Dim dataCopied As Boolean

    On Error GoTo Failed
    Set sourceWb = ThisWorkbook
    oldAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Set resultWb = Workbooks.Add(xlWBATWorksheet)
    Set gRegisterWorkbook = resultWb

    On Error Resume Next
    Set sourceData = sourceWb.Worksheets("data")
    On Error GoTo Failed

    If Not sourceData Is Nothing Then
        sourceData.Copy Before:=resultWb.Worksheets(1)
        Application.DisplayAlerts = False
        resultWb.Worksheets(resultWb.Worksheets.Count).Delete
        Application.DisplayAlerts = oldAlerts
        dataCopied = True
    End If

    resultWb.Activate
    Application.ScreenUpdating = True

    If dataCopied Then
        MsgBox "Создана новая книга для перечня. Справочник data скопирован из книги с макросами.", _
               vbInformation, "Автоматизация перечня"
    Else
        MsgBox "Создана новая книга для перечня, но лист справочника data не найден в книге с макросами." & _
               vbCrLf & "Перечень можно сформировать, однако отдел и ответственные лица могут не заполниться автоматически.", _
               vbExclamation, "Нет справочника data"
    End If

    CreateMultiFolderApplicationsListCorrected
    Exit Sub

Failed:
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = True
    MsgBox "Не удалось подготовить книгу для перечня." & vbCrLf & _
           "Ошибка " & Err.Number & ": " & Err.Description, vbExclamation, "Автоматизация перечня"
End Sub

Public Function GetRegisterWorkbook() As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet

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

    If Not ActiveWorkbook Is Nothing Then
        Set gRegisterWorkbook = ActiveWorkbook
        Set GetRegisterWorkbook = ActiveWorkbook
    End If
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
