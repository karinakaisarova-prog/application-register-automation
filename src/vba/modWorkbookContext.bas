Attribute VB_Name = "modWorkbookContext"
Option Explicit

' Книга с макросами и книга с результатом — разные книги.
' Запускатель создаёт книгу результата; остальные процедуры работают с ней.
Public gRegisterWorkbook As Workbook

Public Sub StartRegisterAutomation()
    Dim sourceWb As Workbook
    Dim resultWb As Workbook
    Dim dataWb As Workbook
    Dim sourceData As Worksheet
    Dim oldAlerts As Boolean
    Dim oldScreen As Boolean
    Dim dataPath As Variant

    oldAlerts = Application.DisplayAlerts
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed

    Set sourceWb = ThisWorkbook
    On Error Resume Next
    Set sourceData = sourceWb.Worksheets("data")
    On Error GoTo Failed

    ' If the macro host has no data sheet, let the user choose the workbook
    ' that already contains the reference sheet. No blank workbook is created
    ' until this step succeeds or the user explicitly proceeds without data.
    If sourceData Is Nothing Then
        dataPath = Application.GetOpenFilename( _
            FileFilter:="Книги Excel (*.xlsx;*.xlsm;*.xlsb),*.xlsx;*.xlsm;*.xlsb", _
            Title:="Выберите Excel-файл со справочным листом data (Отмена — выход)")
        If VarType(dataPath) = vbBoolean And dataPath = False Then
            MsgBox "Запуск отменён. Книга результата не создавалась.", vbInformation, "Автоматизация перечня"
            Exit Sub
        End If

        Set dataWb = Workbooks.Open(Filename:=CStr(dataPath), ReadOnly:=True, UpdateLinks:=False)
        On Error Resume Next
        Set sourceData = dataWb.Worksheets("data")
        On Error GoTo Failed

        If sourceData Is Nothing Then
            dataWb.Close SaveChanges:=False
            Set dataWb = Nothing
            MsgBox "В выбранной книге не найден лист с точным именем 'data'." & vbCrLf & _
                   "Переименуйте справочный лист в data и запустите макрос снова. Книга результата не создавалась.", _
                   vbExclamation, "Нет справочника data"
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Set resultWb = Workbooks.Add(xlWBATWorksheet)
    Set gRegisterWorkbook = resultWb

    sourceData.Copy Before:=resultWb.Worksheets(1)
    Application.DisplayAlerts = False
    resultWb.Worksheets(resultWb.Worksheets.Count).Delete
    Application.DisplayAlerts = oldAlerts

    If Not dataWb Is Nothing Then
        dataWb.Close SaveChanges:=False
        Set dataWb = Nothing
    End If

    resultWb.Activate
    Application.ScreenUpdating = oldScreen

    MsgBox "Создана новая книга для перечня. Справочник data скопирован." & vbCrLf & _
           "Книгу с макросами оставьте открытой, пока используете кнопки на листе результата.", _
           vbInformation, "Автоматизация перечня"

    CreateMultiFolderApplicationsListCorrected
    Exit Sub

Failed:
    On Error Resume Next
    If Not dataWb Is Nothing Then dataWb.Close SaveChanges:=False
    If Not resultWb Is Nothing Then
        If resultWb.Worksheets.Count = 1 And Len(CStr(resultWb.Worksheets(1).Range("A1").Value)) = 0 Then
            Application.DisplayAlerts = False
            resultWb.Close SaveChanges:=False
        End If
    End If
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
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

    ' Do not silently fall back to the macro-host workbook: callers must handle Nothing.
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
