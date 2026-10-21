Attribute VB_Name = "modReportAnalysis"
Option Explicit

' Word report analysis. Uses late binding; Microsoft Word must be installed.
' Results are written to worksheet "Упоминания отчёта".
' This module does not modify the source report.

Private Const RESULT_SHEET As String = "Упоминания отчёта"
Private Const FIRST_RESULT_ROW As Long = 2

Public Sub ScanWordReportMentions()
    Dim fd As FileDialog
    Dim reportPath As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim ws As Worksheet
    Dim para As Object
    Dim rng As Object
    Dim rx As Object
    Dim matches As Object
    Dim mt As Object
    Dim rowOut As Long
    Dim pageNo As Long
    Dim outlineLevel As Variant
    Dim paraText As String
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim errText As String
    Dim cancelled As Boolean

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Выберите основной Word-отчёт"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Документы Word", "*.doc;*.docx"
        If .Show <> -1 Then GoTo CleanExit
        reportPath = .SelectedItems(1)
    End With

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(reportPath, False, True, False)
    Set ws = GetOrCreateReportSheet(RESULT_SHEET)
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "№"
    ws.Cells(1, 2).Value = "Упоминание"
    ws.Cells(1, 3).Value = "Страница"
    ws.Cells(1, 4).Value = "Уровень структуры"
    ws.Cells(1, 5).Value = "Текст абзаца"
    ws.Cells(1, 6).Value = "Файл отчёта"
    rowOut = FIRST_RESULT_ROW

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "приложени[ея]\s*(№|номер)?\s*([0-9]+[А-ЯA-Z]?)"

    For Each para In wdDoc.Paragraphs
        DoEvents
        paraText = Replace(CStr(para.Range.Text), vbCr, "")
        paraText = Replace(paraText, Chr$(7), "")
        If rx.Test(paraText) Then
            Set matches = rx.Execute(paraText)
            pageNo = 0
            On Error Resume Next
            pageNo = para.Range.Information(3) ' wdActiveEndPageNumber
            outlineLevel = para.OutlineLevel
            If Err.Number <> 0 Then
                Err.Clear
                outlineLevel = ""
            End If
            On Error GoTo Failed

            For Each mt In matches
                ws.Cells(rowOut, 1).Value = rowOut - FIRST_RESULT_ROW + 1
                ws.Cells(rowOut, 2).Value = mt.Value
                ws.Cells(rowOut, 3).Value = pageNo
                If IsNumeric(outlineLevel) Then
                    If CLng(outlineLevel) >= 1 And CLng(outlineLevel) <= 9 Then
                        ws.Cells(rowOut, 4).Value = CLng(outlineLevel)
                    Else
                        ws.Cells(rowOut, 4).Value = "Основной текст"
                    End If
                Else
                    ws.Cells(rowOut, 4).Value = outlineLevel
                End If
                ws.Cells(rowOut, 5).Value = Trim$(paraText)
                ws.Cells(rowOut, 6).Value = reportPath
                rowOut = rowOut + 1
            Next mt
        End If
    Next para

    FormatReportResults ws, rowOut - 1
    MsgBox "Проверка завершена. Найдено упоминаний: " & (rowOut - FIRST_RESULT_ROW) & _
           vbCrLf & "Результат записан на лист '" & RESULT_SHEET & "'.", vbInformation, "Проверка Word-отчёта"
    GoTo CleanExit

Failed:
    cancelled = (Err.Number = 18)
    errText = Err.Description
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    On Error GoTo 0
    MsgBox IIf(cancelled, "Проверка отменена.", "Ошибка при анализе Word-отчёта: " & errText), vbExclamation
    GoTo RestoreExcel

CleanExit:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
RestoreExcel:
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
End Sub

Public Sub CountPagesForRegister()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim filePath As String
    Dim ext As String
    Dim pages As Variant
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim wb As Workbook
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim errText As String
    Dim completed As Long
    Dim failed As Long

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "На активном листе нет строк перечня.", vbInformation
        GoTo CleanExit
    End If

    For r = 3 To lastRow
        DoEvents
        filePath = GetRegisterPath(ws.Cells(r, 7))
        If Len(filePath) > 0 And Len(Dir$(filePath)) > 0 Then
            ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))
            pages = Empty
            On Error GoTo FileFailed
            Select Case ext
                Case "doc", "docx"
                    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
                    wdApp.Visible = False
                    Set wdDoc = wdApp.Documents.Open(filePath, False, True, False)
                    pages = wdDoc.ComputeStatistics(2) ' wdStatisticPages
                    wdDoc.Close False
                    Set wdDoc = Nothing
                Case "xls", "xlsx", "xlsm", "xlsb"
                    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
                    pages = CountWorkbookPrintPages(wb)
                    wb.Close False
                    Set wb = Nothing
                Case "pdf"
                    pages = CountPdfPagesWithAcrobat(filePath)
                    If IsNumeric(pages) Then
                        ws.Cells(r, 10).Value = pages
                        completed = completed + 1
                    Else
                        ws.Cells(r, 10).Value = "Требуется Acrobat Pro/Standard"
                        failed = failed + 1
                    End If
                    GoTo NextRow
                Case Else
                    ws.Cells(r, 10).Value = "Неподдерживаемый формат"
                    GoTo NextRow
            End Select
            ws.Cells(r, 10).Value = pages
            completed = completed + 1
        ElseIf Len(filePath) > 0 Then
            ws.Cells(r, 10).Value = "Файл не найден"
            failed = failed + 1
        End If
NextRow:
        Application.StatusBar = "Подсчёт страниц: строка " & r & " из " & lastRow
        On Error GoTo Failed
    Next r

    MsgBox "Подсчёт завершён." & vbCrLf & "Обработано: " & completed & _
           vbCrLf & "Проблемы: " & failed & _
           vbCrLf & "Для подсчёта PDF требуется установленный Adobe Acrobat Pro/Standard с COM-интерфейсом; Reader может не поддерживать этот вызов.", _
           vbInformation, "Подсчёт страниц"
    GoTo CleanExit

FileFailed:
    If Err.Number = 18 Then
        ' Ctrl+Break must reach the top-level handler so Excel state is restored.
        GoTo Failed
    End If
    failed = failed + 1
    ws.Cells(r, 10).Value = "Ошибка чтения"
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    Set wdDoc = Nothing
    If Not wb Is Nothing Then wb.Close False
    Set wb = Nothing
    Err.Clear
    On Error GoTo Failed
    Resume NextRow

Failed:
    errText = Err.Description
    If Err.Number = 18 Then errText = "Операция отменена пользователем (Ctrl+Break)."
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    If Not wb Is Nothing Then wb.Close False
    Application.StatusBar = False
    MsgBox "Подсчёт страниц прерван: " & errText, vbExclamation
CleanExit:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    If Not wb Is Nothing Then wb.Close False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
End Sub

Private Function CountPdfPagesWithAcrobat(ByVal pdfPath As String) As Variant
    Dim pdfDoc As Object
    On Error GoTo PdfUnavailable
    Set pdfDoc = CreateObject("AcroExch.PDDoc")
    If pdfDoc.Open(pdfPath) Then
        CountPdfPagesWithAcrobat = CLng(pdfDoc.GetNumPages)
        pdfDoc.Close
    Else
        CountPdfPagesWithAcrobat = Empty
    End If
    Set pdfDoc = Nothing
    Exit Function
PdfUnavailable:
    On Error Resume Next
    If Not pdfDoc Is Nothing Then pdfDoc.Close
    Set pdfDoc = Nothing
    CountPdfPagesWithAcrobat = Empty
End Function

Private Function CountWorkbookPrintPages(ByVal wb As Workbook) As Long
    Dim sh As Worksheet
    Dim total As Long
    Dim h As Long
    Dim v As Long
    For Each sh In wb.Worksheets
        On Error Resume Next
        sh.DisplayPageBreaks = True
        h = sh.HPageBreaks.Count + 1
        v = sh.VPageBreaks.Count + 1
        If Err.Number <> 0 Then
            Err.Clear
            h = 1
            v = 1
        End If
        On Error GoTo 0
        If Application.WorksheetFunction.CountA(sh.UsedRange) > 0 Then total = total + (h * v)
    Next sh
    CountWorkbookPrintPages = total
End Function

Private Function GetRegisterPath(ByVal c As Range) As String
    On Error Resume Next
    If c.Hyperlinks.Count > 0 Then
        GetRegisterPath = CStr(c.Hyperlinks(1).Address)
    Else
        GetRegisterPath = CStr(c.Value)
    End If
    On Error GoTo 0
End Function

Private Function GetOrCreateReportSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateReportSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateReportSheet Is Nothing Then
        Set GetOrCreateReportSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateReportSheet.Name = sheetName
    End If
End Function

Private Sub FormatReportResults(ByVal ws As Worksheet, ByVal lastRow As Long)
    With ws.Range("A1:F" & Application.Max(1, lastRow))
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlTop
    End With
    With ws.Range("A1:F1")
        .Interior.Color = RGB(198, 224, 180)
        .Font.Bold = True
        .WrapText = True
    End With
    ws.Columns("A:D").AutoFit
    ws.Columns("E").ColumnWidth = 65
    ws.Columns("F").ColumnWidth = 45
    ws.Columns("E:F").WrapText = True
End Sub
