Attribute VB_Name = "modReportAnalysis"
Option Explicit

' Анализирует упоминания приложений в отчёте Word. Для выполнения требуется Microsoft Word.
' Результаты записываются на лист «Упоминания отчёта».
' Исходный отчёт не изменяется.

Private Const RESULT_SHEET As String = "Упоминания отчёта"
Private Const FIRST_RESULT_ROW As Long = 2

Public Sub ScanWordReportMentions()
   Dim fd As FileDialog
    Dim reportPath As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim ws As Worksheet
    Dim backupWs As Worksheet
    Dim damagedWs As Worksheet
    Dim recoverySheetName As String
    Dim backupName As String
    Dim originalSheetExisted As Boolean
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
    Dim recoveryNote As String
    Dim cancelled As Boolean
    Dim oldStatusBar As Variant
    Dim oldCancelKey As XlEnableCancelKey

    If GetRegisterWorkbook() Is Nothing Then
        MsgBox "Сначала запустите StartRegisterAutomation, чтобы создать книгу результата.", vbExclamation, "Книга результата не найдена"
        Exit Sub
    End If

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    oldCancelKey = Application.EnableCancelKey
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

    ' Сохраняет предыдущий результат, пока новая проверка не завершится успешно.
    On Error Resume Next
    Set ws = GetRegisterWorkbook().Worksheets(RESULT_SHEET)
    On Error GoTo Failed
    originalSheetExisted = Not (ws Is Nothing)
    If originalSheetExisted Then
        ws.Copy After:=ws
        Set backupWs = ActiveSheet
        backupName = "tmpMentions_" & Format$(Now, "yymmdd_hhnnss") & "_" & CStr(GetRegisterWorkbook().Worksheets.Count)
        backupWs.Name = backupName
        Set ws = GetRegisterWorkbook().Worksheets(RESULT_SHEET)
    Else
        Set ws = GetOrCreateReportSheet(RESULT_SHEET)
    End If
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
    rx.Pattern = "(приложение|приложения|приложению|приложением|приложении|приложений|прил)\.?\s*(№|номер)?\s*([0-9]+[А-ЯA-Z]?)"

    For Each para In wdDoc.Paragraphs
        DoEvents
        paraText = Replace(CStr(para.Range.Text), vbCr, "")
        paraText = Replace(paraText, Chr$(7), "")
        If rx.Test(paraText) Then
            Set matches = rx.Execute(paraText)
            On Error Resume Next
            Err.Clear
            outlineLevel = para.OutlineLevel
            If Err.Number = 18 Then
                On Error GoTo 0
                GoTo Failed
            ElseIf Err.Number <> 0 Then
                Err.Clear
                outlineLevel = ""
            End If
            On Error GoTo Failed

            For Each mt In matches
                ' Resolve the page from the exact match range, not the paragraph
                ' end: one long paragraph can contain mentions on different pages.
                Set rng = wdDoc.Range(para.Range.Start + CLng(mt.FirstIndex), _
                                      para.Range.Start + CLng(mt.FirstIndex) + CLng(mt.Length))
                pageNo = CLng(rng.Information(1)) ' wdActiveEndAdjustedPageNumber: page number as displayed in Word
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
    If Not backupWs Is Nothing Then
        Application.DisplayAlerts = False
        backupWs.Delete
        Set backupWs = Nothing
        Application.DisplayAlerts = oldAlerts
    End If
    MsgBox "Проверка завершена. Найдено упоминаний: " & (rowOut - FIRST_RESULT_ROW) & _
           vbCrLf & "Результат записан на лист '" & RESULT_SHEET & "'.", vbInformation, "Проверка Word-отчёта"
    GoTo CleanExit

Failed:
    cancelled = (Err.Number = 18)
    errText = Err.Description
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit

    ' Восстанавливает прежний лист результата, чтобы после ошибки или отмены не осталось неполных данных.
    If Not backupWs Is Nothing Then
        Application.DisplayAlerts = False
        Set damagedWs = ws
        If damagedWs Is Nothing Then
            Err.Clear
            backupWs.Name = RESULT_SHEET
            If Err.Number = 0 Then
                Set ws = backupWs
                Set backupWs = Nothing
            Else
                Err.Clear
                recoveryNote = "Не удалось вернуть резервной копии исходное имя. Резервная копия сохранена: " & backupWs.Name
            End If
        Else
            recoverySheetName = GetUniqueMentionRecoveryName()
            Err.Clear
            damagedWs.Name = recoverySheetName
            If Err.Number = 0 Then
                Err.Clear
                backupWs.Name = RESULT_SHEET
                If Err.Number = 0 Then
                    Set ws = backupWs
                    Set backupWs = Nothing
                    Err.Clear
                    damagedWs.Delete
                    If Err.Number <> 0 Then
                        Err.Clear
                        recoveryNote = "Прежние результаты восстановлены; неполный лист сохранён под именем " & recoverySheetName
                    End If
                Else
                    Err.Clear
                    damagedWs.Name = RESULT_SHEET
                    LogRestoreIssue "Не удалось вернуть резервной копии исходное имя. Резервная копия сохранена: " & backupWs.Name
                End If
            Else
                Err.Clear
                recoveryNote = "Не удалось безопасно переименовать неполный лист. Резервная копия сохранена: " & backupWs.Name
            End If
        End If
    ElseIf Not ws Is Nothing And Not originalSheetExisted Then
        Application.DisplayAlerts = False
        ws.Delete
        Set ws = Nothing
    End If
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0
    MsgBox IIf(cancelled, "Проверка отменена. Выполнена попытка сохранить прежние результаты.", _
        "Ошибка при анализе Word-отчёта: " & errText & vbCrLf & _
        "Выполнена попытка восстановления прежних результатов. Проверьте листы книги.") & _
        IIf(Len(recoveryNote) > 0, vbCrLf & recoveryNote, ""), vbExclamation
    GoTo RestoreExcel

CleanExit:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
RestoreExcel:
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
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
    Dim oldPageValues As Variant
    Dim pageSnapshotTaken As Boolean
    Dim oldStatusBar As Variant
    Dim oldCancelKey As XlEnableCancelKey

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    oldCancelKey = Application.EnableCancelKey
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = GetActiveRegisterSheet()
    If ws Is Nothing Then
        MsgBox "Не найден лист перечня в книге результата. Сначала сформируйте перечень.", vbExclamation
        GoTo CleanExit
    End If
    lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row
    If lastRow < 3 Then
        MsgBox "На активном листе нет строк перечня.", vbInformation
        GoTo CleanExit
    End If

    ' Сохраняет исходный столбец с количеством страниц, чтобы восстановить его при прерывании или критической ошибке.
    oldPageValues = ws.Range("J3:J" & lastRow).Formula
    pageSnapshotTaken = True

    For r = 3 To lastRow
        DoEvents
        filePath = GetRegisterPath(ws.Cells(r, 7))
        If Len(filePath) > 0 And Len(Dir$(filePath)) > 0 Then
            ext = GetFileExtensionFromPath(filePath)
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
                    failed = failed + 1
                    GoTo NextRow
            End Select
            ws.Cells(r, 10).Value = pages
            completed = completed + 1
        ElseIf Len(filePath) > 0 Then
            ws.Cells(r, 10).Value = "Файл не найден"
            failed = failed + 1
        Else
            ' Do not leave a stale page count from an earlier run when the
            ' register row has no usable hyperlink/path.
            ws.Cells(r, 10).Value = "Нет ссылки на файл"
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
        ' Прерывание Ctrl+Break должно дойти до основного обработчика, чтобы восстановить настройки Excel.
        GoTo Failed
    End If
    failed = failed + 1
    If Err.Number = vbObjectError + 2501 Then
        ws.Cells(r, 10).Value = "Не удалось определить страницы Excel: проверьте принтер и параметры печати"
    Else
        ws.Cells(r, 10).Value = "Ошибка чтения"
    End If
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
    If pageSnapshotTaken Then ws.Range("J3:J" & lastRow).Formula = oldPageValues
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
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
End Sub

Private Function GetFileExtensionFromPath(ByVal filePath As String) As String
    Dim fileName As String
    Dim separatorPos As Long
    Dim dotPos As Long

    separatorPos = InStrRev(filePath, "\")
    fileName = Mid$(filePath, separatorPos + 1)
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 And dotPos < Len(fileName) Then
        GetFileExtensionFromPath = LCase$(Mid$(fileName, dotPos + 1))
    Else
        GetFileExtensionFromPath = ""
    End If
End Function

Private Function CountPdfPagesWithAcrobat(ByVal pdfPath As String) As Variant
    Dim pdfDoc As Object
    Dim savedErrNumber As Long
    Dim savedErrDescription As String

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
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error Resume Next
    If Not pdfDoc Is Nothing Then pdfDoc.Close
    Set pdfDoc = Nothing
    On Error GoTo 0

    ' Не подавляет Ctrl+Break: вызывающая процедура должна восстановить сохранённые значения количества страниц.
    If savedErrNumber = 18 Then
        Err.Raise savedErrNumber, "CountPdfPagesWithAcrobat", savedErrDescription
    End If
    CountPdfPagesWithAcrobat = Empty
End Function

Private Function CountWorkbookPrintPages(ByVal wb As Workbook) As Long
    Dim sh As Worksheet
    Dim total As Long
    Dim h As Long
    Dim v As Long
    Dim pageBreakError As Long
    Dim pageBreakDescription As String

    For Each sh In wb.Worksheets
        DoEvents
        On Error Resume Next
        Err.Clear
        sh.DisplayPageBreaks = True
        h = sh.HPageBreaks.Count + 1
        v = sh.VPageBreaks.Count + 1
        pageBreakError = Err.Number
        pageBreakDescription = Err.Description
        On Error GoTo 0

        ' Do not report a fabricated one-page result when Excel cannot calculate
        ' print breaks (commonly because no usable printer is configured).
        If pageBreakError = 18 Then
            Err.Raise pageBreakError, "CountWorkbookPrintPages", pageBreakDescription
        ElseIf pageBreakError <> 0 Then
            Err.Raise vbObjectError + 2501, "CountWorkbookPrintPages", _
                      "Excel не смог рассчитать разрывы страниц листа '" & sh.Name & "': " & pageBreakDescription
        End If

        If Application.WorksheetFunction.CountA(sh.UsedRange) > 0 Then total = total + (h * v)
    Next sh
    CountWorkbookPrintPages = total
End Function

Private Function GetRegisterPath(ByVal c As Range) As String
    On Error Resume Next
    If c.Hyperlinks.Count > 0 Then
        GetRegisterPath = CStr(c.Hyperlinks(1).Address)
        If Len(GetRegisterPath) = 0 Then GetRegisterPath = CStr(c.Hyperlinks(1).SubAddress)
    Else
        GetRegisterPath = CStr(c.Value)
    End If
    On Error GoTo 0
    GetRegisterPath = NormalizeReportPath(GetRegisterPath)
End Function

Private Function NormalizeReportPath(ByVal value As String) As String
    value = Trim$(value)
    If LCase$(Left$(value, 8)) = "file:///" Then value = Mid$(value, 9)
    value = Replace(value, "%20", " ")
    value = Replace(value, "/", "\")
    NormalizeReportPath = value
End Function

Private Function GetUniqueMentionRecoveryName() As String
    Dim candidate As String
    Dim suffix As Long
    candidate = "_failed_mentions"
    suffix = 1
    Do While MentionWorksheetExists(candidate)
        candidate = "_failed_mnt_" & CStr(suffix)
        suffix = suffix + 1
    Loop
    GetUniqueMentionRecoveryName = candidate
End Function

Private Function MentionWorksheetExists(ByVal sheetName As String) As Boolean
    Dim testWs As Worksheet
    On Error Resume Next
    Set testWs = GetRegisterWorkbook().Worksheets(sheetName)
    MentionWorksheetExists = Not testWs Is Nothing
    On Error GoTo 0
End Function

Private Function GetOrCreateReportSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateReportSheet = GetRegisterWorkbook().Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateReportSheet Is Nothing Then
        Set GetOrCreateReportSheet = GetRegisterWorkbook().Worksheets.Add(After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count))
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
