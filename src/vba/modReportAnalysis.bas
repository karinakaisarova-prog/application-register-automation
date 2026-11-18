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
    Dim headingText As String
    Dim paraCount As Long
    Dim searchFrom As Long

    Dim registerWb As Workbook

    On Error Resume Next
    Set registerWb = GetRegisterWorkbook()
    If Err.Number <> 0 Or registerWb Is Nothing Then
        Err.Clear
        On Error GoTo 0
        MsgBox "Откройте книгу с листом «data» и повторите запуск макроса.", vbExclamation, "Книга результата не найдена"
        Exit Sub
    End If
    On Error GoTo 0

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
        .Filters.Add "Документы Word", "*.doc;*.docx;*.docm"
        If .Show <> -1 Then GoTo CleanExit
        reportPath = .SelectedItems(1)
    End With

    Set wdApp = OpenHiddenWordForReading()
    Set wdDoc = OpenDocumentReadOnly(wdApp, reportPath)

    ' Сохраняет предыдущий результат, пока новая проверка не завершится успешно.
    On Error Resume Next
    Set ws = GetRegisterWorkbook().Worksheets(RESULT_SHEET)
    On Error GoTo Failed
    originalSheetExisted = Not (ws Is Nothing)
    If originalSheetExisted Then
        ws.Copy After:=ws
        ' Копию берём по позиции: у скрытого листа копия не становится активной,
        ' и ActiveSheet указал бы на чужой лист (например, перечень).
        Set backupWs = ws.Parent.Sheets(ws.Index + 1)
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

    ws.Columns("B").NumberFormat = "@"
    ws.Columns("D:F").NumberFormat = "@"
    headingText = "Основной текст (до первого заголовка)"

    For Each para In wdDoc.Paragraphs
        paraCount = paraCount + 1
        If (paraCount Mod 50) = 0 Then
            DoEvents
            Application.StatusBar = "Проверка отчёта: абзац " & paraCount & " | упоминаний: " & (rowOut - FIRST_RESULT_ROW)
        End If
        paraText = Replace(CStr(para.Range.Text), vbCr, "")
        paraText = Replace(paraText, Chr$(7), "")
        ' неразрывный пробел («№ 5») - как обычный; длина строки не меняется
        paraText = Replace(paraText, ChrW(160), " ")

        ' Уровень структуры - ближайший заголовок выше упоминания
        On Error Resume Next
        Err.Clear
        outlineLevel = para.OutlineLevel
        If Err.Number = 18 Then
            On Error GoTo Failed
            Err.Raise 18
        ElseIf Err.Number <> 0 Then
            Err.Clear
            outlineLevel = 10
        End If
        On Error GoTo Failed
        If IsNumeric(outlineLevel) Then
            If CLng(outlineLevel) >= 1 And CLng(outlineLevel) <= 9 Then
                headingText = "Уровень " & CLng(outlineLevel) & ": " & _
                              Trim$(GetListNumber(para) & " " & Left$(Trim$(paraText), 120))
            End If
        End If

        If rx.Test(paraText) Then
            Set matches = rx.Execute(paraText)

            searchFrom = -1
            For Each mt In matches
                ' Определяет страницу по точному диапазону совпадения, а не по концу абзаца:
                ' в одном длинном абзаце могут быть упоминания на разных страницах.
                ' Позиция по тексту абзаца может сдвигаться из-за полей и скрытого
                ' текста - ищем само упоминание внутри абзаца через Find.
                pageNo = FindMentionPage(wdDoc, para, mt, searchFrom)
                ws.Cells(rowOut, 1).Value = rowOut - FIRST_RESULT_ROW + 1
                ws.Cells(rowOut, 2).Value = mt.Value
                ws.Cells(rowOut, 3).Value = pageNo
                ws.Cells(rowOut, 4).Value = headingText
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
    CloseHiddenWordReader wdApp, wdDoc

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
                    recoveryNote = "Не удалось вернуть резервной копии исходное имя. Резервная копия сохранена: " & backupWs.Name
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
    CloseHiddenWordReader wdApp, wdDoc
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
    Dim wbOpenedHere As Boolean
    Dim fso As Object

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
    Set fso = CreateObject("Scripting.FileSystemObject")

    For r = 3 To lastRow
        DoEvents
        filePath = GetRegisterPath(ws.Cells(r, 7))
        If IsWebAddress(filePath) Then
            ws.Cells(r, 10).Value = "SharePoint: используйте синхронизированную папку"
            failed = failed + 1
            GoTo NextRow
        End If
        If Len(filePath) > 0 And FileExistsSafe(fso, filePath) Then
            ext = GetFileExtensionFromPath(filePath)
            pages = Empty
            On Error GoTo FileFailed
            Select Case ext
                Case "doc", "docx", "docm"
                    If wdApp Is Nothing Then Set wdApp = OpenHiddenWordForReading()
                    Set wdDoc = OpenDocumentReadOnly(wdApp, filePath)
                    pages = wdDoc.ComputeStatistics(2) ' wdStatisticPages
                    wdDoc.Close 0
                    Set wdDoc = Nothing
                Case "xls", "xlsx", "xlsm", "xlsb"
                    ' Книгу, которая уже открыта у пользователя, не закрываем
                    Set wb = FindOpenWorkbook(filePath)
                    wbOpenedHere = False
                    If wb Is Nothing Then
                        Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
                        wbOpenedHere = True
                    End If
                    pages = CountWorkbookPrintPages(wb)
                    If wbOpenedHere Then wb.Close False
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
            ' Не оставляет количество страниц от предыдущего запуска, если
            ' в строке перечня нет действующей ссылки или пути к файлу.
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
    If Not wdDoc Is Nothing Then wdDoc.Close 0
    Set wdDoc = Nothing
    If Not wb Is Nothing And wbOpenedHere Then wb.Close False
    Set wb = Nothing
    Err.Clear
    On Error GoTo Failed
    Resume NextRow

Failed:
    errText = Err.Description
    If Err.Number = 18 Then errText = "Операция отменена пользователем (Ctrl+Break)."
    On Error Resume Next
    If pageSnapshotTaken Then ws.Range("J3:J" & lastRow).Formula = oldPageValues
    CloseHiddenWordReader wdApp, wdDoc
    If Not wb Is Nothing And wbOpenedHere Then wb.Close False
    Set wb = Nothing
    Application.StatusBar = False
    MsgBox "Подсчёт страниц прерван: " & errText, vbExclamation
CleanExit:
    On Error Resume Next
    CloseHiddenWordReader wdApp, wdDoc
    If Not wb Is Nothing And wbOpenedHere Then wb.Close False
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
    Dim oldPageBreaks As Boolean

    For Each sh In wb.Worksheets
        DoEvents
        If sh.Visible <> xlSheetVisible Then GoTo NextSheet   ' скрытые листы не печатаются
        On Error Resume Next
        Err.Clear
        oldPageBreaks = sh.DisplayPageBreaks
        sh.DisplayPageBreaks = True
        h = sh.HPageBreaks.Count + 1
        v = sh.VPageBreaks.Count + 1
        pageBreakError = Err.Number
        pageBreakDescription = Err.Description
        sh.DisplayPageBreaks = oldPageBreaks   ' вид книги (в том числе уже открытой у пользователя) не меняем
        On Error GoTo 0

        ' Не указывает условную одну страницу, если Excel не может определить
        ' разрывы страниц (например, когда не настроен доступный принтер).
        If pageBreakError = 18 Then
            Err.Raise pageBreakError, "CountWorkbookPrintPages", pageBreakDescription
        ElseIf pageBreakError <> 0 Then
            Err.Raise vbObjectError + 2501, "CountWorkbookPrintPages", _
                      "Excel не смог рассчитать разрывы страниц листа '" & sh.Name & "': " & pageBreakDescription
        End If

        If Application.WorksheetFunction.CountA(sh.UsedRange) > 0 Then total = total + (h * v)
NextSheet:
    Next sh
    CountWorkbookPrintPages = total
End Function

Private Function GetRegisterPath(ByVal c As Range) As String
    On Error Resume Next
    If c.Hyperlinks.Count > 0 Then
        GetRegisterPath = ResolveLinkPath(CStr(c.Hyperlinks(1).Address), c.Worksheet.Parent)
        If Len(GetRegisterPath) = 0 Then GetRegisterPath = CStr(c.Hyperlinks(1).SubAddress)
    Else
        GetRegisterPath = CStr(c.Value)
    End If
    On Error GoTo 0
    GetRegisterPath = NormalizeReportPath(GetRegisterPath)
End Function

Private Function NormalizeReportPath(ByVal value As String) As String
    value = Trim$(value)
    If IsWebAddress(value) Then
        NormalizeReportPath = value
        Exit Function
    End If
    If LCase$(Left$(value, 8)) = "file:///" Then
        value = Mid$(value, 9)
    ElseIf LCase$(Left$(value, 7)) = "file://" Then
        value = "\\" & Mid$(value, 8)
    End If
    value = Replace(value, "%20", " ")
    value = Replace(value, "/", "\")
    NormalizeReportPath = value
End Function

Private Function IsWebAddress(ByVal value As String) As Boolean
    value = LCase$(Trim$(value))
    IsWebAddress = (Left$(value, 7) = "http://" Or Left$(value, 8) = "https://")
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

' Невидимый Word для чтения: без диалогов и без макросов документа
' (любое окно в невидимом Word заставило бы Excel ждать бесконечно).
Private Function OpenHiddenWordForReading() As Object
    Dim wdApp As Object
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    wdApp.DisplayAlerts = 0                 ' wdAlertsNone
    On Error Resume Next
    wdApp.AutomationSecurity = 3            ' msoAutomationSecurityForceDisable
    On Error GoTo 0
    Set OpenHiddenWordForReading = wdApp
End Function

' Открывает документ только для чтения. Документ с паролем не спрашивает
' пароль, а даёт ошибку - она попадает в отчёт как «Ошибка чтения».
Private Function OpenDocumentReadOnly(ByVal wdApp As Object, ByVal filePath As String) As Object
    Set OpenDocumentReadOnly = wdApp.Documents.Open(FileName:=filePath, ConfirmConversions:=False, _
        ReadOnly:=True, AddToRecentFiles:=False, PasswordDocument:="#", Visible:=False)
End Function

Private Sub CloseHiddenWordReader(ByRef wdApp As Object, ByRef wdDoc As Object)
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close 0
    Set wdDoc = Nothing
    If Not wdApp Is Nothing Then
        wdApp.NormalTemplate.Saved = True
        wdApp.Quit 0
    End If
    Set wdApp = Nothing
End Sub

' Номер страницы упоминания: Find внутри абзаца по тексту совпадения.
' searchFrom - позиция, с которой искать следующее совпадение этого же абзаца
' (-1 - с начала абзаца).
Private Function FindMentionPage(ByVal wdDoc As Object, ByVal para As Object, ByVal mt As Object, _
                                 ByRef searchFrom As Long) As Long
    Dim rng As Object
    Dim startPos As Long

    startPos = para.Range.Start
    If searchFrom > startPos Then startPos = searchFrom
    Set rng = wdDoc.Range(startPos, para.Range.End)
    With rng.Find
        .ClearFormatting
        .Text = CStr(mt.Value)
        .MatchCase = False
        .MatchWildcards = False
        .Forward = True
        .Wrap = 0                           ' wdFindStop
    End With
    If rng.Find.Execute Then
        FindMentionPage = CLng(rng.Information(1))
        searchFrom = rng.End
    Else
        ' запасной путь - позиция по тексту абзаца
        Set rng = wdDoc.Range(para.Range.Start + CLng(mt.FirstIndex), _
                              para.Range.Start + CLng(mt.FirstIndex) + CLng(mt.Length))
        FindMentionPage = CLng(rng.Information(1))
    End If
End Function

' Номер заголовка из автонумерации («2.1»), если есть
Private Function GetListNumber(ByVal para As Object) As String
    On Error Resume Next
    GetListNumber = CStr(para.Range.ListFormat.ListString)
    If Err.Number = 18 Then
        On Error GoTo 0
        Err.Raise 18
    End If
    Err.Clear
End Function

Private Function FindOpenWorkbook(ByVal filePath As String) As Workbook
    Dim w As Workbook
    For Each w In Application.Workbooks
        If StrComp(w.FullName, filePath, vbTextCompare) = 0 Then
            Set FindOpenWorkbook = w
            Exit Function
        End If
    Next w
End Function

' Проверка существования файла без ошибки на недопустимом или длинном пути
' и с любыми буквами в имени (Dir$ понимает только кодировку Windows).
Private Function FileExistsSafe(ByVal fso As Object, ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExistsSafe = fso.FileExists(filePath)
    If Err.Number = 18 Then
        On Error GoTo 0
        Err.Raise 18
    End If
    Err.Clear
End Function
