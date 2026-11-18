Attribute VB_Name = "modTableNumbering"
Option Explicit

' Проверяет все таблицы документа Word.
' Перед изменением документа формирует предварительный отчёт на листе Excel.
' Для таблиц без подписи добавляет автоматические поля SEQ Word и обновляет поля документа.
' Подписи, которые нельзя безопасно преобразовать (многоуровневый номер «Таблица 2.3»,
' буквенный «Таблица А.1», подпись с полями), не меняются и помечаются для ручной проверки.
' Для выполнения требуется Microsoft Word.

Private Const TABLE_REPORT_SHEET As String = "Проверка таблиц"
Private Const FIRST_DATA_ROW As Long = 2
Private Const WD_CAPTION_POSITION_ABOVE As Long = 0
Private Const WD_FIELD_SEQUENCE As Long = 12
Private Const WD_WITHIN_TABLE As Long = 12
Private Const WD_ALERTS_NONE As Long = 0
Private Const WD_NO_PROTECTION As Long = -1
Private Const MSO_AUTOMATION_SECURITY_FORCE_DISABLE As Long = 3
Private Const CAPTION_SEARCH_LIMIT As Long = 4

' Формат поля: отображает «Таблица № 1»
Private Const SEQ_FIELD_CODE As String = " SEQ Таблица \# ""'Таблица № '0"" "

' Результат разбора подписи над таблицей
Private Const CAP_HAS_SEQ As Long = 1      ' подпись с полем SEQ «Таблица» - нумерация есть
Private Const CAP_CONVERT As Long = 2      ' текстовая подпись «Таблица 5 …» - будет преобразована
Private Const CAP_MISSING As Long = 3      ' подписи нет - будет добавлена
Private Const CAP_MANUAL As Long = 4       ' подпись есть, но преобразовать безопасно нельзя

Public Sub CheckAndNumberWordTables()
    Dim fd As FileDialog
    Dim reportPath As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim tbl As Object
    Dim ws As Worksheet
    Dim r As Long
    Dim tableCount As Long
    Dim missingCount As Long
    Dim manualCaptionCount As Long
    Dim manualCheckCount As Long
    Dim insertedCount As Long
    Dim pageNo As Long
    Dim captionText As String
    Dim capKind As Long
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
    Dim errText As String
    Dim cancelled As Boolean
    Dim answer As VbMsgBoxResult
    Dim backupPath As String
    Dim changesStarted As Boolean
    Dim fsoRestore As Object
    Dim oldStatusBar As Variant
    Dim oldCancelKey As XlEnableCancelKey
    Dim priorWs As Worksheet
    Dim stageWs As Worksheet
    Dim finalWs As Worksheet
    Dim stageName As String, priorBackupName As String
    Dim finalStageName As String, finalOldName As String
    Dim reportCommitted As Boolean, priorRenamed As Boolean
    Dim finalReportCommitted As Boolean, finalOldRenamed As Boolean
    Dim oldTrack As Boolean, trackChanged As Boolean
    Dim registerWb As Workbook
    Dim unconfirmedCount As Long
    Dim restoredOk As Boolean

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
        .Title = "Выберите основной Word-отчёт для проверки таблиц"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Документы Word", "*.doc;*.docx;*.docm"
        If .Show <> -1 Then GoTo CleanExit
        reportPath = .SelectedItems(1)
    End With

    Application.StatusBar = "Открытие документа Word..."
    Set wdApp = OpenHiddenWord()
    ' Пароль «#» вместо запроса: документ с паролем даёт ошибку, а не окно
    ' в невидимом Word, на котором Excel ждал бы бесконечно.
    Set wdDoc = wdApp.Documents.Open(FileName:=reportPath, ConfirmConversions:=False, _
                                     ReadOnly:=False, AddToRecentFiles:=False, PasswordDocument:="#")

    ' Формирует новый отчёт на временном листе и сохраняет предыдущий отчёт
    ' до успешного завершения проверки.
    On Error Resume Next
    Set priorWs = registerWb.Worksheets(TABLE_REPORT_SHEET)
    On Error GoTo Failed
    stageName = GetUniqueTableAuditSheetName("_tmp_table_audit")
    Set stageWs = registerWb.Worksheets.Add(After:=registerWb.Worksheets(registerWb.Worksheets.Count))
    stageWs.Name = stageName
    Set ws = stageWs
    ws.Columns("C:E").NumberFormat = "@"
    ws.Range("A1:G1").Value = Array("№ таблицы", "Страница", "Подпись перед таблицей", _
                                    "Поле SEQ «Таблица»", "Результат проверки", "Путь к отчёту", "Индекс таблицы")
    r = FIRST_DATA_ROW
    tableCount = wdDoc.Tables.Count

    For Each tbl In wdDoc.Tables
        DoEvents
        pageNo = 0
        capKind = ClassifyTableCaption(wdDoc, tbl.Range.Start, captionText)

        On Error Resume Next
        Err.Clear
        pageNo = tbl.Range.Information(3) ' wdActiveEndPageNumber
        If Err.Number = 18 Then GoTo Failed
        Err.Clear
        On Error GoTo Failed

        ws.Cells(r, 1).Value = r - FIRST_DATA_ROW + 1
        ws.Cells(r, 2).Value = pageNo
        ws.Cells(r, 3).Value = captionText
        ws.Cells(r, 4).Value = IIf(capKind = CAP_HAS_SEQ, "Да", "Нет")
        Select Case capKind
            Case CAP_HAS_SEQ
                ws.Cells(r, 5).Value = "Нумерация найдена"
            Case CAP_CONVERT
                ws.Cells(r, 5).Value = "Текстовая подпись без SEQ — будет преобразована"
                manualCaptionCount = manualCaptionCount + 1
            Case CAP_MANUAL
                ws.Cells(r, 5).Value = "Требуется ручная проверка подписи (не меняется)"
                manualCheckCount = manualCheckCount + 1
            Case Else
                ws.Cells(r, 5).Value = "Требуется добавить подпись"
                missingCount = missingCount + 1
        End Select
        ws.Cells(r, 6).Value = reportPath
        ws.Cells(r, 7).Value = r - FIRST_DATA_ROW + 1
        r = r + 1
        If (r Mod 20) = 0 Then Application.StatusBar = "Проверка таблиц: " & (r - FIRST_DATA_ROW) & " из " & tableCount
    Next tbl

    FormatTableReport ws, r - 1

    ' Заменяет предыдущий отчёт только после завершения проверки и форматирования.
    If Not priorWs Is Nothing Then
        priorBackupName = GetUniqueTableAuditSheetName("_bak_table_audit")
        priorWs.Name = priorBackupName
        priorRenamed = True
    End If
    On Error GoTo CommitFailed
    ws.Name = TABLE_REPORT_SHEET
    reportCommitted = True
    Set stageWs = Nothing
    If priorRenamed Then
        Application.DisplayAlerts = False
        priorWs.Delete
        Set priorWs = Nothing
        priorRenamed = False
    End If
    Application.DisplayAlerts = oldAlerts
    On Error GoTo Failed
    ActivateReportSheet registerWb, ws

    If tableCount = 0 Then
        MsgBox "В документе не найдено таблиц через коллекцию Word Tables.", vbInformation, "Проверка таблиц"
        GoTo CleanExit
    End If

    If missingCount = 0 And manualCaptionCount = 0 Then
        MsgBox "Проверка завершена. Найдено таблиц: " & tableCount & _
               vbCrLf & IIf(manualCheckCount > 0, _
                    "Подписей для ручной проверки: " & manualCheckCount & " (макрос их не меняет)." & vbCrLf, _
                    "У всех таблиц найдены подписи с полем SEQ «Таблица»." & vbCrLf) & _
               "Отчёт проверки записан на лист '" & TABLE_REPORT_SHEET & "'.", _
               vbInformation, "Проверка таблиц"
        GoTo CleanExit
    End If

    ' Изменения возможны только в документе, который можно сохранить.
    If wdDoc.ReadOnly Then
        MsgBox "Документ открыт только для чтения (возможно, он открыт в Word у вас или у коллеги)." & vbCrLf & _
               "Предварительный отчёт сохранён на листе '" & TABLE_REPORT_SHEET & "'. Закройте документ и повторите.", _
               vbExclamation, "Проверка таблиц"
        GoTo CleanExit
    End If
    If wdDoc.ProtectionType <> WD_NO_PROTECTION Then
        MsgBox "Документ защищён от изменений. Снимите защиту и повторите." & vbCrLf & _
               "Предварительный отчёт сохранён на листе '" & TABLE_REPORT_SHEET & "'.", _
               vbExclamation, "Проверка таблиц"
        GoTo CleanExit
    End If

    answer = MsgBox("Предварительная проверка завершена." & vbCrLf & _
                    "Всего таблиц: " & tableCount & vbCrLf & _
                    "Таблиц без подписи: " & missingCount & vbCrLf & _
                    "Текстовых подписей без SEQ (будут преобразованы): " & manualCaptionCount & vbCrLf & _
                    "Подписей для ручной проверки (не меняются): " & manualCheckCount & vbCrLf & vbCrLf & _
                    "Добавить подписи отсутствующим таблицам и преобразовать существующие текстовые подписи в поля SEQ?" & _
                    vbCrLf & "Текст после номера будет сохранён. Ссылки в тексте («см. табл. 3») не обновляются." & _
                    IIf(wdDoc.Revisions.Count > 0, vbCrLf & "В документе есть непринятые исправления: " & _
                        wdDoc.Revisions.Count & ". Новые правки будут внесены без режима исправлений.", "") & _
                    vbCrLf & "Перед изменением будет создана резервная копия.", _
                    vbYesNo + vbQuestion + vbDefaultButton2, "Подтверждение нумерации")
    If answer <> vbYes Then
        MsgBox "Изменения не внесены. Предварительный отчёт сохранён на листе '" & TABLE_REPORT_SHEET & "'.", _
               vbInformation, "Проверка таблиц"
        GoTo CleanExit
    End If

    ' Создаёт резервную копию в той же папке только после подтверждения изменений пользователем.
    backupPath = CreateWordBackupPath(reportPath)
    FileCopy reportPath, backupPath
    LogTableBackup ws, backupPath
    changesStarted = True

    ' Режим исправлений выключается на время правок: иначе удалённый текст
    ' остаётся в документе как «исправление» и ломает расчёт позиций.
    oldTrack = wdDoc.TrackRevisions
    wdDoc.TrackRevisions = False
    trackChanged = True

    EnsureTableCaptionLabel wdApp

    ' Обрабатывает таблицы с конца документа, чтобы вставка подписей не смещала расположение ещё не проверенных таблиц.
    For r = tableCount To 1 Step -1
        DoEvents
        Set tbl = wdDoc.Tables(r)
        capKind = ClassifyTableCaption(wdDoc, tbl.Range.Start, captionText)
        Select Case capKind
            Case CAP_CONVERT
                ConvertManualTableCaptionToSequence wdDoc, tbl
                insertedCount = insertedCount + 1
            Case CAP_MISSING
                InsertCustomTableCaption wdApp, wdDoc, tbl
                insertedCount = insertedCount + 1
        End Select
        Application.StatusBar = "Обработка подписей: " & insertedCount & " из " & (missingCount + manualCaptionCount)
    Next r

    wdDoc.Fields.Update
    UpdateAllStoryFields wdDoc
    wdDoc.TrackRevisions = oldTrack
    trackChanged = False
    wdDoc.Save
    ' Документ сохранён успешно: дальнейшие ошибки при построении отчёта в Excel
    ' не должны откатывать документ к резервной копии.
    changesStarted = False

    ' Формирует окончательный отчёт на отдельном временном листе. Текущий
    ' предварительный отчёт сохраняется, пока новый не будет полностью заполнен.
    finalStageName = GetUniqueTableAuditSheetName("_tmp_table_final")
    Set finalWs = registerWb.Worksheets.Add(After:=registerWb.Worksheets(registerWb.Worksheets.Count))
    finalWs.Name = finalStageName
    finalWs.Columns("C:E").NumberFormat = "@"
    finalWs.Range("A1:G1").Value = Array("№ таблицы", "Страница", "Подпись перед таблицей", _
                                         "Поле SEQ «Таблица»", "Результат проверки", "Путь к отчёту", "Индекс таблицы")
    r = FIRST_DATA_ROW
    For Each tbl In wdDoc.Tables
        DoEvents
        pageNo = 0
        capKind = ClassifyTableCaption(wdDoc, tbl.Range.Start, captionText)
        On Error Resume Next
        Err.Clear
        pageNo = tbl.Range.Information(3)
        If Err.Number = 18 Then GoTo Failed
        Err.Clear
        On Error GoTo Failed

        finalWs.Cells(r, 1).Value = r - FIRST_DATA_ROW + 1
        finalWs.Cells(r, 2).Value = pageNo
        finalWs.Cells(r, 3).Value = captionText
        finalWs.Cells(r, 4).Value = IIf(capKind = CAP_HAS_SEQ, "Да", "Нет")
        Select Case capKind
            Case CAP_HAS_SEQ
                finalWs.Cells(r, 5).Value = "Нумерация найдена"
            Case CAP_MANUAL
                finalWs.Cells(r, 5).Value = "Требуется ручная проверка подписи (не менялась)"
            Case Else
                finalWs.Cells(r, 5).Value = "ОШИБКА: после обработки поле SEQ не найдено"
                unconfirmedCount = unconfirmedCount + 1
        End Select
        finalWs.Cells(r, 6).Value = reportPath
        finalWs.Cells(r, 7).Value = r - FIRST_DATA_ROW + 1
        r = r + 1
    Next tbl
    FormatTableReport finalWs, r - 1
    LogTableBackup finalWs, backupPath

    ' Заменяет листы только после подготовки окончательного отчёта; сохраняет старый
    ' отчёт как резервную копию, пока новый лист не получит основное имя.
    finalOldName = GetUniqueTableAuditSheetName("_bak_table_final")
    ws.Name = finalOldName
    finalOldRenamed = True
    On Error GoTo FinalReportCommitFailed
    finalWs.Name = TABLE_REPORT_SHEET
    finalReportCommitted = True
    Set ws = finalWs
    Set finalWs = Nothing
    Application.DisplayAlerts = False
    On Error Resume Next
    registerWb.Worksheets(finalOldName).Delete
    If Err.Number = 0 Then
        finalOldRenamed = False
    Else
        Err.Clear
    End If
    On Error GoTo Failed
    Application.DisplayAlerts = oldAlerts
    ActivateReportSheet registerWb, ws

    MsgBox "Обработка завершена." & vbCrLf & _
           IIf(unconfirmedCount > 0, "ВНИМАНИЕ: у " & unconfirmedCount & " таблиц подпись после обработки не подтвердилась - " & _
               "строки «ОШИБКА» в отчёте, проверьте их в документе." & vbCrLf, "") & _
           "Подписей добавлено или преобразовано: " & insertedCount & vbCrLf & _
           IIf(manualCheckCount > 0, "Подписей для ручной проверки (не менялись): " & manualCheckCount & vbCrLf, "") & _
           "Поля Word обновлены, документ сохранён." & vbCrLf & _
           "Резервная копия: " & backupPath & vbCrLf & _
           "Итоговый список: лист '" & TABLE_REPORT_SHEET & "'.", _
           IIf(unconfirmedCount > 0, vbExclamation, vbInformation), "Нумерация таблиц"
    GoTo CleanExit

FinalReportCommitFailed:
    errText = Err.Description
    On Error Resume Next
    If finalOldRenamed And Not ws Is Nothing Then ws.Name = TABLE_REPORT_SHEET
    If Not finalWs Is Nothing Then
        Application.DisplayAlerts = False
        finalWs.Delete
        Set finalWs = Nothing
    End If
    Application.DisplayAlerts = oldAlerts
    Err.Clear
    On Error GoTo Failed
    GoTo Failed

CommitFailed:
    errText = Err.Description
    On Error Resume Next
    If priorRenamed And Not priorWs Is Nothing Then
        If Not reportCommitted Then priorWs.Name = TABLE_REPORT_SHEET
    End If
    If Not stageWs Is Nothing Then
        If Not reportCommitted Then
            Application.DisplayAlerts = False
            stageWs.Delete
        End If
    End If
    Application.DisplayAlerts = oldAlerts
    Err.Clear
    On Error GoTo Failed

Failed:
    cancelled = (Err.Number = 18)
    If Len(errText) = 0 Then errText = Err.Description
    On Error Resume Next
    If trackChanged And Not wdDoc Is Nothing Then wdDoc.TrackRevisions = oldTrack
    CloseHiddenWord wdApp, wdDoc
    Set wdDoc = Nothing
    Set wdApp = Nothing
    If Not stageWs Is Nothing Then
        If Not reportCommitted Then
            Application.DisplayAlerts = False
            stageWs.Delete
        End If
    End If
    If Not finalWs Is Nothing Then
        Application.DisplayAlerts = False
        finalWs.Delete
        Set finalWs = Nothing
    End If
    If finalOldRenamed And Not finalReportCommitted And Not ws Is Nothing Then
        ws.Name = TABLE_REPORT_SHEET
        finalOldRenamed = False
    End If
    If priorRenamed And Not priorWs Is Nothing And Not reportCommitted Then
        priorWs.Name = TABLE_REPORT_SHEET
    End If
    Application.DisplayAlerts = oldAlerts
    If changesStarted And Len(backupPath) > 0 Then
        Err.Clear
        Set fsoRestore = CreateObject("Scripting.FileSystemObject")
        fsoRestore.CopyFile backupPath, reportPath, True
        restoredOk = (Err.Number = 0)
        If restoredOk Then
            LogTableBackup ws, "Исходный файл восстановлен из копии: " & backupPath
        Else
            Err.Clear
            LogTableBackup ws, "Автовосстановление не удалось. Резервная копия сохранена: " & backupPath
        End If
    End If
    On Error GoTo 0
    If cancelled Then
        MsgBox "Операция отменена пользователем." & RestoreNote(restoredOk, backupPath), _
               vbExclamation, "Проверка таблиц"
    Else
        MsgBox "Ошибка при проверке или нумерации таблиц: " & errText & RestoreNote(restoredOk, backupPath), _
               vbExclamation, "Проверка таблиц"
    End If

CleanExit:
    On Error Resume Next
    CloseHiddenWord wdApp, wdDoc
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
End Sub

' Невидимый Word без диалогов и без запуска макросов документа:
' любое окно в невидимом Word заставило бы Excel ждать бесконечно.
Private Function OpenHiddenWord() As Object
    Dim wdApp As Object
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    wdApp.DisplayAlerts = WD_ALERTS_NONE
    On Error Resume Next
    wdApp.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
    On Error GoTo 0
    Set OpenHiddenWord = wdApp
End Function

' Закрывает документ без сохранения и выходит из Word, не спрашивая о шаблоне Normal.
Private Sub CloseHiddenWord(ByRef wdApp As Object, ByRef wdDoc As Object)
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close 0
    Set wdDoc = Nothing
    If Not wdApp Is Nothing Then
        wdApp.NormalTemplate.Saved = True
        wdApp.Quit 0
    End If
    Set wdApp = Nothing
End Sub

Private Sub ActivateReportSheet(ByVal wb As Workbook, ByVal ws As Worksheet)
    On Error Resume Next
    wb.Activate
    ws.Activate
    On Error GoTo 0
End Sub

' Ближайший непустой абзац над таблицей (не дальше CAPTION_SEARCH_LIMIT абзацев).
' Поиск идёт вверх от таблицы и останавливается на предыдущей таблице: подпись
' соседней таблицы не должна засчитываться этой.
Private Function GetCaptionParagraphBeforeTable(ByVal wdDoc As Object, ByVal tableStart As Long) As Object
    Dim p As Object
    Dim txt As String
    Dim checkedParagraphs As Long

    If tableStart <= 0 Then Exit Function
    Set p = wdDoc.Range(tableStart - 1, tableStart - 1).Paragraphs(1)
    Do While Not p Is Nothing
        If p.Range.Information(WD_WITHIN_TABLE) Then Exit Do
        txt = CleanParagraphText(CStr(p.Range.Text))
        If Len(txt) > 0 Then
            Set GetCaptionParagraphBeforeTable = p.Range
            Exit Function
        End If
        checkedParagraphs = checkedParagraphs + 1
        If checkedParagraphs >= CAPTION_SEARCH_LIMIT Then Exit Do
        Set p = p.Previous
    Loop
End Function

Private Function CleanParagraphText(ByVal txt As String) As String
    txt = Replace(txt, vbCr, "")
    txt = Replace(txt, Chr$(7), "")
    txt = Replace(txt, ChrW(160), " ")
    CleanParagraphText = Trim$(txt)
End Function

' Разбирает подпись над таблицей. captionText - текст подписи для отчёта.
Private Function ClassifyTableCaption(ByVal wdDoc As Object, ByVal tableStart As Long, _
                                      ByRef captionText As String) As Long
    Dim para As Object
    Dim fld As Object
    Dim codeText As String
    Dim hasTableSeq As Boolean
    Dim fieldCount As Long
    Dim listString As String

    captionText = ""
    Set para = GetCaptionParagraphBeforeTable(wdDoc, tableStart)
    If para Is Nothing Then
        ClassifyTableCaption = CAP_MISSING
        Exit Function
    End If

    captionText = CleanParagraphText(CStr(para.Text))
    On Error Resume Next
    listString = CStr(para.ListFormat.ListString)
    On Error GoTo 0

    For Each fld In para.Fields
        fieldCount = fieldCount + 1
        codeText = Trim$(CStr(fld.Code.Text))
        If InStr(1, codeText, "SEQ", vbTextCompare) > 0 And _
           InStr(1, codeText, "ТАБЛИЦА", vbTextCompare) > 0 Then hasTableSeq = True
    Next fld

    If hasTableSeq Then
        If StrComp(Left$(captionText, Len("Таблица")), "Таблица", vbTextCompare) = 0 Then
            ClassifyTableCaption = CAP_HAS_SEQ
        Else
            ClassifyTableCaption = CAP_MANUAL   ' поле есть, но подпись выглядит иначе
        End If
        Exit Function
    End If

    If GetManualCaptionPrefixLength(captionText) > 0 Then
        ' Поля внутри подписи (например, SEQ Table из английского Word) или
        ' автонумерация абзаца не позволяют точно вырезать номер - не трогаем.
        If fieldCount > 0 Or Len(listString) > 0 Then
            ClassifyTableCaption = CAP_MANUAL
        Else
            ClassifyTableCaption = CAP_CONVERT
        End If
    ElseIf LooksLikeTableCaption(captionText) Or (Len(listString) > 0 And _
           InStr(1, listString, "Табл", vbTextCompare) > 0) Then
        ClassifyTableCaption = CAP_MANUAL       ' «Табл. 1», «Таблица 2.3», «Таблица А.1»
    Else
        ClassifyTableCaption = CAP_MISSING
    End If
End Function

' Похоже на подпись таблицы, но не в простом формате «Таблица N»
Private Function LooksLikeTableCaption(ByVal captionText As String) As Boolean
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True
    ' после «Табл.»/«Таблица» сразу номер: цифра или буква с точкой и цифрой («А.1»);
    ' обычный текст «Таблица ниже показывает…» подписью не считается
    rx.Pattern = "^\s*Табл(ица|\.)?\s*(№|N|#)?\s*([0-9]|[А-ЯA-Z]\.?[0-9])"
    LooksLikeTableCaption = rx.Test(captionText)
End Function

Private Sub InsertCustomTableCaption(ByVal wdApp As Object, ByVal wdDoc As Object, ByVal tbl As Object)
    Dim captionPara As Object
    Dim fld As Object
    Dim prefixRange As Object
    Dim foundSeq As Boolean

    ' Просит Word создать настоящий абзац подписи над таблицей, затем меняет формат
    ' поля SEQ, чтобы оно отображало «Таблица № 1» и последующие номера.
    tbl.Range.Select
    wdApp.Selection.InsertCaption Label:="Таблица", Title:="", Position:=WD_CAPTION_POSITION_ABOVE, ExcludeLabel:=False

    Set captionPara = GetCaptionParagraphBeforeTable(wdDoc, tbl.Range.Start)
    If captionPara Is Nothing Then
        Err.Raise vbObjectError + 710, "InsertCustomTableCaption", _
                  "Не удалось найти созданную подпись над таблицей."
    End If

    For Each fld In captionPara.Fields
        If InStr(1, CStr(fld.Code.Text), "SEQ", vbTextCompare) > 0 And _
           InStr(1, CStr(fld.Code.Text), "ТАБЛИЦА", vbTextCompare) > 0 Then
            ' Удаляет обычный текст подписи перед полем: формат поля
            ' задаёт полный вид подписи.
            If fld.Result.Start > captionPara.Start Then
                Set prefixRange = wdDoc.Range(captionPara.Start, fld.Code.Start - 1)
                If prefixRange.Fields.Count = 0 Then prefixRange.Text = ""
            End If
            fld.Code.Text = SEQ_FIELD_CODE
            fld.Update
            foundSeq = True
            Exit For
        End If
    Next fld

    If Not foundSeq Then
        Err.Raise vbObjectError + 711, "InsertCustomTableCaption", _
                  "Word не создал поле SEQ в подписи над таблицей."
    End If
End Sub

Private Function CreateWordBackupPath(ByVal sourcePath As String) As String
    Dim fso As Object
    Dim parentFolder As String
    Dim baseName As String
    Dim extension As String
    Dim candidate As String
    Dim suffix As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    parentFolder = fso.GetParentFolderName(sourcePath)
    baseName = fso.GetBaseName(sourcePath)
    extension = fso.GetExtensionName(sourcePath)
    candidate = fso.BuildPath(parentFolder, baseName & "_backup_" & Format$(Now, "yyyymmdd_hhnnss") & "." & extension)
    suffix = 1
    Do While fso.FileExists(candidate)
        candidate = fso.BuildPath(parentFolder, baseName & "_backup_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & suffix & "." & extension)
        suffix = suffix + 1
    Loop
    CreateWordBackupPath = candidate
End Function

Private Sub LogTableBackup(ByVal ws As Worksheet, ByVal message As String)
    Dim nextRow As Long
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < FIRST_DATA_ROW Then nextRow = FIRST_DATA_ROW
    ws.Cells(nextRow, 1).Value = "Резервная копия"
    ws.Cells(nextRow, 5).Value = message
    ws.Cells(nextRow, 6).Value = Now
    On Error GoTo 0
End Sub

' Длина префикса «Таблица 5» / «Таблица № 5» в начале подписи; 0 - если формат
' не простой (многоуровневый «Таблица 2.3», буквенный «Таблица А.1» и т.п.).
Private Function GetManualCaptionPrefixLength(ByVal captionText As String) As Long
    Dim rx As Object
    Dim matches As Object
    Dim leadingLength As Long
    Dim nextChar As String
    Dim afterNext As String

    captionText = Replace(captionText, ChrW(160), " ")
    Do While leadingLength < Len(captionText) And _
        (Mid$(captionText, leadingLength + 1, 1) = " " Or Mid$(captionText, leadingLength + 1, 1) = vbTab)
        leadingLength = leadingLength + 1
    Loop
    captionText = Mid$(captionText, leadingLength + 1)
    If Len(captionText) = 0 Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = "^Таблица\s*(№|N|#)?\s*[0-9]+"
    If rx.Test(captionText) Then
        Set matches = rx.Execute(captionText)
        GetManualCaptionPrefixLength = matches(0).Length
        If GetManualCaptionPrefixLength < Len(captionText) Then
            nextChar = Mid$(captionText, GetManualCaptionPrefixLength + 1, 1)
            afterNext = Mid$(captionText, GetManualCaptionPrefixLength + 2, 1)
            If nextChar Like "[.,]" And afterNext Like "[0-9А-ЯA-Zа-яa-z]" Then
                GetManualCaptionPrefixLength = 0      ' «Таблица 2.3», «Таблица 1.А» - многоуровневый номер
            ElseIf nextChar <> " " And nextChar <> vbTab And nextChar <> "." And _
               nextChar <> "," And nextChar <> ":" And nextChar <> ";" And _
               nextChar <> ChrW(8211) And nextChar <> ChrW(8212) And nextChar <> "-" Then
                GetManualCaptionPrefixLength = 0      ' «Таблица 5А», «Таблица 5б» и т.п.
            End If
        End If
    End If
End Function

Private Sub ConvertManualTableCaptionToSequence(ByVal wdDoc As Object, ByVal tbl As Object)
    Dim captionPara As Object
    Dim prefixRange As Object
    Dim insertRange As Object
    Dim fld As Object
    Dim captionText As String
    Dim prefixLength As Long
    Dim leadingLength As Long

    Set captionPara = GetCaptionParagraphBeforeTable(wdDoc, tbl.Range.Start)
    If captionPara Is Nothing Then
        Err.Raise vbObjectError + 713, "ConvertManualTableCaptionToSequence", _
                  "Не найдена существующая подпись над таблицей."
    End If
    If captionPara.Fields.Count > 0 Then
        Err.Raise vbObjectError + 715, "ConvertManualTableCaptionToSequence", _
                  "Подпись содержит поля Word - её нельзя преобразовать автоматически."
    End If

    captionText = Replace(CStr(captionPara.Text), vbCr, "")
    captionText = Replace(captionText, Chr$(7), "")
    captionText = Replace(captionText, ChrW(160), " ")
    Do While leadingLength < Len(captionText) And _
        (Mid$(captionText, leadingLength + 1, 1) = " " Or Mid$(captionText, leadingLength + 1, 1) = vbTab)
        leadingLength = leadingLength + 1
    Loop
    prefixLength = GetManualCaptionPrefixLength(captionText)
    If prefixLength = 0 Then
        Err.Raise vbObjectError + 714, "ConvertManualTableCaptionToSequence", _
                  "Не удалось определить номер в существующей подписи таблицы."
    End If

    Set prefixRange = wdDoc.Range(captionPara.Start + leadingLength, _
                                  captionPara.Start + leadingLength + prefixLength)
    prefixRange.Text = ""
    Set insertRange = wdDoc.Range(captionPara.Start + leadingLength, captionPara.Start + leadingLength)
    Set fld = wdDoc.Fields.Add(insertRange, WD_FIELD_SEQUENCE, "Таблица", False)
    fld.Code.Text = SEQ_FIELD_CODE
    fld.Update
End Sub

' Название подписи «Таблица» хранится в самом Word (Application.CaptionLabels),
' а не в документе.
Private Sub EnsureTableCaptionLabel(ByVal wdApp As Object)
    Dim label As Object
    Dim savedErrNumber As Long
    Dim savedErrDescription As String

    On Error Resume Next
    Err.Clear
    Set label = wdApp.CaptionLabels("Таблица")
    savedErrNumber = Err.Number
    savedErrDescription = Err.Description
    On Error GoTo 0

    If savedErrNumber = 18 Or (savedErrNumber <> 0 And savedErrNumber <> 5941) Then
        Err.Raise savedErrNumber, "EnsureTableCaptionLabel", savedErrDescription
    End If
    If label Is Nothing Then wdApp.CaptionLabels.Add "Таблица"
End Sub

Private Sub UpdateAllStoryFields(ByVal wdDoc As Object)
    Dim story As Object
    Dim currentRange As Object

    For Each story In wdDoc.StoryRanges
        Set currentRange = story
        Do While Not currentRange Is Nothing
            currentRange.Fields.Update
            Set currentRange = currentRange.NextStoryRange
        Loop
    Next story
End Sub

Private Function GetUniqueTableAuditSheetName(ByVal baseName As String) As String
    Dim candidate As String
    Dim suffix As Long
    candidate = Left$(baseName, 31)
    suffix = 1
    Do While TableAuditSheetExists(candidate)
        candidate = Left$(baseName, 27) & "_" & CStr(suffix)
        suffix = suffix + 1
    Loop
    GetUniqueTableAuditSheetName = candidate
End Function

Private Function TableAuditSheetExists(ByVal sheetName As String) As Boolean
    Dim probe As Worksheet
    On Error Resume Next
    Set probe = GetRegisterWorkbook().Worksheets(sheetName)
    TableAuditSheetExists = Not probe Is Nothing
    On Error GoTo 0
End Function

Private Sub FormatTableReport(ByVal ws As Worksheet, ByVal lastRow As Long)
    If lastRow < 1 Then lastRow = 1
    With ws.Range("A1:G" & lastRow)
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
    With ws.Range("A1:G1")
        .Interior.Color = RGB(198, 224, 180)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns("A:B").AutoFit
    ws.Columns("C").ColumnWidth = 45
    ws.Columns("D:E").ColumnWidth = 24
    ws.Columns("F").ColumnWidth = 50
    ws.Columns("G").Hidden = True
End Sub

' Текст для сообщения: восстановлен ли документ из резервной копии
Private Function RestoreNote(ByVal restoredOk As Boolean, ByVal backupPath As String) As String
    If Len(backupPath) = 0 Then Exit Function
    If restoredOk Then
        RestoreNote = vbCrLf & "Документ восстановлен из резервной копии: " & backupPath
    Else
        RestoreNote = vbCrLf & "Документ НЕ восстанавливался (изменения уже сохранены или копия недоступна)." & _
                      vbCrLf & "Резервная копия: " & backupPath
    End If
End Function
