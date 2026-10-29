Attribute VB_Name = "modTableNumbering"
Option Explicit

' Проверяет все таблицы документа Word.
' Перед изменением документа формирует предварительный отчёт на листе Excel.
' Для таблиц без подписи добавляет автоматические поля SEQ Word и обновляет поля документа.
' Для выполнения требуется Microsoft Word.

Private Const TABLE_REPORT_SHEET As String = "Проверка таблиц"
Private Const FIRST_DATA_ROW As Long = 2
Private Const WD_STATISTIC_PAGES As Long = 2
Private Const WD_CAPTION_POSITION_ABOVE As Long = 0
Private Const WD_FIELD_SEQUENCE As Long = 12
Private Const WD_WITHIN_TABLE As Long = 12

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
    Dim insertedCount As Long
    Dim pageNo As Long
    Dim captionText As String
    Dim hasSeq As Boolean
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
        .Title = "Выберите основной Word-отчёт для проверки таблиц"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Документы Word", "*.doc;*.docx;*.docm"
        If .Show <> -1 Then GoTo CleanExit
        reportPath = .SelectedItems(1)
    End With

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(reportPath, False, False, False)

    ' Формирует новый отчёт на временном листе и сохраняет предыдущий отчёт
    ' до успешного завершения проверки.
    On Error Resume Next
    Set priorWs = GetRegisterWorkbook().Worksheets(TABLE_REPORT_SHEET)
    On Error GoTo Failed
    stageName = GetUniqueTableAuditSheetName("_tmp_table_audit")
    Set stageWs = GetRegisterWorkbook().Worksheets.Add(After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count))
    stageWs.Name = stageName
    Set ws = stageWs
    ws.Range("A1:G1").Value = Array("№ таблицы", "Страница", "Подпись перед таблицей", _
                                    "Поле SEQ «Таблица»", "Результат проверки", "Путь к отчёту", "Индекс таблицы")
    r = FIRST_DATA_ROW
    tableCount = wdDoc.Tables.Count

    For Each tbl In wdDoc.Tables
        DoEvents
        pageNo = 0
        captionText = GetPreviousCaptionText(wdDoc, tbl.Range.Start)
        hasSeq = HasTableSequenceField(wdDoc, tbl.Range.Start)

        On Error Resume Next
        pageNo = tbl.Range.Information(3) ' wdActiveEndPageNumber
        On Error GoTo Failed

        ws.Cells(r, 1).Value = r - FIRST_DATA_ROW + 1
        ws.Cells(r, 2).Value = pageNo
        ws.Cells(r, 3).Value = captionText
        ws.Cells(r, 4).Value = IIf(hasSeq, "Да", "Нет")
        If hasSeq Then
            ws.Cells(r, 5).Value = "Нумерация найдена"
        ElseIf IsTableCaptionText(captionText) Then
            ws.Cells(r, 5).Value = "Текстовая подпись без SEQ — ручная проверка"
            manualCaptionCount = manualCaptionCount + 1
        Else
            ws.Cells(r, 5).Value = "Требуется добавить подпись"
            missingCount = missingCount + 1
        End If
        ws.Cells(r, 6).Value = reportPath
        ws.Cells(r, 7).Value = r - FIRST_DATA_ROW + 1
        r = r + 1
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
    ws.Activate

    If tableCount = 0 Then
        MsgBox "В документе не найдено таблиц через коллекцию Word Tables.", vbInformation, "Проверка таблиц"
        GoTo CleanExit
    End If

    If missingCount = 0 Then
        If manualCaptionCount > 0 Then
            MsgBox "Проверка завершена. Таблиц: " & tableCount & _
                   vbCrLf & "Текстовых подписей без поля SEQ: " & manualCaptionCount & _
                   vbCrLf & "Эти подписи не изменены автоматически, чтобы не создавать дубликаты. Проверьте их вручную." & _
                   vbCrLf & "Отчёт проверки записан на лист '" & TABLE_REPORT_SHEET & "'.", _
                   vbExclamation, "Проверка таблиц"
        Else
            MsgBox "Проверка завершена. Найдено таблиц: " & tableCount & _
                   vbCrLf & "У всех таблиц найдены подписи с полем SEQ «Таблица»." & _
                   vbCrLf & "Отчёт проверки записан на лист '" & TABLE_REPORT_SHEET & "'.", _
                   vbInformation, "Проверка таблиц"
        End If
        GoTo CleanExit
    End If

    answer = MsgBox("Предварительная проверка завершена." & vbCrLf & _
                    "Всего таблиц: " & tableCount & vbCrLf & _
                    "Таблиц без подписи: " & missingCount & vbCrLf & _
                    "Текстовых подписей без SEQ (будут пропущены): " & manualCaptionCount & vbCrLf & vbCrLf & _
                    "Добавить подписи только таблицам без подписи, обновить поля и сохранить Word-отчёт?" & _
                    vbCrLf & "Текстовые подписи без SEQ автоматически не заменяются, чтобы избежать дубликатов." & _
                    vbCrLf & "Перед изменением рекомендуется иметь резервную копию.", _
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

    EnsureTableCaptionLabel wdDoc

    ' Обрабатывает таблицы с конца документа, чтобы вставка подписей не смещала расположение ещё не проверенных таблиц.
    For r = tableCount To 1 Step -1
        DoEvents
        Set tbl = wdDoc.Tables(r)
        captionText = GetPreviousCaptionText(wdDoc, tbl.Range.Start)
        If Not HasTableSequenceField(wdDoc, tbl.Range.Start) And _
           Not IsTableCaptionText(captionText) Then
            InsertCustomTableCaption wdApp, wdDoc, tbl
            insertedCount = insertedCount + 1
            Application.StatusBar = "Добавление подписей: " & insertedCount & " из " & missingCount
        End If
    Next r

    wdDoc.Fields.Update
    UpdateAllStoryFields wdDoc
    wdDoc.Save

    ' Формирует окончательный отчёт на отдельном временном листе. Текущий
    ' предварительный отчёт сохраняется, пока новый не будет полностью заполнен.
    finalStageName = GetUniqueTableAuditSheetName("_tmp_table_final")
    Set finalWs = GetRegisterWorkbook().Worksheets.Add(After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count))
    finalWs.Name = finalStageName
    finalWs.Range("A1:G1").Value = Array("№ таблицы", "Страница", "Подпись перед таблицей", _
                                         "Поле SEQ «Таблица»", "Результат проверки", "Путь к отчёту", "Индекс таблицы")
    r = FIRST_DATA_ROW
    For Each tbl In wdDoc.Tables
        DoEvents
        pageNo = 0
        captionText = GetPreviousCaptionText(wdDoc, tbl.Range.Start)
        hasSeq = HasTableSequenceField(wdDoc, tbl.Range.Start)
        On Error Resume Next
        pageNo = tbl.Range.Information(3)
        On Error GoTo Failed

        finalWs.Cells(r, 1).Value = r - FIRST_DATA_ROW + 1
        finalWs.Cells(r, 2).Value = pageNo
        finalWs.Cells(r, 3).Value = captionText
        finalWs.Cells(r, 4).Value = IIf(hasSeq, "Да", "Нет")
        If hasSeq Then
            finalWs.Cells(r, 5).Value = "Нумерация найдена"
        ElseIf IsTableCaptionText(captionText) Then
            finalWs.Cells(r, 5).Value = "Подпись без SEQ — проверьте вручную"
        Else
            finalWs.Cells(r, 5).Value = "Проверьте вручную"
        End If
        finalWs.Cells(r, 6).Value = reportPath
        finalWs.Cells(r, 7).Value = r - FIRST_DATA_ROW + 1
        r = r + 1
    Next tbl
    FormatTableReport finalWs, r - 1

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
    ' Удаление старого предварительного отчёта — завершающая очистка, а не часть изменения документа Word.
    ' Если Excel не позволяет удалить лист, сохраняет его как резервную копию и не восстанавливает
    ' исходный документ Word после сохранения итогового отчёта.
    On Error Resume Next
    GetRegisterWorkbook().Worksheets(finalOldName).Delete
    If Err.Number = 0 Then
        finalOldRenamed = False
    Else
        Err.Clear
    End If
    On Error GoTo Failed
    Application.DisplayAlerts = oldAlerts
    ws.Activate

    MsgBox "Обработка завершена." & vbCrLf & _
           "Подписей добавлено: " & insertedCount & vbCrLf & _
           "Поля Word обновлены, документ сохранён." & vbCrLf & _
           "Резервная копия: " & backupPath & vbCrLf & _
           "Итоговый список: лист '" & TABLE_REPORT_SHEET & "'.", _
           vbInformation, "Нумерация таблиц"
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
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
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
        If Err.Number = 0 Then
            LogTableBackup ws, "Исходный файл восстановлен из копии: " & backupPath
        Else
            Err.Clear
            LogTableBackup ws, "Автовосстановление не удалось. Резервная копия сохранена: " & backupPath
        End If
    End If
    On Error GoTo 0
    If cancelled Then
        MsgBox "Операция отменена пользователем. Если подтверждение изменений уже было дано, проверьте документ перед повторным запуском.", _
               vbExclamation, "Проверка таблиц"
    Else
        MsgBox "Ошибка при проверке или нумерации таблиц: " & errText, vbExclamation, "Проверка таблиц"
    End If

CleanExit:
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
End Sub

Private Sub InsertCustomTableCaption(ByVal wdApp As Object, ByVal wdDoc As Object, ByVal tbl As Object)
    Dim captionPara As Object
    Dim fld As Object
    Dim prefixRange As Object
    Dim captionStart As Long
    Dim captionEnd As Long
    Dim foundSeq As Boolean

    ' Просит Word создать настоящий абзац подписи над таблицей, затем меняет формат
    ' поля SEQ, чтобы оно отображало «Таблица № 1» и последующие номера.
    tbl.Range.Select
    wdApp.Selection.InsertCaption "Таблица", "", , WD_CAPTION_POSITION_ABOVE, ""

    captionStart = tbl.Range.Start
    Set captionPara = GetCaptionParagraphBeforeTable(wdDoc, captionStart)
    If captionPara Is Nothing Then
        Err.Raise vbObjectError + 710, "InsertCustomTableCaption", _
                  "Не удалось найти созданную подпись над таблицей."
    End If

    For Each fld In captionPara.Fields
        If InStr(1, CStr(fld.Code.Text), "SEQ", vbTextCompare) > 0 And _
           InStr(1, CStr(fld.Code.Text), "ТАБЛИЦА", vbTextCompare) > 0 Then
            ' Удаляет обычный текст подписи перед полем: формат поля
            ' задаёт полный вид подписи.
            If fld.Range.Start > captionPara.Start Then
                Set prefixRange = wdDoc.Range(captionPara.Start, fld.Range.Start)
                prefixRange.Text = ""
            End If

            fld.Code.Text = " SEQ Таблица \# ""'Таблица № '0"" "
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

Private Function GetCaptionParagraphBeforeTable(ByVal wdDoc As Object, ByVal tableStart As Long) As Object
    Dim paraIndex As Long
    Dim para As Object
    Dim txt As String

    For paraIndex = wdDoc.Paragraphs.Count To 1 Step -1
        Set para = wdDoc.Paragraphs(paraIndex).Range
        If para.Start < tableStart And para.End <= tableStart Then
            txt = Replace(CStr(para.Text), vbCr, "")
            txt = Replace(txt, Chr$(7), "")
            If Len(Trim$(txt)) > 0 Then
                Set GetCaptionParagraphBeforeTable = para
                Exit Function
            End If
        End If
    Next paraIndex
End Function

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
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < FIRST_DATA_ROW Then nextRow = FIRST_DATA_ROW
    ws.Cells(nextRow, 1).Value = "Резервная копия"
    ws.Cells(nextRow, 5).Value = message
    ws.Cells(nextRow, 6).Value = Now
    On Error GoTo 0
End Sub

Private Function HasTableSequenceField(ByVal wdDoc As Object, ByVal tableStart As Long) As Boolean
    Dim paraIndex As Long
    Dim checkedParagraphs As Long
    Dim para As Object
    Dim fld As Object
    Dim codeText As String
    Dim paraText As String
    Dim trimmedText As String

    ' Пропускает пустые абзацы и проверяет ближайший непустой
    ' абзац. Наличия поля SEQ недостаточно: оно должно находиться в подписи,
    ' начинающейся с видимого слова «Таблица».
    For paraIndex = wdDoc.Paragraphs.Count To 1 Step -1
        Set para = wdDoc.Paragraphs(paraIndex).Range
        If para.Start >= tableStart Or para.Information(WD_WITHIN_TABLE) Then
            ' Пропускает текущую таблицу и абзацы внутри предыдущих таблиц.
        ElseIf para.End <= tableStart Then
            checkedParagraphs = checkedParagraphs + 1
            paraText = CStr(para.Text)
            paraText = Replace(paraText, vbCr, "")
            paraText = Replace(paraText, Chr$(7), "")
            paraText = Replace(paraText, ChrW(160), " ")
            trimmedText = Trim$(paraText)

            If Len(trimmedText) > 0 Then
                If StrComp(Left$(trimmedText, Len("Таблица")), "Таблица", vbTextCompare) = 0 Then
                    For Each fld In para.Fields
                        codeText = Trim$(CStr(fld.Code.Text))
                        If InStr(1, codeText, "SEQ", vbTextCompare) > 0 And _
                           InStr(1, codeText, "ТАБЛИЦА", vbTextCompare) > 0 Then
                            HasTableSequenceField = True
                            Exit Function
                        End If
                    Next fld
                End If
                Exit For
            End If

            If checkedParagraphs >= 4 Then Exit For
        End If
    Next paraIndex
End Function

Private Function GetPreviousCaptionText(ByVal wdDoc As Object, ByVal tableStart As Long) As String
    Dim paraIndex As Long
    Dim checkedParagraphs As Long
    Dim para As Object
    Dim txt As String

    If tableStart <= 0 Then Exit Function
    For paraIndex = wdDoc.Paragraphs.Count To 1 Step -1
        Set para = wdDoc.Paragraphs(paraIndex).Range
        If para.Start >= tableStart Then
            ' Пропускает абзацы, расположенные в текущей таблице или внутри неё.
        ElseIf para.End <= tableStart Then
            If para.Information(WD_WITHIN_TABLE) = False Then
                checkedParagraphs = checkedParagraphs + 1
                txt = CStr(para.Text)
                txt = Replace(txt, vbCr, "")
                txt = Replace(txt, Chr$(7), "")
                txt = Replace(txt, ChrW(160), " ")
                If Len(Trim$(txt)) > 0 Then
                    GetPreviousCaptionText = Trim$(txt)
                    Exit Function
                End If
                If checkedParagraphs >= 4 Then Exit For
            End If
        End If
    Next paraIndex
End Function

Private Function IsTableCaptionText(ByVal captionText As String) As Boolean
    captionText = Replace(captionText, ChrW(160), " ")
    captionText = Trim$(captionText)
    If Len(captionText) < Len("Таблица") Then Exit Function
    IsTableCaptionText = (StrComp(Left$(captionText, Len("Таблица")), "Таблица", vbTextCompare) = 0)
End Function

Private Sub EnsureTableCaptionLabel(ByVal wdDoc As Object)
    Dim label As Object
    On Error Resume Next
    Set label = wdDoc.CaptionLabels("Таблица")
    If label Is Nothing Then
        Err.Clear
        wdDoc.CaptionLabels.Add "Таблица"
    End If
    On Error GoTo 0
End Sub

Private Sub UpdateAllStoryFields(ByVal wdDoc As Object)
    Dim story As Object
    Dim currentRange As Object
    On Error Resume Next
    For Each story In wdDoc.StoryRanges
        Set currentRange = story
        Do
            currentRange.Fields.Update
            Set currentRange = currentRange.NextStoryRange
        Loop Until currentRange Is Nothing
    Next story
    On Error GoTo 0
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
