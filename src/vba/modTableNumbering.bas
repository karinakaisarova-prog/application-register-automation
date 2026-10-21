Attribute VB_Name = "modTableNumbering"
Option Explicit

' Checks every table exposed by Word's document Tables collection.
' Produces a preview on an Excel sheet before modifying the Word document.
' Missing captions are inserted as Word SEQ fields, then document fields are updated.
' Word is automated late-bound; Microsoft Word must be installed.

Private Const TABLE_REPORT_SHEET As String = "Проверка таблиц"
Private Const FIRST_DATA_ROW As Long = 2
Private Const WD_STATISTIC_PAGES As Long = 2
Private Const WD_CAPTION_POSITION_ABOVE As Long = 0
Private Const WD_FIELD_SEQUENCE As Long = 12

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

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts

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

    Set ws = GetOrCreateTableReportSheet(TABLE_REPORT_SHEET)
    ws.Cells.Clear
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
        Else
            ws.Cells(r, 5).Value = "Требуется добавить подпись"
            missingCount = missingCount + 1
        End If
        ws.Cells(r, 6).Value = reportPath
        ws.Cells(r, 7).Value = r - FIRST_DATA_ROW + 1
        r = r + 1
    Next tbl

    FormatTableReport ws, r - 1
    ws.Activate

    If tableCount = 0 Then
        MsgBox "В документе не найдено таблиц через коллекцию Word Tables.", vbInformation, "Проверка таблиц"
        GoTo CleanExit
    End If

    If missingCount = 0 Then
        MsgBox "Проверка завершена. Найдено таблиц: " & tableCount & _
               vbCrLf & "У всех таблиц найдены подписи с полем SEQ «Таблица»." & _
               vbCrLf & "Отчёт проверки записан на лист '" & TABLE_REPORT_SHEET & "'.", _
               vbInformation, "Проверка таблиц"
        GoTo CleanExit
    End If

    answer = MsgBox("Предварительная проверка завершена." & vbCrLf & _
                    "Всего таблиц: " & tableCount & vbCrLf & _
                    "Таблиц без поля SEQ «Таблица»: " & missingCount & vbCrLf & vbCrLf & _
                    "Добавить подписи над таблицами без нумерации, обновить поля документа и сохранить Word-отчёт?" & _
                    vbCrLf & "Перед изменением рекомендуется иметь резервную копию.", _
                    vbYesNo + vbQuestion + vbDefaultButton2, "Подтверждение нумерации")
    If answer <> vbYes Then
        MsgBox "Изменения не внесены. Предварительный отчёт сохранён на листе '" & TABLE_REPORT_SHEET & "'.", _
               vbInformation, "Проверка таблиц"
        GoTo CleanExit
    End If

    ' Create a same-folder backup only after the user confirms the proposed changes.
    backupPath = CreateWordBackupPath(reportPath)
    FileCopy reportPath, backupPath
    LogTableBackup ws, backupPath
    changesStarted = True

    EnsureTableCaptionLabel wdDoc

    ' Iterate backwards so inserting captions does not disturb positions of tables not yet processed.
    For r = tableCount To 1 Step -1
        DoEvents
        Set tbl = wdDoc.Tables(r)
        If Not HasTableSequenceField(wdDoc, tbl.Range.Start) Then
            tbl.Range.Select
            wdApp.Selection.InsertCaption "Таблица", "", , WD_CAPTION_POSITION_ABOVE, ""
            insertedCount = insertedCount + 1
            Application.StatusBar = "Добавление подписей: " & insertedCount & " из " & missingCount
        End If
    Next r

    wdDoc.Fields.Update
    UpdateAllStoryFields wdDoc
    wdDoc.Save

    ' Refresh the preview with the final state.
    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array("№ таблицы", "Страница", "Подпись перед таблицей", _
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

        ws.Cells(r, 1).Value = r - FIRST_DATA_ROW + 1
        ws.Cells(r, 2).Value = pageNo
        ws.Cells(r, 3).Value = captionText
        ws.Cells(r, 4).Value = IIf(hasSeq, "Да", "Нет")
        ws.Cells(r, 5).Value = IIf(hasSeq, "Нумерация найдена", "Проверьте вручную")
        ws.Cells(r, 6).Value = reportPath
        ws.Cells(r, 7).Value = r - FIRST_DATA_ROW + 1
        r = r + 1
    Next tbl
    FormatTableReport ws, r - 1

    MsgBox "Обработка завершена." & vbCrLf & _
           "Подписей добавлено: " & insertedCount & vbCrLf & _
           "Поля Word обновлены, документ сохранён." & vbCrLf & _
           "Резервная копия: " & backupPath & vbCrLf & _
           "Итоговый список: лист '" & TABLE_REPORT_SHEET & "'.", _
           vbInformation, "Нумерация таблиц"
    GoTo CleanExit

Failed:
    cancelled = (Err.Number = 18)
    errText = Err.Description
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
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
    Application.StatusBar = False
    Application.EnableCancelKey = xlInterrupt
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
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < FIRST_DATA_ROW Then nextRow = FIRST_DATA_ROW
    ws.Cells(nextRow, 1).Value = "Резервная копия"
    ws.Cells(nextRow, 5).Value = message
    ws.Cells(nextRow, 6).Value = Now
    On Error GoTo 0
End Sub

Private Function HasTableSequenceField(ByVal wdDoc As Object, ByVal tableStart As Long) As Boolean
    Dim p As Long
    Dim probe As Object
    Dim para As Object
    Dim fld As Object
    Dim codeText As String
    Dim stepsBack As Long

    ' Search the few paragraphs immediately before the table, stopping at the first
    ' non-empty paragraph. A matching Word SEQ field is the criterion for a numbered caption.
    For stepsBack = 1 To 4
        p = tableStart - stepsBack
        If p < 0 Then Exit For
        On Error Resume Next
        Set probe = wdDoc.Range(0, tableStart - stepsBack + 1)
        Set para = probe.Paragraphs.Last.Range
        If Err.Number <> 0 Or para Is Nothing Then
            Err.Clear
            On Error GoTo 0
            Exit For
        End If
        On Error GoTo 0

        If para.Fields.Count > 0 Then
            For Each fld In para.Fields
                codeText = UCase$(Trim$(CStr(fld.Code.Text)))
                If InStr(1, codeText, "SEQ", vbTextCompare) > 0 And _
                   InStr(1, codeText, "ТАБЛИЦА", vbTextCompare) > 0 Then
                    HasTableSequenceField = True
                    Exit Function
                End If
            Next fld
        End If

        If Len(Trim$(Replace(Replace(CStr(para.Text), vbCr, ""), Chr$(7), ""))) > 0 Then Exit For
    Next stepsBack
End Function

Private Function GetPreviousCaptionText(ByVal wdDoc As Object, ByVal tableStart As Long) As String
    Dim probe As Object
    Dim para As Object
    Dim txt As String

    On Error Resume Next
    If tableStart <= 0 Then Exit Function
    Set probe = wdDoc.Range(0, tableStart)
    Set para = probe.Paragraphs.Last.Range
    If Not para Is Nothing Then
        txt = CStr(para.Text)
        txt = Replace(txt, vbCr, "")
        txt = Replace(txt, Chr$(7), "")
        GetPreviousCaptionText = Trim$(txt)
    End If
    On Error GoTo 0
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

Private Function GetOrCreateTableReportSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateTableReportSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateTableReportSheet Is Nothing Then
        Set GetOrCreateTableReportSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateTableReportSheet.Name = sheetName
    End If
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
