Attribute VB_Name = "modBulletListAudit"
Option Explicit

' Finds consecutive paragraphs that use typed dash/bullet characters instead of
' Word's native list formatting. Shows a preview before any edits.
' Tests should be run on a copy of the report. Microsoft Word must be installed.

Private Const BULLET_REPORT_SHEET As String = "Проверка списков"
Private Const FIRST_DATA_ROW As Long = 2
Private Const WD_LIST_NO_NUMBERING As Long = 0

Public Sub CheckAndFixWordBulletLists()
    Dim fd As FileDialog
    Dim reportPath As String
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim ws As Worksheet
    Dim para As Object
    Dim i As Long, j As Long, pCount As Long
    Dim runStart As Long, runEnd As Long, runLength As Long
    Dim rowOut As Long, groupNo As Long
    Dim pageNo As Long
    Dim txt As String
    Dim markerOffset As Long, markerLength As Long
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim errText As String
    Dim cancelled As Boolean
    Dim answer As VbMsgBoxResult
    Dim groupsFound As Long, paragraphsFound As Long, paragraphsFixed As Long

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Выберите Word-отчёт для проверки маркированных списков"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Документы Word", "*.doc;*.docx;*.docm"
        If .Show <> -1 Then GoTo CleanExit
        reportPath = .SelectedItems(1)
    End With

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(reportPath, False, True, False)

    Set ws = GetOrCreateBulletReportSheet(BULLET_REPORT_SHEET)
    ws.Cells.Clear
    ws.Range("A1:G1").Value = Array("Группа", "Страница", "Текст абзаца", "Найденная маркировка", _
                                    "Предлагаемое действие", "Начальный абзац (служебное)", "Конечный абзац (служебное)")
    rowOut = FIRST_DATA_ROW
    pCount = wdDoc.Paragraphs.Count
    i = 1

    Do While i <= pCount
        DoEvents
        Set para = wdDoc.Paragraphs(i)
        txt = CleanBulletParagraphText(CStr(para.Range.Text))
        markerOffset = 0
        markerLength = 0

        If Not IsNativeWordList(para) Then
            If GetTypedMarker(txt, markerOffset, markerLength) Then
                runStart = i
                j = i
                Do While j <= pCount
                    Set para = wdDoc.Paragraphs(j)
                    txt = CleanBulletParagraphText(CStr(para.Range.Text))
                    markerOffset = 0
                    markerLength = 0
                    If IsNativeWordList(para) Then Exit Do
                    If Not GetTypedMarker(txt, markerOffset, markerLength) Then Exit Do
                    j = j + 1
                Loop
                runEnd = j - 1
                runLength = runEnd - runStart + 1

                ' Avoid converting isolated dash-led paragraphs: they may be ordinary prose.
                If runLength >= 2 Then
                    groupNo = groupNo + 1
                    groupsFound = groupsFound + 1
                    For i = runStart To runEnd
                        Set para = wdDoc.Paragraphs(i)
                        txt = CleanBulletParagraphText(CStr(para.Range.Text))
                        markerOffset = 0
                        markerLength = 0
                        Call GetTypedMarker(txt, markerOffset, markerLength)
                        pageNo = 0
                        On Error Resume Next
                        pageNo = para.Range.Information(3) ' wdActiveEndPageNumber
                        On Error GoTo Failed
                        ws.Cells(rowOut, 1).Value = groupNo
                        ws.Cells(rowOut, 2).Value = pageNo
                        ws.Cells(rowOut, 3).Value = txt
                        ws.Cells(rowOut, 4).Value = Left$(Trim$(txt), markerLength)
                        ws.Cells(rowOut, 5).Value = "Удалить ручной маркер и применить список Word"
                        ws.Cells(rowOut, 6).Value = i
                        ws.Cells(rowOut, 7).Value = runEnd
                        rowOut = rowOut + 1
                        paragraphsFound = paragraphsFound + 1
                    Next i
                End If
                i = j
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop

    FormatBulletReport ws, rowOut - 1
    ws.Activate

    If paragraphsFound = 0 Then
        MsgBox "Не найдено групп из двух или более подряд идущих абзацев с ручными маркерами, которые можно уверенно преобразовать." & _
               vbCrLf & "Одиночные абзацы с тире намеренно не меняются, чтобы не испортить обычный текст." & _
               vbCrLf & "Отчёт проверки записан на лист '" & BULLET_REPORT_SHEET & "'.", _
               vbInformation, "Проверка списков"
        GoTo CleanExit
    End If

    answer = MsgBox("Предварительная проверка завершена." & vbCrLf & _
                    "Найдено групп: " & groupsFound & vbCrLf & _
                    "Абзацев-кандидатов: " & paragraphsFound & vbCrLf & vbCrLf & _
                    "Преобразовать эти группы в настоящие маркированные списки Word и сохранить документ?" & _
                    vbCrLf & "Перед исправлением рекомендуется сделать копию файла.", _
                    vbYesNo + vbQuestion + vbDefaultButton2, "Подтверждение исправлений")
    If answer <> vbYes Then
        MsgBox "Изменения не внесены. Предварительный отчёт сохранён на листе '" & BULLET_REPORT_SHEET & "'.", _
               vbInformation, "Проверка списков"
        GoTo CleanExit
    End If

    ' Process groups from the end of the document backwards so earlier paragraph indexes remain valid.
    For i = pCount To 1 Step -1
        If IsNumeric(ws.Cells(FIRST_DATA_ROW, 6).Value) Then Exit For
    Next i

    Dim groupStartRows() As Long, groupEndRows() As Long, groupCount As Long
    Dim rr As Long, priorGroup As Variant
    ReDim groupStartRows(1 To groupsFound)
    ReDim groupEndRows(1 To groupsFound)
    rr = FIRST_DATA_ROW
    priorGroup = Empty
    Do While rr < rowOut
        If IsEmpty(priorGroup) Or ws.Cells(rr, 1).Value <> priorGroup Then
            groupCount = groupCount + 1
            groupStartRows(groupCount) = CLng(ws.Cells(rr, 6).Value)
            groupEndRows(groupCount) = CLng(ws.Cells(rr, 7).Value)
            priorGroup = ws.Cells(rr, 1).Value
        End If
        rr = rr + 1
    Loop

    Dim g As Long, k As Long, firstP As Long, lastP As Long
    Dim wordRange As Object, currentPara As Object
    Dim rawText As String, offset As Long, removeLen As Long
    For g = groupCount To 1 Step -1
        DoEvents
        firstP = groupStartRows(g)
        lastP = groupEndRows(g)
        For k = lastP To firstP Step -1
            Set currentPara = wdDoc.Paragraphs(k)
            rawText = CStr(currentPara.Range.Text)
            offset = 0
            removeLen = 0
            If GetTypedMarker(CleanBulletParagraphText(rawText), offset, removeLen) Then
                ' offset includes leading whitespace; preserve indentation before the marker.
                Dim leadingSpaces As Long, cleanText As String
                cleanText = CleanBulletParagraphText(rawText)
                leadingSpaces = 0
                Do While leadingSpaces < Len(cleanText) And _
                    (Mid$(cleanText, leadingSpaces + 1, 1) = " " Or Mid$(cleanText, leadingSpaces + 1, 1) = vbTab)
                    leadingSpaces = leadingSpaces + 1
                Loop
                Set wordRange = wdDoc.Range(currentPara.Range.Start + leadingSpaces, _
                                            currentPara.Range.Start + leadingSpaces + removeLen)
                wordRange.Delete
            End If
        Next k
        Set wordRange = wdDoc.Range(wdDoc.Paragraphs(firstP).Range.Start, wdDoc.Paragraphs(lastP).Range.End)
        wordRange.ListFormat.ApplyBulletDefault
        paragraphsFixed = paragraphsFixed + (lastP - firstP + 1)
        Application.StatusBar = "Исправление маркированных списков: группа " & (groupCount - g + 1) & " из " & groupCount
    Next g

    wdDoc.Save
    ' Update the audit sheet status; preserve the original preview text.
    ws.Range("E2:E" & rowOut - 1).Value = "Исправлено: применён маркированный список Word"
    MsgBox "Исправление завершено." & vbCrLf & _
           "Групп преобразовано: " & groupCount & vbCrLf & _
           "Абзацев обработано: " & paragraphsFixed & vbCrLf & _
           "Документ сохранён. Проверьте визуально отступы и переносы.", _
           vbInformation, "Маркированные списки Word"
    GoTo CleanExit

Failed:
    cancelled = (Err.Number = 18)
    errText = Err.Description
    On Error Resume Next
    If Not wdDoc Is Nothing Then
        If Not cancelled Then wdDoc.Close False
    End If
    If Not wdApp Is Nothing Then wdApp.Quit
    On Error GoTo 0
    If cancelled Then
        MsgBox "Операция отменена. Если изменения уже начались, проверьте документ и восстановите копию при необходимости.", _
               vbExclamation, "Проверка списков"
    Else
        MsgBox "Ошибка при проверке или исправлении маркированных списков: " & errText, _
               vbExclamation, "Проверка списков"
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

Private Function GetTypedMarker(ByVal paragraphText As String, ByRef markerOffset As Long, ByRef markerLength As Long) As Boolean
    Dim i As Long, ch As String, marker As String
    paragraphText = Replace(paragraphText, vbCr, "")
    paragraphText = Replace(paragraphText, Chr$(7), "")
    i = 1
    Do While i <= Len(paragraphText) And (Mid$(paragraphText, i, 1) = " " Or Mid$(paragraphText, i, 1) = vbTab)
        i = i + 1
    Loop
    If i > Len(paragraphText) Then Exit Function
    ch = Mid$(paragraphText, i, 1)
    If ch <> "-" And ch <> ChrW(8211) And ch <> ChrW(8212) And ch <> ChrW(8226) Then Exit Function
    If i = Len(paragraphText) Then Exit Function
    If Mid$(paragraphText, i + 1, 1) <> " " And Mid$(paragraphText, i + 1, 1) <> vbTab Then Exit Function

    markerOffset = i - 1
    marker = ch
    i = i + 1
    Do While i <= Len(paragraphText) And (Mid$(paragraphText, i, 1) = " " Or Mid$(paragraphText, i, 1) = vbTab)
        i = i + 1
    Loop
    markerLength = i - (markerOffset + 1)
    GetTypedMarker = True
End Function

Private Function IsNativeWordList(ByVal para As Object) As Boolean
    On Error Resume Next
    IsNativeWordList = (CLng(para.Range.ListFormat.ListType) <> WD_LIST_NO_NUMBERING)
    If Err.Number <> 0 Then
        Err.Clear
        IsNativeWordList = False
    End If
    On Error GoTo 0
End Function

Private Function CleanBulletParagraphText(ByVal value As String) As String
    value = Replace(value, vbCr, "")
    value = Replace(value, Chr$(7), "")
    CleanBulletParagraphText = value
End Function

Private Function GetOrCreateBulletReportSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateBulletReportSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateBulletReportSheet Is Nothing Then
        Set GetOrCreateBulletReportSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateBulletReportSheet.Name = sheetName
    End If
End Function

Private Sub FormatBulletReport(ByVal ws As Worksheet, ByVal lastRow As Long)
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
    ws.Columns("C").ColumnWidth = 60
    ws.Columns("D:E").ColumnWidth = 34
    ws.Columns("F:G").Hidden = True
End Sub
