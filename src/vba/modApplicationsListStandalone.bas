Attribute VB_Name = "modApplicationsListStandalone"
Option Explicit

' Application Register Automation
' Main list builder for desktop Excel on Windows.
' Standalone version: only the application-register builder; no installer or other VBA modules required.
'
' Entry point: CreateMultiFolderApplicationsListCorrected
' Folder depth: up to MAX_FOLDER_DEPTH below each selected root.
' Manual values in column H ("Раздел") are preserved by full file path.
' No FreezePanes are used.

Private Const HEADER_ROW As Long = 2
Private Const FIRST_DATA_ROW As Long = 3
Private Const MAX_FOLDER_DEPTH As Long = 3
Private Const TEMP_SORT_COL As Long = 11
Private Const DATA_SHEET_NAME As String = "data"

Private mFiles As Collection
Private mSections As Object
Private mSeenPaths As Object
Private mLogText As String
Private mRootPath As String

Public Sub CreateMultiFolderApplicationsListCorrected()
    Dim dlg As FileDialog
    Dim roots As Collection
    Dim rootItem As Variant
    Dim ws As Worksheet
    Dim sheetName As String
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    Dim cancelOrError As Boolean
    Dim errText As String
    Dim backupWs As Worksheet
    Dim targetWasExisting As Boolean
    Dim rebuildStarted As Boolean
    Dim originalSheetName As String
    Dim backupSheetName As String
    Dim oldStatusBar As Variant
    Dim oldCancelKey As XlEnableCancelKey
    Dim damagedWs As Worksheet
    Dim recoverySheetName As String

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    oldStatusBar = Application.StatusBar
    oldCancelKey = Application.EnableCancelKey

    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    mLogText = "Формирование перечня приложений" & vbCrLf

    Set roots = New Collection
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    With dlg
        .Title = "Выберите одну или несколько исходных папок"
        .AllowMultiSelect = True
        If .Show <> -1 Then GoTo CleanExit
        For Each rootItem In .SelectedItems
            roots.Add NormalizeSelectedRootPath(CStr(rootItem))
        Next rootItem
    End With

    sheetName = SafeSheetName(GetFolderName(CStr(roots(1))))
    If LCase$(sheetName) = LCase$(DATA_SHEET_NAME) Then sheetName = "Перечень приложений"

    Set mSections = CreateObject("Scripting.Dictionary")
    mSections.CompareMode = vbTextCompare
    Set ws = Nothing
    If WorksheetExists(sheetName) Then
        Set ws = FindWorksheet(sheetName)
    End If
    targetWasExisting = Not ws Is Nothing
    originalSheetName = sheetName
    If targetWasExisting Then
        SaveExistingSections ws, mSections
        ' Keep a worksheet-level recovery copy before any destructive rebuild.
        backupSheetName = GetUniqueBackupSheetName()
        ws.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Set backupWs = ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        backupWs.Name = backupSheetName
        LogLine "Создана временная резервная копия листа: " & originalSheetName
    End If

    Set mFiles = New Collection
    Set mSeenPaths = CreateObject("Scripting.Dictionary")
    mSeenPaths.CompareMode = vbTextCompare
    For Each rootItem In roots
        mRootPath = CStr(rootItem)
        LogLine "Поиск файлов: " & mRootPath
        CollectFiles mRootPath, mRootPath, 0
    Next rootItem

    If mFiles.Count = 0 Then
        If Not backupWs Is Nothing Then
            backupWs.Delete
            Set backupWs = Nothing
            LogLine "Файлы не найдены; исходный лист оставлен без изменений, временная копия удалена."
        End If
        MsgBox "Подходящие файлы не найдены." & vbCrLf & mLogText, vbInformation, "Перечень приложений"
        GoTo CleanExit
    End If

    If ws Is Nothing Then Set ws = GetOrCreateSheet(sheetName)
    rebuildStarted = True
    BuildOutputSheet ws, mFiles, mSections
    CheckDuplicateApplicationNumbers ws
    ApplyLookups ws
    ws.Columns(TEMP_SORT_COL).Hidden = True
    ws.Activate
    ' Панель кнопок не создаётся: модуль работает самостоятельно.
    If Not backupWs Is Nothing Then
        Application.DisplayAlerts = False
        backupWs.Delete
        Set backupWs = Nothing
        LogLine "Временная резервная копия удалена после успешного формирования."
    End If
    Application.StatusBar = False
    MsgBox "Перечень сформирован. Найдено файлов: " & mFiles.Count & vbCrLf & vbCrLf & mLogText, _
           vbInformation, "Перечень приложений"
    GoTo CleanExit

Failed:
    cancelOrError = (Err.Number = 18)
    errText = Err.Description
    If cancelOrError Then
        LogLine "Операция отменена пользователем."
    Else
        LogLine "Ошибка " & Err.Number & ": " & errText
    End If
    On Error Resume Next
    If Not backupWs Is Nothing Then
        Application.DisplayAlerts = False
        If Not rebuildStarted Then
            ' The original sheet has not been changed; discard only the unused copy.
            backupWs.Delete
            Set backupWs = Nothing
            LogLine "Формирование не началось; исходный лист оставлен без изменений."
        Else
            ' Never delete the damaged worksheet before the backup has taken
            ' the original name successfully. If a rename fails, retain the
            ' backup and leave the remaining worksheet(s) available for recovery.
            Set damagedWs = Nothing
            Set damagedWs = ThisWorkbook.Worksheets(originalSheetName)
            If damagedWs Is Nothing Then
                Err.Clear
                backupWs.Name = originalSheetName
                If Err.Number = 0 Then
                    Set ws = backupWs
                    Set backupWs = Nothing
                    LogLine "Исходный лист восстановлен из резервной копии."
                Else
                    LogLine "Не удалось восстановить исходное имя листа. Резервная копия сохранена: " & backupWs.Name
                End If
            Else
                recoverySheetName = GetUniqueRecoverySheetName()
                Err.Clear
                damagedWs.Name = recoverySheetName
                If Err.Number = 0 Then
                    Err.Clear
                    backupWs.Name = originalSheetName
                    If Err.Number = 0 Then
                        Set ws = backupWs
                        Set backupWs = Nothing
                        Err.Clear
                        damagedWs.Delete
                        If Err.Number = 0 Then
                            LogLine "Исходный лист восстановлен из резервной копии."
                        Else
                            LogLine "Лист восстановлен. Повреждённый лист сохранён под именем: " & recoverySheetName
                        End If
                    Else
                        Err.Clear
                        damagedWs.Name = originalSheetName
                        LogLine "Не удалось переименовать резервную копию в исходное имя. Резервная копия сохранена: " & backupWs.Name
                    End If
                Else
                    LogLine "Не удалось безопасно переименовать исходный лист. Резервная копия сохранена: " & backupWs.Name
                End If
            End If
        End If
    ElseIf Not targetWasExisting Then
        Set ws = Nothing
        If Len(Trim$(sheetName)) > 0 Then
            On Error Resume Next
            Set ws = FindWorksheet(sheetName)
            On Error GoTo 0
        End If
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
        LogLine "Новый неполный лист удалён."
    End If
    Application.StatusBar = False
    MsgBox IIf(cancelOrError, "Операция отменена.", "Не удалось завершить формирование перечня: " & errText) & _
           vbCrLf & vbCrLf & mLogText, vbExclamation, "Перечень приложений"

CleanExit:
    On Error Resume Next
    If Not backupWs Is Nothing Then
        ' A remaining backup may be the only recoverable copy after a failed restore.
        ' Keep it rather than infer success from Err.Number, which may have been cleared.
        LogLine "Резервная копия сохранена для проверки: " & backupWs.Name
    End If
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
End Sub

Private Function GetUniqueRecoverySheetName() As String
    Dim candidate As String
    Dim suffix As Long
    candidate = "_failed_register"
    suffix = 1
    Do While WorksheetExists(candidate)
        candidate = "_failed_reg_" & CStr(suffix)
        suffix = suffix + 1
    Loop
    GetUniqueRecoverySheetName = candidate
End Function

Private Function GetUniqueBackupSheetName() As String
    Dim candidate As String
    Dim suffix As Long
    candidate = "_bak_register"
    suffix = 1
    Do While WorksheetExists(candidate)
        candidate = "_bak_register_" & CStr(suffix)
        suffix = suffix + 1
    Loop
    GetUniqueBackupSheetName = candidate
End Function

Private Sub CollectFiles(ByVal rootPath As String, ByVal currentPath As String, ByVal depth As Long)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim fileItem As Object
    Dim ext As String
    Dim info As Object
    Dim objectName As String

    On Error GoTo FolderError
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(currentPath)

    For Each fileItem In folder.Files
        DoEvents
        Application.StatusBar = "Проверка файла: " & CStr(fileItem.Name)
        ext = LCase$(fso.GetExtensionName(fileItem.Name))
        If IsSupportedExtension(ext) Then
            If Not mSeenPaths.Exists(CStr(fileItem.Path)) Then
                mSeenPaths.Add CStr(fileItem.Path), True
                Set info = ParseFileName(CStr(fileItem.Name))
                objectName = GetObjectName(currentPath, CStr(fileItem.Name), rootPath)
                Dim rec As Object
                Set rec = CreateObject("Scripting.Dictionary")
                rec("Path") = CStr(fileItem.Path)
                rec("FileName") = CStr(fileItem.Name)
                rec("Extension") = ext
                rec("Object") = objectName
                rec("AppNumber") = CStr(info("AppNumber"))
                rec("Title") = CStr(info("Title"))
                rec("IsReport") = IsMainReport(CStr(fileItem.Name))
                rec("Folder") = CStr(currentPath)
                mFiles.Add rec
            Else
                LogLine "Повторный путь пропущен: " & CStr(fileItem.Path)
            End If
        End If
    Next fileItem

    If depth < MAX_FOLDER_DEPTH Then
        For Each subFolder In folder.SubFolders
            DoEvents
            Application.StatusBar = "Обработка папки: " & CStr(subFolder.Path)
            CollectFiles rootPath, CStr(subFolder.Path), depth + 1
        Next subFolder
    End If
    Exit Sub

FolderError:
    If Err.Number = 18 Then
        ' Do not swallow Ctrl+Break; let the top-level handler restore Excel state.
        Err.Raise 18
    End If
    LogLine "Не удалось прочитать папку: " & currentPath & " (" & Err.Description & ")"
    Err.Clear
End Sub

Private Function IsSupportedExtension(ByVal ext As String) As Boolean
    IsSupportedExtension = (ext = "doc" Or ext = "docx" Or ext = "pdf" Or _
                            ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "xlsb")
End Function

Private Function IsMainReport(ByVal fileName As String) As Boolean
    Dim normalized As String
    normalized = LCase$(Trim$(fileName))
    IsMainReport = (Left$(normalized, 6) = "отчет_" Or _
                    Left$(normalized, 6) = "отчёт_" Or _
                    Left$(normalized, 6) = "отчет " Or _
                    Left$(normalized, 6) = "отчёт ")
End Function

Private Function ParseFileName(ByVal fileName As String) As Object
    Dim fso As Object
    Dim baseName As String
    Dim result As Object
    Dim rx As Object
    Dim matches As Object
    Dim appNumber As String
    Dim titleText As String
    Dim p As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    baseName = fso.GetBaseName(fileName)
    Set result = CreateObject("Scripting.Dictionary")
    appNumber = ""
    titleText = baseName

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = "(приложение|приложения|приложению|приложением|приложении|приложений|прил)\.?\s*(№|номер)?\s*([0-9]+[А-ЯA-Z]?)"
    If rx.Test(baseName) Then
        Set matches = rx.Execute(baseName)
        appNumber = matches(0).SubMatches(2)
        p = InStr(1, baseName, matches(0).Value, vbTextCompare)
        titleText = Mid$(baseName, p + Len(matches(0).Value))
        titleText = Replace(titleText, "_", " ")
        titleText = Trim$(titleText)
        Do While Len(titleText) > 0 And (Left$(titleText, 1) = "-" Or Left$(titleText, 1) = "—")
            titleText = Trim$(Mid$(titleText, 2))
        Loop
    ElseIf IsMainReport(baseName) Then
        titleText = Mid$(baseName, 7)
        appNumber = ""
        titleText = Replace(titleText, "_", " ")
    Else
        titleText = Replace(baseName, "_", " ")
    End If

    If Len(titleText) = 0 Then titleText = baseName
    result("AppNumber") = appNumber
    result("Title") = titleText
    Set ParseFileName = result
End Function

Private Function GetObjectName(ByVal currentPath As String, ByVal fileName As String, ByVal selectedRootPath As String) As String
    Dim fso As Object
    Dim folderPath As String
    Dim parentPath As String
    Dim folderName As String
    Dim candidate As String
    Dim objectNames As Variant

    ' Only inspect folders inside the selected root, including the root itself.
    ' A parent folder outside the selected root must not determine the object.
    objectNames = Array("КЗТ", "СМТ", "АМТ", "Саратов", "НТУ", "КТ")
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = RemoveTrailingSlash(currentPath)
    selectedRootPath = RemoveTrailingSlash(selectedRootPath)

    Do While Len(folderPath) > 0
        folderName = GetFolderName(folderPath)
        candidate = FindObjectInText(folderName, objectNames)
        If Len(candidate) > 0 Then
            GetObjectName = candidate
            Exit Function
        End If

        ' Stop after checking the selected root; never inspect its parent.
        If StrComp(folderPath, selectedRootPath, vbTextCompare) = 0 Then Exit Do
        parentPath = fso.GetParentFolderName(folderPath)
        If Len(parentPath) = 0 Or StrComp(parentPath, folderPath, vbTextCompare) = 0 Then Exit Do
        folderPath = RemoveTrailingSlash(parentPath)
    Loop

    ' If no approved object was found in the selected folder chain, inspect
    ' the filename. Return blank rather than guessing an unapproved object.
    GetObjectName = FindObjectInText(fileName, objectNames)
End Function

Private Function FindObjectInText(ByVal sourceText As String, ByVal objectNames As Variant) As String
    Dim rx As Object
    Dim i As Long
    Dim objectName As String

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True

    For i = LBound(objectNames) To UBound(objectNames)
        objectName = CStr(objectNames(i))
        ' Match a complete code/name, not a substring inside another word.
        rx.Pattern = "(^|[^А-Яа-яЁёA-Za-z0-9])" & objectName & "([^А-Яа-яЁёA-Za-z0-9]|$)"
        If rx.Test(sourceText) Then
            FindObjectInText = objectName
            Exit Function
        End If
    Next i
End Function

Private Sub SaveExistingSections(ByVal ws As Worksheet, ByVal target As Object)
    Dim lastRow As Long
    Dim r As Long
    Dim filePath As String
    Dim sectionValue As String

    On Error GoTo Done
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then Exit Sub

    For r = FIRST_DATA_ROW To lastRow
        filePath = GetCellLinkPath(ws.Cells(r, 7))
        sectionValue = CStr(ws.Cells(r, 8).Value)
        If Len(filePath) > 0 And Len(sectionValue) > 0 Then target(NormalizeFilePath(filePath)) = sectionValue
    Next r
Done:
End Sub

Private Function GetCellLinkPath(ByVal c As Range) As String
    On Error Resume Next
    If c.Hyperlinks.Count > 0 Then
        GetCellLinkPath = CStr(c.Hyperlinks(1).Address)
        If Len(GetCellLinkPath) = 0 Then GetCellLinkPath = CStr(c.Hyperlinks(1).SubAddress)
    Else
        GetCellLinkPath = CStr(c.Value)
    End If
    On Error GoTo 0
End Function

Private Function NormalizeFilePath(ByVal value As String) As String
    value = Trim$(value)
    If LCase$(Left$(value, 8)) = "file:///" Then value = Mid$(value, 9)
    value = Replace(value, "%20", " ")
    value = Replace(value, "/", "\")
    NormalizeFilePath = LCase$(value)
End Function

Private Sub BuildOutputSheet(ByVal ws As Worksheet, ByVal files As Collection, ByVal sections As Object)
    Dim headers As Variant
    Dim i As Long
    Dim r As Long
    Dim rec As Object
    Dim sectionValue As String
    Dim sortKey As String

    headers = Array("№ п/п", "Объект", "№ Прил.", "Наименование", "Отдел", _
                    "Ответственное лицо", "Ссылка на файл", "Раздел", "Стр. отчёта", "Кол-во страниц")

    ws.Cells.Clear
    On Error Resume Next
    ws.Hyperlinks.Delete
    On Error GoTo 0

    With ws.Range("A1:J1")
        .Merge
        .Value = "Перечень приложений"
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 30

    For i = LBound(headers) To UBound(headers)
        ws.Cells(HEADER_ROW, i + 1).Value = headers(i)
    Next i
    ws.Cells(HEADER_ROW, TEMP_SORT_COL).Value = "Временный ключ"

    r = FIRST_DATA_ROW
    For i = 1 To files.Count
        Set rec = files(i)
        ws.Cells(r, 2).Value = rec("Object")
        ws.Cells(r, 3).Value = rec("AppNumber")
        ws.Cells(r, 4).Value = rec("Title")
        ws.Cells(r, 7).Value = "Открыть файл"
        ws.Hyperlinks.Add Anchor:=ws.Cells(r, 7), Address:=CStr(rec("Path")), TextToDisplay:="Открыть файл"

        sectionValue = ""
        If sections.Exists(NormalizeFilePath(CStr(rec("Path")))) Then sectionValue = CStr(sections(NormalizeFilePath(CStr(rec("Path")))))
        ws.Cells(r, 8).Value = sectionValue

        If IsSpecialDepartmentFile(CStr(rec("FileName")), CStr(rec("Title"))) Then
            ws.Cells(r, 5).Value = "—"
        End If

        sortKey = UCase$(CStr(rec("Object"))) & "|" & IIf(CBool(rec("IsReport")), "0", "1") & "|" & _
                  Format$(Val(CStr(rec("AppNumber"))), "000000") & "|" & UCase$(CStr(rec("Title")))
        ws.Cells(r, TEMP_SORT_COL).Value = sortKey
        r = r + 1
    Next i

    If r > FIRST_DATA_ROW Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range(ws.Cells(FIRST_DATA_ROW, TEMP_SORT_COL), ws.Cells(r - 1, TEMP_SORT_COL)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(r - 1, TEMP_SORT_COL))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
        For i = FIRST_DATA_ROW To r - 1
            ws.Cells(i, 1).Value = i - FIRST_DATA_ROW + 1
        Next i
    End If

    ApplyTableFormatting ws, r - 1
End Sub

Private Sub CheckDuplicateApplicationNumbers(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim objectName As String
    Dim appNumber As String
    Dim key As String
    Dim counts As Object
    Dim duplicateCount As Long

    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then Exit Sub

    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare

    ' Count each application number separately within each object.
    For r = FIRST_DATA_ROW To lastRow
        objectName = Trim$(CStr(ws.Cells(r, 2).Value))
        appNumber = Trim$(CStr(ws.Cells(r, 3).Value))
        If Len(objectName) > 0 And Len(appNumber) > 0 Then
            key = UCase$(objectName) & "|" & NormalizeApplicationNumberForDuplicateCheck(appNumber)
            If counts.Exists(key) Then
                counts(key) = CLng(counts(key)) + 1
            Else
                counts.Add key, 1
            End If
        End If
    Next r

    ' Highlight every row participating in a duplicate, and record it in the run log.
    For r = FIRST_DATA_ROW To lastRow
        objectName = Trim$(CStr(ws.Cells(r, 2).Value))
        appNumber = Trim$(CStr(ws.Cells(r, 3).Value))
        If Len(objectName) > 0 And Len(appNumber) > 0 Then
            key = UCase$(objectName) & "|" & NormalizeApplicationNumberForDuplicateCheck(appNumber)
            If counts.Exists(key) Then
                If CLng(counts(key)) > 1 Then
                    ws.Cells(r, 2).Interior.Color = RGB(255, 199, 206)
                    ws.Cells(r, 3).Interior.Color = RGB(255, 199, 206)
                    ws.Cells(r, 2).Font.Color = RGB(156, 0, 6)
                    ws.Cells(r, 3).Font.Color = RGB(156, 0, 6)
                    LogLine "Повтор номера приложения № " & appNumber & _
                            " у объекта '" & objectName & "' (строка " & r & ")."
                    duplicateCount = duplicateCount + 1
                End If
            End If
        End If
    Next r

    If duplicateCount = 0 Then
        LogLine "Повторов номеров приложений внутри объектов не обнаружено."
    Else
        LogLine "Проверка нумерации: отмечено строк с повторяющимися номерами: " & duplicateCount & "."
    End If
End Sub

Private Function NormalizeApplicationNumberForDuplicateCheck(ByVal appNumber As String) As String
    Dim value As String
    value = UCase$(Trim$(appNumber))

    ' Leading zeros do not make a different application number (01 = 1, 001A = 1A).
    ' Preserve a trailing letter suffix because 1A and 1B are distinct numbers.
    Do While Len(value) > 1 And Left$(value, 1) = "0" And _
             Mid$(value, 2, 1) >= "0" And Mid$(value, 2, 1) <= "9"
        value = Mid$(value, 2)
    Loop

    NormalizeApplicationNumberForDuplicateCheck = value
End Function

Private Sub ApplyLookups(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim appName As String

    If Not WorksheetExists(DATA_SHEET_NAME) Then
        LogLine "Предупреждение: лист 'data' не найден; столбцы 'Отдел' и 'Ответственное лицо' не заполнены автоматически."
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    For r = FIRST_DATA_ROW To lastRow
        appName = Trim$(CStr(ws.Cells(r, 4).Value))
        ' The exact source filename was checked while building the row. Preserve
        ' its special dash here; do not infer special status from a generic title.
        If Trim$(CStr(ws.Cells(r, 5).Value)) <> "—" Then
            ws.Cells(r, 5).FormulaLocal = "=ЕСЛИОШИБКА(ИНДЕКС(data!$B:$B;ПОИСКПОЗ($D" & r & ";data!$A:$A;0));"""")"
        End If
        ws.Cells(r, 6).FormulaLocal = "=ЕСЛИОШИБКА(ИНДЕКС(data!$C:$F;ПОИСКПОЗ($D" & r & ";data!$A:$A;0);ПОИСКПОЗ($B" & r & ";data!$C$1:$F$1;0));"""")"
    Next r
End Sub

Private Function IsSpecialDepartmentFile(ByVal fileNameOrNumber As String, ByVal titleText As String) As Boolean
    Dim fso As Object
    Dim baseName As String
    Dim normalizedFile As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    baseName = fso.GetBaseName(fileNameOrNumber)
    normalizedFile = LCase$(Trim$(Replace(baseName, "ё", "е")))

    ' Match only exact base filenames, never a generic parsed title.
    IsSpecialDepartmentFile = (normalizedFile = "отчет_2 кв 2026" Or _
                               normalizedFile = "реестр рисков")
End Function

Private Sub ApplyTableFormatting(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim fullRange As Range
    Dim widths As Variant
    Dim i As Long

    If lastRow < HEADER_ROW Then lastRow = HEADER_ROW
    Set fullRange = ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow, TEMP_SORT_COL))
    With fullRange
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, 10))
        .Interior.Color = RGB(198, 224, 180)
        .Font.Color = RGB(31, 78, 45)
        .Font.Italic = True
        .Font.Bold = False
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .RowHeight = 32
    End With

    widths = Array(6, 12, 8, 11, 16, 22, 16, 11, 10, 10)
    For i = LBound(widths) To UBound(widths)
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i
    ws.Columns(TEMP_SORT_COL).Hidden = True
    ws.Columns(1).HorizontalAlignment = xlCenter
    ws.Columns(3).HorizontalAlignment = xlCenter
    ws.Columns(8).HorizontalAlignment = xlCenter
    ws.Columns(9).HorizontalAlignment = xlCenter
    ws.Columns(10).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(lastRow, 10)).RowHeight = 20
    ws.Columns(4).WrapText = True
    ws.Columns(6).WrapText = True
    ws.Columns(7).WrapText = True
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = Nothing
    If WorksheetExists(sheetName) Then
        Set ws = FindWorksheet(sheetName)
    End If
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = FindWorksheet(sheetName)
    WorksheetExists = Not ws Is Nothing
End Function

Private Function FindWorksheet(ByVal sheetName As String) As Worksheet
    Dim foundSheet As Worksheet
    If Len(Trim$(sheetName)) = 0 Then Exit Function
    On Error Resume Next
    Set foundSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    Set FindWorksheet = foundSheet
End Function

Private Function SafeSheetName(ByVal value As String) As String
    Dim invalidChars As Variant
    Dim i As Long
    invalidChars = Array(":", "\", "/", "?", "*", "[", "]")
    For i = LBound(invalidChars) To UBound(invalidChars)
        value = Replace(value, CStr(invalidChars(i)), "_")
    Next i
    value = Trim$(value)
    If Len(value) = 0 Then value = "Перечень приложений"
    If Len(value) > 31 Then value = Left$(value, 31)
    SafeSheetName = value
End Function

' FolderPicker can return an HTTP(S) SharePoint URL instead of a Windows path.
' Convert it to the Windows WebDAV UNC form before calling FileSystemObject.
' For example:
' http://server:55555/SD/Folder -> \\server@55555\DavWWWRoot\SD\Folder
Private Function NormalizeSelectedRootPath(ByVal selectedPath As String) As String
    Dim schemeEnd As Long
    Dim slashAfterHost As Long
    Dim scheme As String
    Dim authority As String
    Dim urlPath As String
    Dim serverName As String
    Dim portText As String
    Dim colonPos As Long
    Dim webDavServer As String

    selectedPath = Trim$(selectedPath)
    schemeEnd = InStr(1, selectedPath, "://", vbTextCompare)

    If schemeEnd = 0 Then
        NormalizeSelectedRootPath = selectedPath
        Exit Function
    End If

    scheme = LCase$(Left$(selectedPath, schemeEnd - 1))
    If scheme <> "http" And scheme <> "https" Then
        NormalizeSelectedRootPath = selectedPath
        Exit Function
    End If

    slashAfterHost = InStr(schemeEnd + 3, selectedPath, "/")
    If slashAfterHost = 0 Then
        authority = Mid$(selectedPath, schemeEnd + 3)
        urlPath = ""
    Else
        authority = Mid$(selectedPath, schemeEnd + 3, slashAfterHost - (schemeEnd + 3))
        urlPath = Mid$(selectedPath, slashAfterHost + 1)
    End If

    ' Remove query strings or fragments if the picker supplied them.
    If InStr(urlPath, "?") > 0 Then urlPath = Left$(urlPath, InStr(urlPath, "?") - 1)
    If InStr(urlPath, "#") > 0 Then urlPath = Left$(urlPath, InStr(urlPath, "#") - 1)
    urlPath = Replace(urlPath, "%20", " ")
    urlPath = Replace(urlPath, "/", "\")

    serverName = authority
    portText = ""
    colonPos = InStrRev(authority, ":")
    If colonPos > 0 Then
        serverName = Left$(authority, colonPos - 1)
        portText = Mid$(authority, colonPos + 1)
    End If

    If Len(serverName) = 0 Then
        NormalizeSelectedRootPath = selectedPath
        Exit Function
    End If

    webDavServer = serverName
    If scheme = "https" Then
        webDavServer = webDavServer & "@SSL"
        If Len(portText) > 0 And portText <> "443" Then webDavServer = webDavServer & "@" & portText
    Else
        If Len(portText) > 0 And portText <> "80" Then webDavServer = webDavServer & "@" & portText
    End If

    NormalizeSelectedRootPath = "\\" & webDavServer & "\DavWWWRoot"
    If Len(urlPath) > 0 Then NormalizeSelectedRootPath = NormalizeSelectedRootPath & "\" & urlPath

    LogLine "Адрес SharePoint преобразован в путь Windows WebDAV: " & NormalizeSelectedRootPath
End Function

Private Function GetFolderName(ByVal folderPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFolderName = fso.GetFolder(RemoveTrailingSlash(folderPath)).Name
End Function

Private Function RemoveTrailingSlash(ByVal value As String) As String
    Do While Len(value) > 3 And (Right$(value, 1) = "\" Or Right$(value, 1) = "/")
        value = Left$(value, Len(value) - 1)
    Loop
    RemoveTrailingSlash = value
End Function

Private Sub LogLine(ByVal message As String)
    mLogText = mLogText & Format$(Now, "hh:nn:ss") & "  " & message & vbCrLf
    Application.StatusBar = message
    DoEvents
End Sub
