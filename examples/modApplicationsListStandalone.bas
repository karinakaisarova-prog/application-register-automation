Attribute VB_Name = "modApplicationsListStandalone"
Option Explicit

' Автоматизация перечня приложений
' Основной макрос формирования перечня в настольном Excel под Windows.
' Отдельная версия: только макрос формирования перечня; установщик и другие модули VBA не требуются.
'
' Основная процедура формирования перечня: CreateMultiFolderApplicationsListCorrected
' Глубина поиска: до MAX_FOLDER_DEPTH уровней внутри каждой выбранной папки.
' Ручные значения столбца H («Раздел») сохраняются по полному пути к файлу.
' Закрепление областей не используется.

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
Private mCurrentRootIsWebDav As Boolean
Private mFolderErrorCount As Long

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
    mFolderErrorCount = 0

    Set roots = New Collection
    If MsgBox("Выберите способ указания исходной папки:" & vbCrLf & vbCrLf & _
              "Да — выбрать одну или несколько папок через Проводник Windows." & vbCrLf & _
              "Нет — вставить веб-адрес папки SharePoint." & vbCrLf & _
              "Отмена — выйти без изменений.", _
              vbYesNoCancel + vbQuestion + vbDefaultButton1, "Выбор исходной папки") = vbYes Then
        Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
        With dlg
            .Title = "Выберите одну или несколько исходных папок"
            .AllowMultiSelect = True
            If .Show <> -1 Then GoTo CleanExit
            For Each rootItem In .SelectedItems
                roots.Add NormalizeSelectedRootPath(CStr(rootItem))
            Next rootItem
        End With
    Else
        Dim sharePointUrl As String
        sharePointUrl = Trim$(InputBox("Вставьте полный адрес папки SharePoint, начиная с http:// или https://", _
                                       "Адрес папки SharePoint"))
        If Len(sharePointUrl) = 0 Then GoTo CleanExit
        If Len(sharePointUrl) >= 254 Then
            MsgBox "Адрес слишком длинный: окно ввода может обрезать его." & vbCrLf & _
                   "Выберите синхронизированную папку SharePoint через Проводник Windows.", _
                   vbExclamation, "Слишком длинный адрес"
            GoTo CleanExit
        End If
        If LCase$(Left$(sharePointUrl, 7)) <> "http://" And _
           LCase$(Left$(sharePointUrl, 8)) <> "https://" Then
            MsgBox "Укажите полный веб-адрес, начинающийся с http:// или https://.", vbExclamation, "Неверный адрес"
            GoTo CleanExit
        End If
        roots.Add NormalizeSelectedRootPath(sharePointUrl)
    End If

    sheetName = SafeSheetName(GetFolderName(CStr(roots(1))))
    If LCase$(sheetName) = LCase$(DATA_SHEET_NAME) Then sheetName = "Перечень приложений"

    Set mSections = CreateObject("Scripting.Dictionary")
    mSections.CompareMode = vbTextCompare
    Set ws = Nothing
    If WorksheetExists(sheetName) Then
        Set ws = FindWorksheet(sheetName)
    End If
    If Not ws Is Nothing Then
        If Not IsRegisterSheet(ws) Then
            LogLine "Лист «" & ws.Name & "» не является перечнем и не изменён."
            Set ws = Nothing
            sheetName = GetUniqueListSheetName(sheetName)
        End If
    End If
    targetWasExisting = Not ws Is Nothing
    originalSheetName = sheetName
    If targetWasExisting Then
        SaveExistingSections ws, mSections
        ' Перед перестроением создаёт резервную копию листа для восстановления.
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
        mCurrentRootIsWebDav = (InStr(1, mRootPath, "\\DavWWWRoot", vbTextCompare) > 0)
        LogLine "Поиск файлов: " & mRootPath
        CollectFiles mRootPath, mRootPath, 0
    Next rootItem

    If mFolderErrorCount > 0 Then
        If MsgBox("При чтении папок возникло ошибок: " & mFolderErrorCount & "." & vbCrLf & _
                  "Если продолжить, перечень может оказаться неполным." & vbCrLf & vbCrLf & _
                  "Для SharePoint выбирайте папку через синхронизированную папку OneDrive в Проводнике Windows, " & _
                  "а не веб-адрес в браузере." & vbCrLf & vbCrLf & _
                  "Продолжить с найденными файлами?", _
                  vbYesNo + vbExclamation + vbDefaultButton2, "Обнаружены недоступные папки") <> vbYes Then
            If Not backupWs Is Nothing Then
                backupWs.Delete
                Set backupWs = Nothing
                LogLine "Исходный лист не изменён; временная резервная копия удалена."
            End If
            LogLine "Построение остановлено: пользователь не подтвердил создание неполного перечня."
            GoTo CleanExit
        End If
    End If

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
            ' Исходный лист не менялся; удаляет только ненужную копию.
            Err.Clear
            backupWs.Delete
            If Err.Number = 0 Then
                Set backupWs = Nothing
                LogLine "Формирование не началось; исходный лист оставлен без изменений."
            Else
                Err.Clear
                LogLine "Не удалось удалить временную резервную копию; она сохранена для восстановления."
            End If
        Else
            ' Не удаляет повреждённый лист, пока резервная копия не получит его исходное имя.
            ' Если переименование не удалось, сохраняет резервную копию
            ' и оставляет доступные листы для восстановления.
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
        If Not ws Is Nothing Then
            Err.Clear
            ws.Delete
            If Err.Number = 0 Then
                Set ws = Nothing
                LogLine "Новый неполный лист «" & sheetName & "» удалён."
            Else
                LogLine "Не удалось удалить новый неполный лист «" & sheetName & "»: " & Err.Description
                Err.Clear
            End If
        Else
            LogLine "Не удалось подтвердить удаление нового неполного листа «" & sheetName & "»: ссылка на лист недоступна."
        End If
    End If
    Application.StatusBar = False
    MsgBox IIf(cancelOrError, "Операция отменена.", "Не удалось завершить формирование перечня: " & errText) & _
           vbCrLf & vbCrLf & mLogText, vbExclamation, "Перечень приложений"

CleanExit:
    On Error Resume Next
    If Not backupWs Is Nothing Then
        ' После неудачного восстановления резервная копия может остаться единственной доступной копией.
        ' Сохраняет её, не полагаясь на Err.Number, который мог быть сброшен.
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
        If Not IsSkippedFile(fileItem) Then
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
        End If
    Next fileItem

    If depth < MAX_FOLDER_DEPTH Then
        For Each subFolder In folder.SubFolders
            DoEvents
            Application.StatusBar = "Обработка папки: " & CStr(subFolder.Path)
            CollectFiles rootPath, CStr(subFolder.Path), depth + 1
        Next subFolder
    ElseIf folder.SubFolders.Count > 0 Then
        mFolderErrorCount = mFolderErrorCount + 1
        LogLine "НЕ ПРОСМОТРЕНЫ вложенные папки глубже " & MAX_FOLDER_DEPTH & " уровней: " & currentPath
    End If
    Exit Sub

FolderError:
    If Err.Number = 18 Then
        ' Не перехватываем Ctrl+Break: обработчик верхнего уровня восстановит настройки Excel.
        Err.Raise 18
    End If
    mFolderErrorCount = mFolderErrorCount + 1
    LogLine "ОШИБКА ЧТЕНИЯ ПАПКИ: " & currentPath & " (ошибка " & Err.Number & ": " & Err.Description & ")"
    If mCurrentRootIsWebDav Then
        LogLine "Подсказка: это WebDAV-путь. Проверьте службу WebClient Windows, подключение к серверу, авторизацию и права доступа. Если путь не открывается в Проводнике, выберите синхронизированную папку SharePoint."
    ElseIf Left$(currentPath, 2) = "\\" Then
        LogLine "Подсказка: проверьте доступ к сетевой папке и права чтения."
    Else
        LogLine "Подсказка: проверьте, существует ли папка, синхронизирована ли она с OneDrive и доступны ли файлы локально."
    End If
    Err.Clear
End Sub

Private Function IsSkippedFile(ByVal fileItem As Object) As Boolean
    Dim registerPath As String

    If Left$(CStr(fileItem.Name), 2) = "~$" Then
        IsSkippedFile = True
        Exit Function
    End If

    If (CLng(fileItem.Attributes) And 2) <> 0 Then
        IsSkippedFile = True
        Exit Function
    End If

    On Error Resume Next
    registerPath = ThisWorkbook.FullName
    If Err.Number = 18 Then
        On Error GoTo 0
        Err.Raise 18
    End If
    Err.Clear
    On Error GoTo 0

    If Len(registerPath) > 0 Then
        IsSkippedFile = (StrComp(CStr(fileItem.Path), registerPath, vbTextCompare) = 0)
    End If
End Function

Private Function IsRegisterSheet(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    IsRegisterSheet = (Trim$(CStr(ws.Cells(2, 1).Text)) = "№ п/п" And _
                       Trim$(CStr(ws.Cells(2, 6).Text)) = "Ответственное лицо" And _
                       Trim$(CStr(ws.Cells(2, 8).Text)) = "Раздел")
End Function

Private Function GetUniqueListSheetName(ByVal baseName As String) As String
    Dim candidate As String
    Dim suffix As Long

    suffix = 2
    Do
        candidate = Left$(baseName, 31 - Len(" (" & suffix & ")")) & " (" & suffix & ")"
        If Not WorksheetExists(candidate) Then Exit Do
        suffix = suffix + 1
    Loop
    GetUniqueListSheetName = candidate
End Function

Private Function IsSupportedExtension(ByVal ext As String) As Boolean
    IsSupportedExtension = (ext = "doc" Or ext = "docx" Or ext = "docm" Or ext = "pdf" Or _
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

    ' Проверяет только выбранную папку и вложенные в неё папки.
    ' Родительская папка за пределами выбранной не должна определять объект.
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

        ' После проверки выбранной папки не проверяет родительские папки.
        If StrComp(folderPath, selectedRootPath, vbTextCompare) = 0 Then Exit Do
        parentPath = fso.GetParentFolderName(folderPath)
        If Len(parentPath) = 0 Or StrComp(parentPath, folderPath, vbTextCompare) = 0 Then Exit Do
        folderPath = RemoveTrailingSlash(parentPath)
    Loop

    ' Если в пути выбранной папки не найден подходящий объект, проверяет имя файла.
    ' Если объект определить нельзя, оставляет значение пустым, а не угадывает.
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
        ' Ищет полное обозначение или название, а не часть другого слова.
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

    If ws Is Nothing Then Exit Sub
    If target Is Nothing Then Exit Sub
    On Error GoTo ReadFailed
    lastRow = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row
    If lastRow < FIRST_DATA_ROW Then Exit Sub

    For r = FIRST_DATA_ROW To lastRow
        On Error GoTo RowFailed
        filePath = GetCellLinkPath(ws.Cells(r, 7))
        sectionValue = CStr(ws.Cells(r, 8).Value)
        If Len(filePath) > 0 And Len(sectionValue) > 0 Then
            target(NormalizeFilePath(filePath)) = sectionValue
        End If
NextRow:
        On Error GoTo 0
    Next r
    Exit Sub

RowFailed:
    If Err.Number = 18 Then
        On Error GoTo 0
        Err.Raise 18
    End If
    LogLine "Не удалось сохранить значение столбца «Раздел», строка " & r & ": " & Err.Description
    Err.Clear
    Resume NextRow

ReadFailed:
    If Err.Number = 18 Then
        On Error GoTo 0
        Err.Raise 18
    End If
    LogLine "Не удалось прочитать существующие значения столбца «Раздел»: " & Err.Description
    Err.Clear
End Sub

Private Function GetCellLinkPath(ByVal c As Range) As String
    On Error GoTo ReadFailed
    If c.Hyperlinks.Count > 0 Then
        GetCellLinkPath = ResolveLinkPath(Trim$(CStr(c.Hyperlinks(1).Address)), c.Worksheet.Parent)
    Else
        GetCellLinkPath = Trim$(CStr(c.Value))
    End If
    Exit Function

ReadFailed:
    If Err.Number = 18 Then
        On Error GoTo 0
        Err.Raise 18
    End If
    GetCellLinkPath = ""
End Function

Private Function ResolveLinkPath(ByVal addr As String, ByVal wb As Workbook) As String
    Dim fso As Object
    Dim basePath As String
    Dim workbookPath As String
    Dim operationError As Long
    Dim baseIsAbsolute As Boolean

    addr = Trim$(addr)
    ResolveLinkPath = addr
    If Len(addr) = 0 Then Exit Function
    If addr Like "[A-Za-z]:*" Then Exit Function
    If Left$(addr, 2) = "\\" Then Exit Function
    If Left$(addr, 2) = "//" Then Exit Function
    If InStr(1, addr, "://", vbTextCompare) > 0 Then Exit Function
    If LCase$(Left$(addr, 5)) = "file:" Then Exit Function

    On Error Resume Next
    Err.Clear
    basePath = CStr(wb.BuiltinDocumentProperties("Hyperlink base").Value)
    operationError = Err.Number
    If operationError = 18 Then
        On Error GoTo 0
        Err.Raise 18
    End If
    If operationError <> 0 Then
        Err.Clear
        basePath = ""
    End If

    If Len(basePath) = 0 Then
        Err.Clear
        basePath = wb.Path
        operationError = Err.Number
        If operationError = 18 Then
            On Error GoTo 0
            Err.Raise 18
        End If
        If operationError <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
    End If

    If Len(basePath) = 0 Or InStr(1, basePath, "://", vbTextCompare) > 0 Then
        On Error GoTo 0
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    operationError = Err.Number
    If operationError = 18 Then
        On Error GoTo 0
        Err.Raise 18
    End If
    If operationError <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    basePath = Replace(basePath, "/", "\")
    baseIsAbsolute = False
    If Len(basePath) >= 3 Then
        baseIsAbsolute = (Mid$(basePath, 2, 1) = ":" And Mid$(basePath, 3, 1) = "\")
    End If
    If Left$(basePath, 2) = "\\" Then baseIsAbsolute = True

    If Not baseIsAbsolute Then
        If (Len(basePath) >= 2 And Mid$(basePath, 2, 1) = ":") Or Left$(basePath, 1) = "\" Then
            On Error GoTo 0
            Exit Function
        End If

        Err.Clear
        workbookPath = CStr(wb.Path)
        operationError = Err.Number
        If operationError = 18 Then
            On Error GoTo 0
            Err.Raise 18
        End If
        If operationError <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If

        workbookPath = Replace(workbookPath, "/", "\")
        If Not ((Len(workbookPath) >= 3 And Mid$(workbookPath, 2, 1) = ":" And Mid$(workbookPath, 3, 1) = "\") Or _
                Left$(workbookPath, 2) = "\\") Then
            On Error GoTo 0
            Exit Function
        End If

        Err.Clear
        basePath = fso.GetAbsolutePathName(fso.BuildPath(workbookPath, basePath))
        operationError = Err.Number
        If operationError = 18 Then
            On Error GoTo 0
            Err.Raise 18
        End If
        If operationError <> 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
    End If

    Err.Clear
    ResolveLinkPath = fso.GetAbsolutePathName(fso.BuildPath(basePath, Replace(addr, "/", "\")))
    operationError = Err.Number
    If operationError = 18 Then
        On Error GoTo 0
        Err.Raise 18
    End If
    If operationError <> 0 Then
        Err.Clear
        ResolveLinkPath = addr
    End If
    On Error GoTo 0
End Function

Private Function NormalizeFilePath(ByVal value As String) As String
    value = Trim$(value)
    If LCase$(Left$(value, 7)) = "file://" Then
        value = Mid$(value, 8)
        If Left$(value, 1) = "/" Then
            value = Mid$(value, 2)
        Else
            value = "\\" & value
        End If
    End If
    value = DecodeUrlPath(value)
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
    Dim hyperlinkDeleteError As Long

    headers = Array("№ п/п", "Объект", "№ Прил.", "Наименование", "Отдел", _
                    "Ответственное лицо", "Ссылка на файл", "Раздел", "Стр. отчёта", "Кол-во страниц")

    ws.Cells.Clear
    On Error Resume Next
    Err.Clear
    ws.Hyperlinks.Delete
    hyperlinkDeleteError = Err.Number
    On Error GoTo 0
    If hyperlinkDeleteError = 18 Then Err.Raise 18

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

    ' Считает номера приложений отдельно для каждого объекта.
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

    ' Выделяет все строки с повторяющимися номерами и записывает их в журнал запуска.
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

    ' Начальные нули не создают новый номер: 01 считается номером 1, а 001A — номером 1A.
    ' Сохраняет буквенный суффикс, поскольку 1A и 1B — разные номера.
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
        ' Точное имя исходного файла проверено при формировании строки. Сохраняет
        ' специальное тире и не определяет его только по общему названию.
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

    ' Сопоставляет только точные базовые имена файлов, а не общее разобранное название.
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
    Dim lookupError As Long
    On Error Resume Next
    Err.Clear
    Set ws = FindWorksheet(sheetName)
    lookupError = Err.Number
    On Error GoTo 0
    If lookupError = 18 Then Err.Raise 18
    WorksheetExists = Not ws Is Nothing
End Function

Private Function FindWorksheet(ByVal sheetName As String) As Worksheet
    Dim foundSheet As Worksheet
    Dim lookupError As Long
    If Len(Trim$(sheetName)) = 0 Then Exit Function
    On Error Resume Next
    Err.Clear
    Set foundSheet = ThisWorkbook.Worksheets(sheetName)
    lookupError = Err.Number
    On Error GoTo 0
    If lookupError = 18 Then Err.Raise 18
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

' Окно выбора папки может вернуть веб-адрес SharePoint вместо пути Windows.
' Перед обращением к файловой системе преобразует адрес в сетевой путь WebDAV Windows.
' Например:
' Пример преобразования веб-адреса папки в сетевой путь WebDAV Windows.
Private Function NormalizeSelectedRootPath(ByVal selectedPath As String) As String
    Dim schemeEnd As Long
    Dim authorityEnd As Long
    Dim scheme As String
    Dim authority As String
    Dim urlPath As String
    Dim serverName As String
    Dim portText As String
    Dim colonPos As Long
    Dim portNumber As Long
    Dim webDavServer As String
    Dim delimiter As String
    Dim i As Long

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

    For i = schemeEnd + 3 To Len(selectedPath)
        delimiter = Mid$(selectedPath, i, 1)
        If delimiter = "/" Or delimiter = "?" Or delimiter = "#" Then
            authorityEnd = i
            Exit For
        End If
    Next i

    If authorityEnd = 0 Then
        authority = Mid$(selectedPath, schemeEnd + 3)
        urlPath = ""
    Else
        authority = Mid$(selectedPath, schemeEnd + 3, authorityEnd - (schemeEnd + 3))
        If delimiter = "/" Then
            urlPath = Mid$(selectedPath, authorityEnd + 1)
        Else
            urlPath = ""
        End If
    End If

    If InStr(urlPath, "?") > 0 Then urlPath = Left$(urlPath, InStr(urlPath, "?") - 1)
    If InStr(urlPath, "#") > 0 Then urlPath = Left$(urlPath, InStr(urlPath, "#") - 1)
    urlPath = DecodeUrlPath(urlPath)
    urlPath = Replace(urlPath, "/", "\")

    If Len(authority) = 0 Then
        Err.Raise vbObjectError + 2101, "NormalizeSelectedRootPath", _
                  "В адресе SharePoint не указано имя сервера."
    End If

    If InStr(1, authority, "[", vbBinaryCompare) > 0 Or _
       InStr(1, authority, "]", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 2102, "NormalizeSelectedRootPath", _
                  "Адреса IPv6 пока не поддерживаются при преобразовании SharePoint в WebDAV."
    End If

    serverName = authority
    portText = ""
    colonPos = InStr(1, authority, ":", vbBinaryCompare)
    If colonPos > 0 Then
        If InStr(colonPos + 1, authority, ":", vbBinaryCompare) > 0 Then
            Err.Raise vbObjectError + 2103, "NormalizeSelectedRootPath", _
                      "В адресе SharePoint обнаружено несколько двоеточий. Проверьте имя сервера и порт."
        End If

        serverName = Left$(authority, colonPos - 1)
        portText = Mid$(authority, colonPos + 1)

        If Len(serverName) = 0 Then
            Err.Raise vbObjectError + 2104, "NormalizeSelectedRootPath", _
                      "В адресе SharePoint не указано имя сервера."
        End If

        If Len(portText) = 0 Then
            Err.Raise vbObjectError + 2105, "NormalizeSelectedRootPath", _
                      "После двоеточия в адресе SharePoint не указан порт."
        End If

        If Len(portText) > 5 Then
            Err.Raise vbObjectError + 2106, "NormalizeSelectedRootPath", _
                      "Порт SharePoint должен быть целым числом от 1 до 65535."
        End If

        For i = 1 To Len(portText)
            If AscW(Mid$(portText, i, 1)) < 48 Or AscW(Mid$(portText, i, 1)) > 57 Then
                Err.Raise vbObjectError + 2107, "NormalizeSelectedRootPath", _
                          "Порт SharePoint должен содержать только цифры."
            End If
        Next i

        portNumber = CLng(portText)
        If portNumber < 1 Or portNumber > 65535 Then
            Err.Raise vbObjectError + 2108, "NormalizeSelectedRootPath", _
                      "Порт SharePoint должен быть целым числом от 1 до 65535."
        End If
        portText = CStr(portNumber)
    End If

    If Len(serverName) = 0 Then
        Err.Raise vbObjectError + 2109, "NormalizeSelectedRootPath", _
                  "В адресе SharePoint не указано имя сервера."
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

Private Function DecodeUrlPath(ByVal value As String) As String
    Dim result As String
    Dim i As Long
    Dim b1 As Long
    Dim b2 As Long
    Dim b3 As Long
    Dim b4 As Long
    Dim codePoint As Long
    Dim highSurrogate As Long
    Dim lowSurrogate As Long
    Dim sequenceLength As Long
    Dim decoded As Boolean

    i = 1
    Do While i <= Len(value)
        decoded = False
        If TryReadEncodedByte(value, i, b1) Then
            If b1 < 128 Then
                result = result & ChrW$(b1)
                sequenceLength = 3
                decoded = True
            ElseIf b1 >= 194 And b1 <= 223 Then
                If TryReadEncodedByte(value, i + 3, b2) Then
                    If b2 >= 128 And b2 <= 191 Then
                        codePoint = (b1 And 31) * 64 + (b2 And 63)
                        result = result & ChrW$(codePoint)
                        sequenceLength = 6
                        decoded = True
                    End If
                End If
            ElseIf b1 >= 224 And b1 <= 239 Then
                If TryReadEncodedByte(value, i + 3, b2) And _
                   TryReadEncodedByte(value, i + 6, b3) Then
                    If b2 >= 128 And b2 <= 191 And b3 >= 128 And b3 <= 191 Then
                        codePoint = (b1 And 15) * 4096 + (b2 And 63) * 64 + (b3 And 63)
                        If codePoint >= 2048 And _
                           (codePoint < 55296 Or codePoint > 57343) Then
                            result = result & ChrW$(codePoint)
                            sequenceLength = 9
                            decoded = True
                        End If
                    End If
                End If
            ElseIf b1 >= 240 And b1 <= 244 Then
                If TryReadEncodedByte(value, i + 3, b2) And _
                   TryReadEncodedByte(value, i + 6, b3) And _
                   TryReadEncodedByte(value, i + 9, b4) Then
                    If b2 >= 128 And b2 <= 191 And _
                       b3 >= 128 And b3 <= 191 And _
                       b4 >= 128 And b4 <= 191 Then
                        codePoint = (b1 And 7) * 262144 + _
                                    (b2 And 63) * 4096 + _
                                    (b3 And 63) * 64 + _
                                    (b4 And 63)
                        If codePoint >= 65536 And codePoint <= 1114111 Then
                            codePoint = codePoint - 65536
                            highSurrogate = 55296 + (codePoint \ 1024)
                            lowSurrogate = 56320 + (codePoint Mod 1024)
                            result = result & ChrW$(highSurrogate) & ChrW$(lowSurrogate)
                            sequenceLength = 12
                            decoded = True
                        End If
                    End If
                End If
            End If
        End If

        If decoded Then
            i = i + sequenceLength
        ElseIf Mid$(value, i, 1) = "%" And i + 2 <= Len(value) Then
            result = result & Mid$(value, i, 3)
            i = i + 3
        Else
            result = result & Mid$(value, i, 1)
            i = i + 1
        End If
    Loop

    DecodeUrlPath = result
End Function

Private Function TryReadEncodedByte(ByVal value As String, ByVal position As Long, ByRef byteValue As Long) As Boolean
    Dim hexPair As String

    If position < 1 Or position + 2 > Len(value) Then Exit Function
    If Mid$(value, position, 1) <> "%" Then Exit Function

    hexPair = Mid$(value, position + 1, 2)
    If InStr(1, "0123456789ABCDEF", UCase$(Left$(hexPair, 1)), vbBinaryCompare) = 0 Then Exit Function
    If InStr(1, "0123456789ABCDEF", UCase$(Right$(hexPair, 1)), vbBinaryCompare) = 0 Then Exit Function

    byteValue = CLng("&H" & hexPair)
    TryReadEncodedByte = True
End Function

Private Function GetFolderName(ByVal folderPath As String) As String
    Dim fso As Object
    Dim normalizedPath As String
    Dim separatorPos As Long
    Dim accessError As Long

    normalizedPath = RemoveTrailingSlash(Trim$(folderPath))
    On Error Resume Next
    Err.Clear
    Set fso = CreateObject("Scripting.FileSystemObject")
    accessError = Err.Number
    If accessError = 0 Then
        If Not fso Is Nothing Then
            GetFolderName = fso.GetFolder(normalizedPath).Name
            accessError = Err.Number
        End If
    End If
    On Error GoTo 0
    If accessError = 18 Then Err.Raise 18

    ' Сетевой путь WebDAV может быть недоступен уже при определении имени листа.
    ' Использует последний фрагмент пути, чтобы последующая проверка записала реальную ошибку доступа.
    If Len(GetFolderName) = 0 Then
        normalizedPath = Replace(normalizedPath, "/", "\")
        separatorPos = InStrRev(normalizedPath, "\")
        If separatorPos > 0 And separatorPos < Len(normalizedPath) Then
            GetFolderName = Mid$(normalizedPath, separatorPos + 1)
        Else
            GetFolderName = normalizedPath
        End If
    End If
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
