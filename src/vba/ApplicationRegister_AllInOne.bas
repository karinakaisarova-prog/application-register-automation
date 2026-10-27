Attribute VB_Name = "ApplicationRegister_AllInOne"
Option Explicit

' ===== Source module: modWorkbookContext.bas =====
' Книга с макросами, справочником data и перечнем — одна и та же книга.
' Запускатель работает в ThisWorkbook и не создаёт отдельную книгу результата.
Public gRegisterWorkbook As Workbook

Public Sub StartRegisterAutomation()
    Dim sourceWb As Workbook
    Dim sourceData As Worksheet
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo Failed

    Set sourceWb = ThisWorkbook
    On Error Resume Next
    Set sourceData = sourceWb.Worksheets("data")
    On Error GoTo Failed

    If sourceData Is Nothing Then
        MsgBox "В этой книге не найден лист с точным именем 'data'." & vbCrLf & _
               "Откройте подготовленную книгу, содержащую справочник data и макросы. " & _
               "Перечень будет создан в этой же книге.", _
               vbExclamation, "Нет справочника data"
        Exit Sub
    End If

    Set gRegisterWorkbook = sourceWb
    sourceWb.Activate

    MsgBox "Перечень будет создан в этой же книге, где находится лист data." & vbCrLf & _
           "Исходные документы останутся в выбранных папках; в Excel будут записаны ссылки на них.", _
           vbInformation, "Автоматизация перечня"

    CreateMultiFolderApplicationsListCorrected
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error GoTo 0
    MsgBox "Не удалось запустить формирование перечня." & vbCrLf & _
           "Ошибка " & failureNumber & ": " & failureDescription, _
           vbExclamation, "Автоматизация перечня"
End Sub

Public Function GetRegisterWorkbook() As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet

    ' Prefer the workbook explicitly selected by the startup procedure.
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

    ' The macro workbook is the intended home for the data sheet and register.
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("data")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set gRegisterWorkbook = ThisWorkbook
        Set GetRegisterWorkbook = ThisWorkbook
        Exit Function
    End If

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


' ===== Source module: modRegisterPanel.bas =====
' Панель управления макросами на листе перечня.
' Кнопки размещаются справа от таблицы, начиная с колонки L,
' чтобы не сдвигать строки перечня и не менять формулы.

Private Const RegisterPanel_PANEL_PREFIX As String = "regPanel_"
Private Const RegisterPanel_PANEL_LEFT_COLUMN As Long = 12
Private Const RegisterPanel_BUTTON_WIDTH As Double = 190
Private Const RegisterPanel_BUTTON_HEIGHT As Double = 30
Private Const RegisterPanel_BUTTON_GAP As Double = 6

Public Sub AddRegisterControlPanel()
    Dim ws As Worksheet

    Set ws = GetActiveRegisterSheet()
    If ws Is Nothing Then
        MsgBox "Не найден лист перечня в книге результата. Сначала сформируйте перечень и откройте его лист.", _
               vbExclamation, "Панель макросов"
        Exit Sub
    End If

    CreateRegisterControlPanel ws
    MsgBox "Панель макросов добавлена справа от таблицы. Если её не видно, прокрутите лист вправо.", _
           vbInformation, "Панель макросов"
End Sub

Public Sub CreateRegisterControlPanel(ByVal ws As Worksheet)
    Dim buttonNames As Variant
    Dim buttonCaptions As Variant
    Dim macroNames As Variant
    Dim i As Long
    Dim shp As Shape
    Dim leftPos As Double
    Dim topPos As Double
    Dim actionName As String
    Dim workbookRef As String

    If ws Is Nothing Then Exit Sub

    DeleteRegisterControlPanel ws

    buttonCaptions = Array( _
        "Повторно создать перечень", _
        "Анализ упоминаний в отчёте", _
        "Подсчитать страницы", _
        "Создать data_full", _
        "Переименовать файлы", _
        "Откатить переименование", _
        "Проверить нумерацию таблиц Word", _
        "Проверить маркированные списки Word", _
        "Открыть памятку", _
        "Обновить панель")

    macroNames = Array( _
        "CreateMultiFolderApplicationsListCorrected", _
        "ScanWordReportMentions", _
        "CountPagesForRegister", _
        "BuildDataFull", _
        "RenameFilesFromRegister", _
        "RollbackLastRenameBatch", _
        "CheckAndNumberWordTables", _
        "CheckAndFixWordBulletLists", _
        "OpenMacroLaunchGuide", _
        "AddRegisterControlPanel")

    workbookRef = Replace(ThisWorkbook.Name, "'", "''") & "'!"
    leftPos = ws.Columns(RegisterPanel_PANEL_LEFT_COLUMN).Left
    topPos = ws.Rows(1).Top

    For i = LBound(buttonCaptions) To UBound(buttonCaptions)
        Set shp = ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            leftPos, _
            topPos + i * (RegisterPanel_BUTTON_HEIGHT + RegisterPanel_BUTTON_GAP), _
            RegisterPanel_BUTTON_WIDTH, _
            RegisterPanel_BUTTON_HEIGHT)

        shp.Name = RegisterPanel_PANEL_PREFIX & CStr(i + 1)
        shp.TextFrame2.TextRange.Text = CStr(buttonCaptions(i))
        shp.TextFrame2.TextRange.Font.Name = "Arial"
        shp.TextFrame2.TextRange.Font.Size = 9
        shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
        shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        shp.Placement = xlFreeFloating
        shp.OnAction = "'" & workbookRef & CStr(macroNames(i))
    Next i
End Sub

Public Sub DeleteRegisterControlPanel(ByVal ws As Worksheet)
    Dim i As Long
    If ws Is Nothing Then Exit Sub

    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(RegisterPanel_PANEL_PREFIX)) = RegisterPanel_PANEL_PREFIX Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub

Public Sub OpenMacroLaunchGuide()
    On Error GoTo OpenError
    ThisWorkbook.FollowHyperlink Address:="https://github.com/karinakaisarova-prog/application-register-automation/blob/main/docs/macros-launch-order.md"
    Exit Sub
OpenError:
    MsgBox "Не удалось открыть памятку автоматически." & vbCrLf & _
           "Откройте ссылку вручную:" & vbCrLf & _
           "https://github.com/karinakaisarova-prog/application-register-automation/blob/main/docs/macros-launch-order.md", _
           vbExclamation, "Памятка по запуску"
End Sub


' ===== Source module: modApplicationsList.bas =====
' Application Register Automation
' Main list builder for desktop Excel on Windows.
' Requires no additional VBA references (FileSystemObject and Word are late-bound).
'
' Entry point: CreateMultiFolderApplicationsListCorrected
' Folder depth: up to ApplicationsList_MAX_FOLDER_DEPTH below each selected root.
' Manual values in column H ("Раздел") are preserved by full file path.
' No FreezePanes are used.

Private Const ApplicationsList_HEADER_ROW As Long = 2
Private Const ApplicationsList_FIRST_DATA_ROW As Long = 3
Private Const ApplicationsList_MAX_FOLDER_DEPTH As Long = 3
Private Const ApplicationsList_TEMP_SORT_COL As Long = 11
Private Const ApplicationsList_DATA_SHEET_NAME As String = "data"

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
            roots.Add CStr(rootItem)
        Next rootItem
    End With

    sheetName = SafeSheetName(GetFolderName(CStr(roots(1))))
    If LCase$(sheetName) = LCase$(ApplicationsList_DATA_SHEET_NAME) Then sheetName = "Перечень приложений"

    Set mSections = CreateObject("Scripting.Dictionary")
    mSections.CompareMode = vbTextCompare
    On Error Resume Next
    Set ws = GetRegisterWorkbook().Worksheets(sheetName)
    On Error GoTo Failed
    targetWasExisting = Not ws Is Nothing
    originalSheetName = sheetName
    If targetWasExisting Then
        SaveExistingSections ws, mSections
        ' Keep a worksheet-level recovery copy before any destructive rebuild.
        backupSheetName = GetUniqueBackupSheetName()
        ws.Copy After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count)
        Set backupWs = GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count)
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
    ws.Columns(ApplicationsList_TEMP_SORT_COL).Hidden = True
    ws.Activate
    CreateRegisterControlPanel ws
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
            Set damagedWs = GetRegisterWorkbook().Worksheets(originalSheetName)
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
        Set ws = GetRegisterWorkbook().Worksheets(sheetName)
        If Not ws Is Nothing Then ws.Delete
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

    If depth < ApplicationsList_MAX_FOLDER_DEPTH Then
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
    If lastRow < ApplicationsList_FIRST_DATA_ROW Then Exit Sub

    For r = ApplicationsList_FIRST_DATA_ROW To lastRow
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
        ws.Cells(ApplicationsList_HEADER_ROW, i + 1).Value = headers(i)
    Next i
    ws.Cells(ApplicationsList_HEADER_ROW, ApplicationsList_TEMP_SORT_COL).Value = "Временный ключ"

    r = ApplicationsList_FIRST_DATA_ROW
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
        ws.Cells(r, ApplicationsList_TEMP_SORT_COL).Value = sortKey
        r = r + 1
    Next i

    If r > ApplicationsList_FIRST_DATA_ROW Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range(ws.Cells(ApplicationsList_FIRST_DATA_ROW, ApplicationsList_TEMP_SORT_COL), ws.Cells(r - 1, ApplicationsList_TEMP_SORT_COL)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange ws.Range(ws.Cells(ApplicationsList_HEADER_ROW, 1), ws.Cells(r - 1, ApplicationsList_TEMP_SORT_COL))
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
        For i = ApplicationsList_FIRST_DATA_ROW To r - 1
            ws.Cells(i, 1).Value = i - ApplicationsList_FIRST_DATA_ROW + 1
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
    If lastRow < ApplicationsList_FIRST_DATA_ROW Then Exit Sub

    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare

    ' Count each application number separately within each object.
    For r = ApplicationsList_FIRST_DATA_ROW To lastRow
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
    For r = ApplicationsList_FIRST_DATA_ROW To lastRow
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

    If Not WorksheetExists(ApplicationsList_DATA_SHEET_NAME) Then
        LogLine "Предупреждение: лист 'data' не найден; столбцы 'Отдел' и 'Ответственное лицо' не заполнены автоматически."
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    For r = ApplicationsList_FIRST_DATA_ROW To lastRow
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

    If lastRow < ApplicationsList_HEADER_ROW Then lastRow = ApplicationsList_HEADER_ROW
    Set fullRange = ws.Range(ws.Cells(ApplicationsList_HEADER_ROW, 1), ws.Cells(lastRow, ApplicationsList_TEMP_SORT_COL))
    With fullRange
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    With ws.Range(ws.Cells(ApplicationsList_HEADER_ROW, 1), ws.Cells(ApplicationsList_HEADER_ROW, 10))
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
    ws.Columns(ApplicationsList_TEMP_SORT_COL).Hidden = True
    ws.Columns(1).HorizontalAlignment = xlCenter
    ws.Columns(3).HorizontalAlignment = xlCenter
    ws.Columns(8).HorizontalAlignment = xlCenter
    ws.Columns(9).HorizontalAlignment = xlCenter
    ws.Columns(10).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(ApplicationsList_FIRST_DATA_ROW, 1), ws.Cells(lastRow, 10)).RowHeight = 20
    ws.Columns(4).WrapText = True
    ws.Columns(6).WrapText = True
    ws.Columns(7).WrapText = True
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetRegisterWorkbook().Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = GetRegisterWorkbook().Worksheets.Add(After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetRegisterWorkbook().Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
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


' ===== Source module: modReportAnalysis.bas =====
' Word report analysis. Uses late binding; Microsoft Word must be installed.
' Results are written to worksheet "Упоминания отчёта".
' This module does not modify the source report.

Private Const ReportAnalysis_RESULT_SHEET As String = "Упоминания отчёта"
Private Const ReportAnalysis_FIRST_RESULT_ROW As Long = 2

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

    ' Preserve previous results until the new scan completes successfully.
    On Error Resume Next
    Set ws = GetRegisterWorkbook().Worksheets(ReportAnalysis_RESULT_SHEET)
    On Error GoTo Failed
    originalSheetExisted = Not (ws Is Nothing)
    If originalSheetExisted Then
        ws.Copy After:=ws
        Set backupWs = ActiveSheet
        backupName = "tmpMentions_" & Format$(Now, "yymmdd_hhnnss") & "_" & CStr(GetRegisterWorkbook().Worksheets.Count)
        backupWs.Name = backupName
        Set ws = GetRegisterWorkbook().Worksheets(ReportAnalysis_RESULT_SHEET)
    Else
        Set ws = GetOrCreateReportSheet(ReportAnalysis_RESULT_SHEET)
    End If
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "№"
    ws.Cells(1, 2).Value = "Упоминание"
    ws.Cells(1, 3).Value = "Страница"
    ws.Cells(1, 4).Value = "Уровень структуры"
    ws.Cells(1, 5).Value = "Текст абзаца"
    ws.Cells(1, 6).Value = "Файл отчёта"
    rowOut = ReportAnalysis_FIRST_RESULT_ROW

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
                ws.Cells(rowOut, 1).Value = rowOut - ReportAnalysis_FIRST_RESULT_ROW + 1
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
    MsgBox "Проверка завершена. Найдено упоминаний: " & (rowOut - ReportAnalysis_FIRST_RESULT_ROW) & _
           vbCrLf & "Результат записан на лист '" & ReportAnalysis_RESULT_SHEET & "'.", vbInformation, "Проверка Word-отчёта"
    GoTo CleanExit

Failed:
    cancelled = (Err.Number = 18)
    errText = Err.Description
    On Error Resume Next
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit

    ' Roll back the output sheet so a failed/cancelled scan cannot leave partial results.
    If Not backupWs Is Nothing Then
        Application.DisplayAlerts = False
        Set damagedWs = ws
        If damagedWs Is Nothing Then
            Err.Clear
            backupWs.Name = ReportAnalysis_RESULT_SHEET
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
                backupWs.Name = ReportAnalysis_RESULT_SHEET
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
                    damagedWs.Name = ReportAnalysis_RESULT_SHEET
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

    ' Keep the original page-count column so Ctrl+Break or a fatal error can roll back.
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

    ' Do not swallow Ctrl+Break: let the caller restore the page-count snapshot.
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

        ' Do not let the local best-effort page-break handling swallow Ctrl+Break.
        If pageBreakError = 18 Then
            Err.Raise pageBreakError, "CountWorkbookPrintPages", pageBreakDescription
        ElseIf pageBreakError <> 0 Then
            h = 1
            v = 1
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


' ===== Source module: modTableNumbering.bas =====
' Checks every table exposed by Word's document Tables collection.
' Produces a preview on an Excel sheet before modifying the Word document.
' Missing captions are inserted as Word SEQ fields, then document fields are updated.
' Word is automated late-bound; Microsoft Word must be installed.

Private Const TableNumbering_TABLE_REPORT_SHEET As String = "Проверка таблиц"
Private Const TableNumbering_FIRST_DATA_ROW As Long = 2
Private Const TableNumbering_WD_STATISTIC_PAGES As Long = 2
Private Const TableNumbering_WD_CAPTION_POSITION_ABOVE As Long = 0
Private Const TableNumbering_WD_FIELD_SEQUENCE As Long = 12
Private Const TableNumbering_WD_WITHIN_TABLE As Long = 12

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

    ' Build the new report on a temporary sheet. Keep the previous report intact
    ' until the scan has completed successfully.
    On Error Resume Next
    Set priorWs = GetRegisterWorkbook().Worksheets(TableNumbering_TABLE_REPORT_SHEET)
    On Error GoTo Failed
    stageName = GetUniqueTableAuditSheetName("_tmp_table_audit")
    Set stageWs = GetRegisterWorkbook().Worksheets.Add(After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count))
    stageWs.Name = stageName
    Set ws = stageWs
    ws.Range("A1:G1").Value = Array("№ таблицы", "Страница", "Подпись перед таблицей", _
                                    "Поле SEQ «Таблица»", "Результат проверки", "Путь к отчёту", "Индекс таблицы")
    r = TableNumbering_FIRST_DATA_ROW
    tableCount = wdDoc.Tables.Count

    For Each tbl In wdDoc.Tables
        DoEvents
        pageNo = 0
        captionText = GetPreviousCaptionText(wdDoc, tbl.Range.Start)
        hasSeq = HasTableSequenceField(wdDoc, tbl.Range.Start)

        On Error Resume Next
        pageNo = tbl.Range.Information(3) ' wdActiveEndPageNumber
        On Error GoTo Failed

        ws.Cells(r, 1).Value = r - TableNumbering_FIRST_DATA_ROW + 1
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
        ws.Cells(r, 7).Value = r - TableNumbering_FIRST_DATA_ROW + 1
        r = r + 1
    Next tbl

    FormatTableReport ws, r - 1

    ' Commit only after the scan and formatting are complete.
    If Not priorWs Is Nothing Then
        priorBackupName = GetUniqueTableAuditSheetName("_bak_table_audit")
        priorWs.Name = priorBackupName
        priorRenamed = True
    End If
    On Error GoTo CommitFailed
    ws.Name = TableNumbering_TABLE_REPORT_SHEET
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
                   vbCrLf & "Отчёт проверки записан на лист '" & TableNumbering_TABLE_REPORT_SHEET & "'.", _
                   vbExclamation, "Проверка таблиц"
        Else
            MsgBox "Проверка завершена. Найдено таблиц: " & tableCount & _
                   vbCrLf & "У всех таблиц найдены подписи с полем SEQ «Таблица»." & _
                   vbCrLf & "Отчёт проверки записан на лист '" & TableNumbering_TABLE_REPORT_SHEET & "'.", _
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
        MsgBox "Изменения не внесены. Предварительный отчёт сохранён на листе '" & TableNumbering_TABLE_REPORT_SHEET & "'.", _
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

    ' Build the final-state report on a separate temporary sheet. The current
    ' preview remains intact until the refreshed report is completely populated.
    finalStageName = GetUniqueTableAuditSheetName("_tmp_table_final")
    Set finalWs = GetRegisterWorkbook().Worksheets.Add(After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count))
    finalWs.Name = finalStageName
    finalWs.Range("A1:G1").Value = Array("№ таблицы", "Страница", "Подпись перед таблицей", _
                                         "Поле SEQ «Таблица»", "Результат проверки", "Путь к отчёту", "Индекс таблицы")
    r = TableNumbering_FIRST_DATA_ROW
    For Each tbl In wdDoc.Tables
        DoEvents
        pageNo = 0
        captionText = GetPreviousCaptionText(wdDoc, tbl.Range.Start)
        hasSeq = HasTableSequenceField(wdDoc, tbl.Range.Start)
        On Error Resume Next
        pageNo = tbl.Range.Information(3)
        On Error GoTo Failed

        finalWs.Cells(r, 1).Value = r - TableNumbering_FIRST_DATA_ROW + 1
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
        finalWs.Cells(r, 7).Value = r - TableNumbering_FIRST_DATA_ROW + 1
        r = r + 1
    Next tbl
    FormatTableReport finalWs, r - 1

    ' Swap only after the final report is complete; keep the old preview as a
    ' rollback point until the new sheet has the canonical name.
    finalOldName = GetUniqueTableAuditSheetName("_bak_table_final")
    ws.Name = finalOldName
    finalOldRenamed = True
    On Error GoTo FinalReportCommitFailed
    finalWs.Name = TableNumbering_TABLE_REPORT_SHEET
    finalReportCommitted = True
    Set ws = finalWs
    Set finalWs = Nothing
    Application.DisplayAlerts = False
    ' Deleting the old preview is cleanup, not part of the Word edit transaction.
    ' If Excel refuses to delete it, keep the backup sheet rather than restoring
    ' the Word file after the final report has already been committed.
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
           "Итоговый список: лист '" & TableNumbering_TABLE_REPORT_SHEET & "'.", _
           vbInformation, "Нумерация таблиц"
    GoTo CleanExit

FinalReportCommitFailed:
    errText = Err.Description
    On Error Resume Next
    If finalOldRenamed And Not ws Is Nothing Then ws.Name = TableNumbering_TABLE_REPORT_SHEET
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
        If Not reportCommitted Then priorWs.Name = TableNumbering_TABLE_REPORT_SHEET
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
        ws.Name = TableNumbering_TABLE_REPORT_SHEET
        finalOldRenamed = False
    End If
    If priorRenamed And Not priorWs Is Nothing And Not reportCommitted Then
        priorWs.Name = TableNumbering_TABLE_REPORT_SHEET
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

    ' Let Word create a real caption paragraph above the table, then change its
    ' SEQ field format so the field itself displays "Таблица № 1", etc.
    tbl.Range.Select
    wdApp.Selection.InsertCaption "Таблица", "", , TableNumbering_WD_CAPTION_POSITION_ABOVE, ""

    captionStart = tbl.Range.Start
    Set captionPara = GetCaptionParagraphBeforeTable(wdDoc, captionStart)
    If captionPara Is Nothing Then
        Err.Raise vbObjectError + 710, "InsertCustomTableCaption", _
                  "Не удалось найти созданную подпись над таблицей."
    End If

    For Each fld In captionPara.Fields
        If InStr(1, CStr(fld.Code.Text), "SEQ", vbTextCompare) > 0 And _
           InStr(1, CStr(fld.Code.Text), "ТАБЛИЦА", vbTextCompare) > 0 Then
            ' Remove Word's literal label before the field; the field format
            ' supplies the complete visible caption.
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
    If nextRow < TableNumbering_FIRST_DATA_ROW Then nextRow = TableNumbering_FIRST_DATA_ROW
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

    ' Skip a few empty paragraphs, then inspect only the nearest non-empty
    ' paragraph. A SEQ field alone is not enough: it must be in a caption-like
    ' paragraph beginning with the visible label "Таблица".
    For paraIndex = wdDoc.Paragraphs.Count To 1 Step -1
        Set para = wdDoc.Paragraphs(paraIndex).Range
        If para.Start >= tableStart Or para.Information(TableNumbering_WD_WITHIN_TABLE) Then
            ' Skip the current table and paragraphs inside any earlier table.
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
            ' Skip paragraphs at or inside the current table.
        ElseIf para.End <= tableStart Then
            If para.Information(TableNumbering_WD_WITHIN_TABLE) = False Then
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


' ===== Source module: modBulletListAudit.bas =====
' Finds consecutive paragraphs that use typed dash/bullet characters instead of
' Word's native list formatting. Shows a preview before any edits.
' Tests should be run on a copy of the report. Microsoft Word must be installed.

Private Const BulletListAudit_BULLET_REPORT_SHEET As String = "Проверка списков"
Private Const BulletListAudit_FIRST_DATA_ROW As Long = 2
Private Const BulletListAudit_WD_LIST_NO_NUMBERING As Long = 0

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
    Dim runIndent As Double, currentIndent As Double, unsafeGroupsFound As Long
    Dim unsafeGroup As Boolean
    Dim backupPath As String
    Dim changesStarted As Boolean
    Dim fsoRestore As Object
    Dim oldStatusBar As Variant
    Dim oldCancelKey As XlEnableCancelKey
    Dim priorWs As Worksheet
    Dim stageWs As Worksheet
    Dim stageName As String, priorBackupName As String
    Dim reportCommitted As Boolean, priorRenamed As Boolean

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
        .Title = "Выберите Word-отчёт для проверки маркированных списков"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Документы Word", "*.doc;*.docx;*.docm"
        If .Show <> -1 Then GoTo CleanExit
        reportPath = .SelectedItems(1)
    End With

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(reportPath, False, False, False)

    ' Scan into a staging sheet so interruption cannot erase the previous report.
    On Error Resume Next
    Set priorWs = GetRegisterWorkbook().Worksheets(BulletListAudit_BULLET_REPORT_SHEET)
    On Error GoTo Failed
    stageName = GetUniqueBulletAuditSheetName("_tmp_bullet_audit")
    Set stageWs = GetRegisterWorkbook().Worksheets.Add(After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count))
    stageWs.Name = stageName
    Set ws = stageWs
    ws.Range("A1:G1").Value = Array("Группа", "Страница", "Текст абзаца", "Найденная маркировка", _
                                    "Предлагаемое действие", "Начальный абзац (служебное)", "Конечный абзац (служебное)")
    rowOut = BulletListAudit_FIRST_DATA_ROW
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
                unsafeGroup = False
                runIndent = GetBulletIndentLevel(para, txt)
                j = i
                Do While j <= pCount
                    Set para = wdDoc.Paragraphs(j)
                    txt = CleanBulletParagraphText(CStr(para.Range.Text))
                    markerOffset = 0
                    markerLength = 0
                    If IsNativeWordList(para) Then Exit Do
                    If Not GetTypedMarker(txt, markerOffset, markerLength) Then Exit Do
                    currentIndent = GetBulletIndentLevel(para, txt)
                    If currentIndent <> runIndent Then unsafeGroup = True
                    j = j + 1
                Loop
                runEnd = j - 1
                runLength = runEnd - runStart + 1

                ' Avoid converting isolated dash-led paragraphs: they may be ordinary prose.
                If runLength >= 2 Then
                    If unsafeGroup Then unsafeGroupsFound = unsafeGroupsFound + 1
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

    ' Commit the completed scan by swapping sheet names only after all rows are ready.
    If Not priorWs Is Nothing Then
        priorBackupName = GetUniqueBulletAuditSheetName("_bak_bullet_audit")
        priorWs.Name = priorBackupName
        priorRenamed = True
    End If
    On Error GoTo CommitFailed
    ws.Name = BulletListAudit_BULLET_REPORT_SHEET
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

    If unsafeGroupsFound > 0 Then
        MsgBox "Найдены группы ручных маркеров с разными отступами (возможные вложенные пункты): " & unsafeGroupsFound & _
               vbCrLf & "Отчёт сохранён на листе '" & BulletListAudit_BULLET_REPORT_SHEET & "'. Автоматическое исправление не выполнено, чтобы не потерять структуру вложенности." & _
               vbCrLf & "Проверьте эти группы вручную и запустите макрос повторно после подготовки документа.", _
               vbExclamation, "Требуется ручная проверка вложенных списков"
        GoTo CleanExit
    End If

    If paragraphsFound = 0 Then
        MsgBox "Не найдено групп из двух или более подряд идущих абзацев с ручными маркерами, которые можно уверенно преобразовать." & _
               vbCrLf & "Одиночные абзацы с тире намеренно не меняются, чтобы не испортить обычный текст." & _
               vbCrLf & "Отчёт проверки записан на лист '" & BulletListAudit_BULLET_REPORT_SHEET & "'.", _
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
        MsgBox "Изменения не внесены. Предварительный отчёт сохранён на листе '" & BulletListAudit_BULLET_REPORT_SHEET & "'.", _
               vbInformation, "Проверка списков"
        GoTo CleanExit
    End If

    ' Create a same-folder backup only after the user confirms the proposed changes.
    backupPath = CreateBulletBackupPath(reportPath)
    FileCopy reportPath, backupPath
    LogBulletBackup ws, backupPath
    changesStarted = True

    ' Process groups from the end of the document backwards so earlier paragraph indexes remain valid.
    Dim groupStartRows() As Long, groupEndRows() As Long, groupCount As Long
    Dim rr As Long, priorGroup As Variant
    ReDim groupStartRows(1 To groupsFound)
    ReDim groupEndRows(1 To groupsFound)
    rr = BulletListAudit_FIRST_DATA_ROW
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
           "Документ сохранён. Проверьте визуально отступы и переносы." & vbCrLf & _
           "Резервная копия: " & backupPath, _
           vbInformation, "Маркированные списки Word"
    GoTo CleanExit

CommitFailed:
    errText = Err.Description
    On Error Resume Next
    If priorRenamed And Not priorWs Is Nothing Then
        If Not reportCommitted Then priorWs.Name = BulletListAudit_BULLET_REPORT_SHEET
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
    If Err.Number = 18 Then cancelled = True
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
    If priorRenamed And Not priorWs Is Nothing And Not reportCommitted Then
        priorWs.Name = BulletListAudit_BULLET_REPORT_SHEET
    End If
    Application.DisplayAlerts = oldAlerts
    If changesStarted And Len(backupPath) > 0 Then
        Err.Clear
        Set fsoRestore = CreateObject("Scripting.FileSystemObject")
        fsoRestore.CopyFile backupPath, reportPath, True
        If Err.Number = 0 Then
            LogBulletBackup ws, "Исходный файл восстановлен из копии: " & backupPath
        Else
            Err.Clear
            LogBulletBackup ws, "Автовосстановление не удалось. Резервная копия сохранена: " & backupPath
        End If
    End If
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
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
End Sub

Private Function GetUniqueBulletAuditSheetName(ByVal baseName As String) As String
    Dim candidate As String
    Dim suffix As Long
    candidate = Left$(baseName, 31)
    suffix = 1
    Do While BulletAuditSheetExists(candidate)
        candidate = Left$(baseName, 27) & "_" & CStr(suffix)
        suffix = suffix + 1
    Loop
    GetUniqueBulletAuditSheetName = candidate
End Function

Private Function BulletAuditSheetExists(ByVal sheetName As String) As Boolean
    Dim probe As Worksheet
    On Error Resume Next
    Set probe = GetRegisterWorkbook().Worksheets(sheetName)
    BulletAuditSheetExists = Not probe Is Nothing
    On Error GoTo 0
End Function

Private Function CreateBulletBackupPath(ByVal sourcePath As String) As String
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
    CreateBulletBackupPath = candidate
End Function

Private Sub LogBulletBackup(ByVal ws As Worksheet, ByVal message As String)
    Dim nextRow As Long
    On Error Resume Next
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < BulletListAudit_FIRST_DATA_ROW Then nextRow = BulletListAudit_FIRST_DATA_ROW
    ws.Cells(nextRow, 1).Value = "Резервная копия"
    ws.Cells(nextRow, 5).Value = message
    ws.Cells(nextRow, 6).Value = Now
    On Error GoTo 0
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
    If ch <> "-" And ch <> ChrW(8211) And ch <> ChrW(8212) And _
       ch <> ChrW(8226) And ch <> ChrW(9679) And ch <> ChrW(9675) And _
       ch <> ChrW(9702) And ch <> ChrW(9642) And ch <> ChrW(9643) And _
       ch <> ChrW(183) And ch <> "*" Then Exit Function
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

Private Function GetBulletIndentLevel(ByVal para As Object, ByVal paragraphText As String) As Double
    Dim i As Long
    Dim ch As String
    On Error Resume Next
    GetBulletIndentLevel = CDbl(para.Format.LeftIndent)
    If Err.Number <> 0 Then
        Err.Clear
        GetBulletIndentLevel = 0
    End If
    On Error GoTo 0

    paragraphText = Replace(paragraphText, vbCr, "")
    paragraphText = Replace(paragraphText, Chr$(7), "")
    For i = 1 To Len(paragraphText)
        ch = Mid$(paragraphText, i, 1)
        If ch = " " Then
            GetBulletIndentLevel = GetBulletIndentLevel + 1
        ElseIf ch = vbTab Then
            GetBulletIndentLevel = GetBulletIndentLevel + 4
        Else
            Exit For
        End If
    Next i
End Function

Private Function IsNativeWordList(ByVal para As Object) As Boolean
    On Error Resume Next
    IsNativeWordList = (CLng(para.Range.ListFormat.ListType) <> BulletListAudit_WD_LIST_NO_NUMBERING)
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
    Set GetOrCreateBulletReportSheet = GetRegisterWorkbook().Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateBulletReportSheet Is Nothing Then
        Set GetOrCreateBulletReportSheet = GetRegisterWorkbook().Worksheets.Add(After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count))
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


' ===== Source module: modRenameFiles.bas =====
' Safe file renaming from the existing register sheet.
' Columns: B=Object, C=Application number, D=Title, G=File hyperlink.
' Preview + explicit confirmation; never overwrites an existing destination.

Private Const RenameFiles_FIRST_DATA_ROW As Long = 3
Private Const RenameFiles_COL_OBJECT As Long = 2
Private Const RenameFiles_COL_APP_NUMBER As Long = 3
Private Const RenameFiles_COL_TITLE As Long = 4
Private Const RenameFiles_COL_LINK As Long = 7
Private Const RenameFiles_LOG_SHEET As String = "Журнал переименования"

Public Sub RenameFilesFromRegister()
    Dim ws As Worksheet, logWs As Worksheet
    Dim lastRow As Long, r As Long, logRow As Long
    Dim oldPath As String, newPath As String, newName As String, ext As String
    Dim objectName As String, appNo As String, titleText As String
    Dim plan As Collection, rec As Object, seen As Object, seenSource As Object, item As Variant
    Dim preview As String, problems As String, failure As String, batchId As String
    Dim renamed As Long, failed As Long
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim oldStatusBar As Variant
    Dim oldCancelKey As XlEnableCancelKey
    Dim renamedItems As Collection
    Dim rollbackItem As Variant
    Dim rollbackIndex As Long
    Dim rollbackError As String

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    oldCancelKey = Application.EnableCancelKey
    Set renamedItems = New Collection
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = GetActiveRegisterSheet()
    If ws Is Nothing Then
        MsgBox "Не найден лист перечня в книге результата. Сначала сформируйте перечень.", vbExclamation
        GoTo CleanExit
    End If
    lastRow = ws.Cells(ws.Rows.Count, RenameFiles_COL_LINK).End(xlUp).Row
    If lastRow < RenameFiles_FIRST_DATA_ROW Then
        MsgBox "На активном листе нет строк для переименования.", vbInformation
        GoTo CleanExit
    End If

    Set plan = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Set seenSource = CreateObject("Scripting.Dictionary")
    seenSource.CompareMode = vbTextCompare

    For r = RenameFiles_FIRST_DATA_ROW To lastRow
        DoEvents
        oldPath = RegisterFilePath(ws.Cells(r, RenameFiles_COL_LINK))
        If Len(oldPath) = 0 Then GoTo ContinueRow
        If Len(Dir$(oldPath)) = 0 Then
            problems = problems & "Строка " & r & ": исходный файл не найден — " & oldPath & vbCrLf
            GoTo ContinueRow
        End If
        If seenSource.Exists(oldPath) Then
            problems = problems & "Строка " & r & ": этот исходный файл уже указан в строке " & _
                       seenSource(oldPath) & "; повторная обработка пропущена." & vbCrLf
            GoTo ContinueRow
        End If

        objectName = Trim$(CStr(ws.Cells(r, RenameFiles_COL_OBJECT).Value))
        appNo = Trim$(CStr(ws.Cells(r, RenameFiles_COL_APP_NUMBER).Value))
        titleText = Trim$(CStr(ws.Cells(r, RenameFiles_COL_TITLE).Value))
        If Len(objectName) = 0 Or Len(appNo) = 0 Or Len(titleText) = 0 Then
            problems = problems & "Строка " & r & ": заполните Объект, № Прил. и Наименование." & vbCrLf
            GoTo ContinueRow
        End If

        ext = FileExtensionFromPath(oldPath)
        newName = CleanFileName(objectName & "_Приложение № " & appNo & "_" & titleText) & ext
        newPath = Left$(oldPath, InStrRev(oldPath, Application.PathSeparator)) & newName

        ' The first row with complete naming fields is the candidate for this source,
        ' even when its name is already correct or the target is blocked. Later rows
        ' must not rename the same physical file using a different register row.
        seenSource(oldPath) = r
        If StrComp(oldPath, newPath, vbTextCompare) = 0 Then
            problems = problems & "Строка " & r & ": имя файла уже соответствует полям перечня; переименование не требуется." & vbCrLf
            GoTo ContinueRow
        End If
        If seen.Exists(newPath) Then
            problems = problems & "Строка " & r & ": дублируется целевое имя — " & newName & vbCrLf
            GoTo ContinueRow
        End If
        seen(newPath) = True
        If Len(Dir$(newPath)) > 0 Then
            problems = problems & "Строка " & r & ": целевой файл уже существует — " & newPath & vbCrLf
            GoTo ContinueRow
        End If

        Set rec = CreateObject("Scripting.Dictionary")
        rec("Row") = r
        rec("OldPath") = oldPath
        rec("NewPath") = newPath
        rec("OldName") = Mid$(oldPath, InStrRev(oldPath, Application.PathSeparator) + 1)
        rec("NewName") = newName
        plan.Add rec
ContinueRow:
    Next r

    If plan.Count = 0 Then
        MsgBox "Нет безопасных переименований для выполнения." & IIf(Len(problems) > 0, vbCrLf & problems, ""), vbInformation
        GoTo CleanExit
    End If

    preview = "Будут переименованы файлы: " & plan.Count & vbCrLf & vbCrLf
    For Each item In plan
        preview = preview & CStr(item("OldName")) & vbCrLf & "  → " & CStr(item("NewName")) & vbCrLf & vbCrLf
        If Len(preview) > 5500 Then
            preview = preview & "... Остальные файлы будут записаны в журнал." & vbCrLf
            Exit For
        End If
    Next item
    If Len(problems) > 0 Then preview = preview & "Пропущенные строки:" & vbCrLf & problems
    If MsgBox(preview & vbCrLf & "Продолжить? Исходные файлы будут переименованы.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Подтверждение переименования") <> vbYes Then GoTo CleanExit

    Set logWs = GetOrCreateRenameLog(RenameFiles_LOG_SHEET)
    If Len(CStr(logWs.Cells(1, 1).Value)) = 0 Then
        logWs.Range("A1:G1").Value = Array("Дата", "Лист перечня", "Строка перечня", "Старый путь", "Новое имя", "Новый путь", "Статус/ошибка")
    End If
    If Len(CStr(logWs.Cells(1, 8).Value)) = 0 Then logWs.Cells(1, 8).Value = "ID пакета"
    logRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row + 1
    batchId = Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(logRow, "000000")

    For Each item In plan
        DoEvents
        failure = TryRenameFile(CStr(item("OldPath")), CStr(item("NewPath")))
        ' Track the filesystem change immediately, before any worksheet/log operations can fail.
        If Len(failure) = 0 Then renamedItems.Add item
        logWs.Cells(logRow, 1).Value = Now
        logWs.Cells(logRow, 2).Value = ws.Name
        logWs.Cells(logRow, 3).Value = item("Row")
        logWs.Cells(logRow, 4).Value = item("OldPath")
        logWs.Cells(logRow, 5).Value = item("NewName")
        logWs.Cells(logRow, 6).Value = item("NewPath")
        If Len(failure) = 0 Then
            ' A successful filesystem rename is not enough: the register must point to it.
            ' If hyperlink replacement fails, raise into the batch handler so the file rename
            ' is rolled back rather than leaving a stale or missing link.
            On Error Resume Next
            Err.Clear
            ws.Cells(CLng(item("Row")), RenameFiles_COL_LINK).Hyperlinks.Delete
            If Err.Number = 0 Then
                ws.Hyperlinks.Add Anchor:=ws.Cells(CLng(item("Row")), RenameFiles_COL_LINK), _
                    Address:=CStr(item("NewPath")), TextToDisplay:="Открыть файл"
            End If
            If Err.Number <> 0 Then
                failure = "Не удалось обновить гиперссылку в строке " & item("Row") & ": " & Err.Description
                Err.Clear
                On Error GoTo Failed
                Err.Raise vbObjectError + 2401, "RenameFilesFromRegister", failure
            End If
            On Error GoTo Failed
            logWs.Cells(logRow, 7).Value = "Переименован"
            logWs.Cells(logRow, 8).Value = batchId
            renamed = renamed + 1
        Else
            logWs.Cells(logRow, 7).Value = "Ошибка: " & failure
            logWs.Cells(logRow, 8).Value = batchId
            failed = failed + 1
        End If
        logRow = logRow + 1
        Application.StatusBar = "Переименование: " & (renamed + failed) & " из " & plan.Count
    Next item

    FormatRenameLog logWs
    MsgBox "Переименование завершено. Успешно: " & renamed & "; ошибок: " & failed & _
           IIf(Len(problems) > 0, vbCrLf & vbCrLf & "Пропуски:" & vbCrLf & problems, ""), vbInformation
    GoTo CleanExit

Failed:
    failure = Err.Description
    If Err.Number = 18 Then failure = "Операция отменена пользователем (Ctrl+Break)."
    ' If a batch was interrupted, reverse every rename completed in this run.
    ' Disable another Ctrl+Break while recovery is in progress; otherwise a second
    ' interruption could stop the rollback halfway through.
    On Error Resume Next
    Application.EnableCancelKey = xlDisabled
    If Not renamedItems Is Nothing Then
        ' Undo in reverse order, like a transaction stack.
        For rollbackIndex = renamedItems.Count To 1 Step -1
            Set rollbackItem = renamedItems(rollbackIndex)
            rollbackError = TryRenameFile(CStr(rollbackItem("NewPath")), CStr(rollbackItem("OldPath")))
            If Len(rollbackError) = 0 Then
                rollbackError = ""
                If Not ws Is Nothing Then
                    rollbackError = RestoreRegisterHyperlink(ws, CLng(rollbackItem("Row")), CStr(rollbackItem("OldPath")))
                End If
                If Len(rollbackError) > 0 Then
                    problems = problems & "Файл возвращён, но ссылка в строке " & CLng(rollbackItem("Row")) & _
                               " не восстановлена: " & rollbackError & vbCrLf
                End If
                ' Journal records both filesystem recovery and any hyperlink repair problem.
                If Not logWs Is Nothing And Len(batchId) > 0 Then
                    For r = 2 To logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row
                        If Trim$(CStr(logWs.Cells(r, 8).Value)) = batchId And _
                           StrComp(CStr(logWs.Cells(r, 4).Value), CStr(rollbackItem("OldPath")), vbTextCompare) = 0 Then
                            If Len(rollbackError) = 0 Then
                                logWs.Cells(r, 7).Value = "Автоматически отменён после прерывания"
                            Else
                                logWs.Cells(r, 7).Value = "Файл возвращён; ошибка восстановления ссылки: " & rollbackError
                            End If
                            Exit For
                        End If
                    Next r
                End If
            Else
                problems = problems & "Не удалось автоматически отменить: " & CStr(rollbackItem("NewPath")) & _
                           " (" & rollbackError & ")" & vbCrLf
                If Not logWs Is Nothing And Len(batchId) > 0 Then
                    For r = 2 To logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row
                        If Trim$(CStr(logWs.Cells(r, 8).Value)) = batchId And _
                           StrComp(CStr(logWs.Cells(r, 4).Value), CStr(rollbackItem("OldPath")), vbTextCompare) = 0 Then
                            logWs.Cells(r, 7).Value = "Ошибка автоматического отката: " & rollbackError
                            Exit For
                        End If
                    Next r
                End If
            End If
        Next rollbackIndex
    End If
    On Error GoTo 0
    MsgBox failure & vbCrLf & _
           IIf(Len(problems) > 0, "Проблемы и пропуски:" & vbCrLf & problems & vbCrLf, "") & _
           "Выполненные в этом запуске переименования автоматически отменены, где это удалось. " & _
           "Проверьте журнал и пути файлов.", vbExclamation
CleanExit:
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
End Sub

Public Sub RollbackLastRenameBatch()
    Dim logWs As Worksheet, registerWs As Worksheet
    Dim lastRow As Long, r As Long, registerRow As Long, logRow As Long
    Dim oldPath As String, newPath As String, statusText As String
    Dim restored As Long, skipped As Long, errText As String, sheetName As String, batchId As String
    Dim oldScreen As Boolean, oldEvents As Boolean, oldAlerts As Boolean
    Dim oldStatusBar As Variant, oldCancelKey As XlEnableCancelKey
    Dim reversedItems As Collection, item As Object, doneItem As Object
    Dim failure As String, recoveryError As String, itemPath As Variant
    Dim candidateOldPath As String, candidateNewPath As String
    Dim recoveryIndex As Long

    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    oldStatusBar = Application.StatusBar
    oldCancelKey = Application.EnableCancelKey
    Set reversedItems = New Collection
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    If GetRegisterWorkbook() Is Nothing Then
        MsgBox "Не найдена книга результата. Сначала запустите автоматизацию и сформируйте перечень.", _
               vbExclamation, "Откат переименования"
        GoTo CleanExit
    End If

    On Error Resume Next
    Set logWs = GetRegisterWorkbook().Worksheets(RenameFiles_LOG_SHEET)
    On Error GoTo Failed
    If logWs Is Nothing Then
        MsgBox "Журнал переименования не найден. Пока нет операций, которые можно откатить.", _
               vbInformation, "Откат переименования"
        GoTo CleanExit
    End If
    lastRow = logWs.Cells(logWs.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "Журнал переименования пуст.", vbInformation
        GoTo CleanExit
    End If

    batchId = ""
    ' Select the newest batch based on actual file paths, not journal status alone.
    ' This lets a partially completed rollback be retried after a previous error.
    For r = lastRow To 2 Step -1
        candidateOldPath = CStr(logWs.Cells(r, 4).Value)
        candidateNewPath = CStr(logWs.Cells(r, 6).Value)
        If Len(Trim$(CStr(logWs.Cells(r, 8).Value))) > 0 And _
           Len(candidateOldPath) > 0 And Len(candidateNewPath) > 0 Then
            If Len(Dir$(candidateNewPath)) > 0 And Len(Dir$(candidateOldPath)) = 0 Then
                batchId = Trim$(CStr(logWs.Cells(r, 8).Value))
                Exit For
            End If
        End If
    Next r
    If Len(batchId) = 0 Then
        MsgBox "В журнале нет переименованных файлов, доступных для отката. Старые записи без ID пакета автоматически не откатываются.", vbExclamation
        GoTo CleanExit
    End If
    If MsgBox("Отменить последнее переименование (пакет " & batchId & ")? " & _
              "Если исходное имя уже занято, файл будет пропущен. При прерывании макрос попытается вернуть уже откатанные файлы к новым именам.", _
              vbYesNo + vbQuestion + vbDefaultButton2, "Откат переименования") <> vbYes Then GoTo CleanExit

    For r = lastRow To 2 Step -1
        DoEvents
        Application.StatusBar = "Откат переименования: запись " & r & " из " & lastRow
        statusText = CStr(logWs.Cells(r, 7).Value)
        If Trim$(CStr(logWs.Cells(r, 8).Value)) = batchId Then
            oldPath = CStr(logWs.Cells(r, 4).Value)
            newPath = CStr(logWs.Cells(r, 6).Value)
            If Len(Dir$(oldPath)) > 0 And Len(Dir$(newPath)) = 0 Then
                ' Filesystem already reflects a completed rollback; reconcile the journal.
                logWs.Cells(r, 7).Value = "Откатан"
            ElseIf Len(Dir$(newPath)) = 0 Or Len(Dir$(oldPath)) > 0 Then
                logWs.Cells(r, 7).Value = "Откат пропущен: конфликт путей (проверьте оба имени)"
                skipped = skipped + 1
            Else
                errText = TryRenameFile(newPath, oldPath)
                If Len(errText) = 0 Then
                    ' Track immediately after the filesystem change so interruption recovery can find it.
                    Set doneItem = CreateObject("Scripting.Dictionary")
                    doneItem("LogRow") = r
                    doneItem("OldPath") = oldPath
                    doneItem("NewPath") = newPath
                    doneItem("SheetName") = CStr(logWs.Cells(r, 2).Value)
                    doneItem("RegisterRow") = CLng(Val(logWs.Cells(r, 3).Value))
                    reversedItems.Add doneItem

                    logWs.Cells(r, 7).Value = "Откатан"
                    sheetName = CStr(doneItem("SheetName"))
                    registerRow = CLng(doneItem("RegisterRow"))
                    Set registerWs = Nothing
                    On Error Resume Next
                    Set registerWs = GetRegisterWorkbook().Worksheets(sheetName)
                    On Error GoTo Failed
                    recoveryError = ""
                    If Not registerWs Is Nothing And registerRow > 0 Then
                        recoveryError = RestoreRegisterHyperlink(registerWs, registerRow, oldPath)
                    Else
                        recoveryError = "Не найден лист или строка перечня для обновления ссылки."
                    End If
                    If Len(recoveryError) = 0 Then
                        logWs.Cells(r, 7).Value = "Откатан"
                    Else
                        logWs.Cells(r, 7).Value = "Файл возвращён; ошибка восстановления ссылки: " & recoveryError
                        skipped = skipped + 1
                    End If
                    restored = restored + 1
                Else
                    logWs.Cells(r, 7).Value = "Ошибка отката: " & errText
                    skipped = skipped + 1
                End If
            End If
        End If
    Next r
    FormatRenameLog logWs
    MsgBox "Откат завершён. Восстановлено: " & restored & "; пропущено: " & skipped, vbInformation
    GoTo CleanExit

Failed:
    failure = Err.Description
    If Err.Number = 18 Then failure = "Операция отменена пользователем (Ctrl+Break)."
    ' Best-effort recovery: restore files already renamed back to their post-rename paths.
    On Error Resume Next
    Application.EnableCancelKey = xlDisabled
    ' Reverse the rollback actions in reverse order to restore the pre-rollback state.
    For recoveryIndex = reversedItems.Count To 1 Step -1
        Set item = reversedItems(recoveryIndex)
        recoveryError = TryRenameFile(CStr(item("OldPath")), CStr(item("NewPath")))
        logRow = CLng(item("LogRow"))
        If Len(recoveryError) = 0 Then
            logWs.Cells(logRow, 7).Value = "Переименован"
            Set registerWs = Nothing
            Set registerWs = GetRegisterWorkbook().Worksheets(CStr(item("SheetName")))
            recoveryError = ""
            If Not registerWs Is Nothing And CLng(item("RegisterRow")) > 0 Then
                recoveryError = RestoreRegisterHyperlink(registerWs, CLng(item("RegisterRow")), CStr(item("NewPath")))
            Else
                recoveryError = "Не найден лист или строка перечня для обновления ссылки."
            End If
            If Len(recoveryError) > 0 Then
                logWs.Cells(logRow, 7).Value = "Файл возвращён к новому имени; ошибка восстановления ссылки: " & recoveryError
                failure = failure & vbCrLf & "Файл восстановлен, но ссылка не обновлена: " & CStr(item("NewPath")) & _
                          " (" & recoveryError & ")"
            End If
        Else
            logWs.Cells(logRow, 7).Value = "Ошибка восстановления после прерывания: " & recoveryError
            failure = failure & vbCrLf & "Не удалось вернуть файл к новому имени: " & CStr(item("OldPath")) & _
                      " (" & recoveryError & ")"
        End If
    Next recoveryIndex
    On Error GoTo 0
    MsgBox "Откат прерван. Макрос попытался вернуть уже обработанные файлы к состоянию до отката." & vbCrLf & _
           failure & vbCrLf & "Проверьте журнал переименования и фактические имена файлов.", vbExclamation
    GoTo CleanExit

CleanExit:
    On Error Resume Next
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
    On Error GoTo 0
End Sub


Private Function RestoreRegisterHyperlink(ByVal targetWs As Worksheet, ByVal rowNo As Long, ByVal filePath As String) As String
    On Error GoTo LinkError
    targetWs.Cells(rowNo, RenameFiles_COL_LINK).Hyperlinks.Delete
    targetWs.Hyperlinks.Add Anchor:=targetWs.Cells(rowNo, RenameFiles_COL_LINK), _
        Address:=filePath, TextToDisplay:="Открыть файл"
    Exit Function
LinkError:
    RestoreRegisterHyperlink = Err.Description
    Err.Clear
End Function

Private Function TryRenameFile(ByVal oldPath As String, ByVal newPath As String) As String
    On Error GoTo RenameError
    If Len(Dir$(oldPath)) = 0 Then
        TryRenameFile = "Исходный файл не найден."
        Exit Function
    End If
    If Len(Dir$(newPath)) > 0 Then
        TryRenameFile = "Целевой файл уже существует."
        Exit Function
    End If
    Name oldPath As newPath
    Exit Function
RenameError:
    TryRenameFile = Err.Description
    Err.Clear
End Function

Private Function RegisterFilePath(ByVal c As Range) As String
    On Error Resume Next
    If c.Hyperlinks.Count > 0 Then
        RegisterFilePath = CStr(c.Hyperlinks(1).Address)
        If Len(RegisterFilePath) = 0 Then RegisterFilePath = CStr(c.Hyperlinks(1).SubAddress)
    Else
        RegisterFilePath = CStr(c.Value)
    End If
    On Error GoTo 0
    RegisterFilePath = NormalizeRegisterPath(RegisterFilePath)
End Function

Private Function NormalizeRegisterPath(ByVal value As String) As String
    value = Trim$(value)
    If LCase$(Left$(value, 8)) = "file:///" Then value = Mid$(value, 9)
    value = Replace(value, "%20", " ")
    value = Replace(value, "/", "\")
    NormalizeRegisterPath = value
End Function

Private Function FileExtensionFromPath(ByVal filePath As String) As String
    Dim slashPos As Long, dotPos As Long
    slashPos = InStrRev(filePath, Application.PathSeparator)
    dotPos = InStrRev(filePath, ".")
    If dotPos > slashPos Then FileExtensionFromPath = Mid$(filePath, dotPos)
End Function

Private Function CleanFileName(ByVal value As String) As String
    Dim invalidChars As Variant, i As Long
    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(invalidChars) To UBound(invalidChars)
        value = Replace(value, CStr(invalidChars(i)), "_")
    Next i
    Do While InStr(value, "__") > 0
        value = Replace(value, "__", "_")
    Loop
    CleanFileName = Trim$(value)
End Function

Private Function GetOrCreateRenameLog(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateRenameLog = GetRegisterWorkbook().Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateRenameLog Is Nothing Then
        Set GetOrCreateRenameLog = GetRegisterWorkbook().Worksheets.Add(After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count))
        GetOrCreateRenameLog.Name = sheetName
    End If
End Function

Private Sub FormatRenameLog(ByVal ws As Worksheet)
    With ws.UsedRange
        .Font.Name = "Times New Roman"
        .Borders.LineStyle = xlContinuous
    End With
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(198, 224, 180)
    ws.Columns("A:H").AutoFit
End Sub


' ===== Source module: modDataFull.bas =====
' Builds a unique list of Section + Responsible Person pairs from the active register.
' Source columns: F = Responsible Person, H = Section; headers are on row 2, data starts on row 3.
' Output sheet: data_full. Blank values are ignored.

Private Const DataFull_SOURCE_FIRST_ROW As Long = 3
Private Const DataFull_COL_RESPONSIBLE As Long = 6
Private Const DataFull_COL_SECTION As Long = 8
Private Const DataFull_OUTPUT_SHEET As String = "data_full"

Public Sub BuildDataFull()
    Dim sourceWs As Worksheet
    Dim outputWs As Worksheet
    Dim backupWs As Worksheet
    Dim damagedWs As Worksheet
    Dim recoverySheetName As String
    Dim outputWasExisting As Boolean
    Dim operationCompleted As Boolean
    Dim lastRow As Long
    Dim r As Long
    Dim sectionText As String
    Dim personText As String
    Dim key As String
    Dim pairs As Object
    Dim item As Variant
    Dim pairValue As Variant
    Dim outRow As Long
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean
    Dim oldAlerts As Boolean
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

    Set sourceWs = GetActiveRegisterSheet()
    If sourceWs Is Nothing Then
        MsgBox "Не найден лист перечня в книге результата. Сначала сформируйте перечень.", vbExclamation
        GoTo CleanExit
    End If
    If LCase$(sourceWs.Name) = LCase$(DataFull_OUTPUT_SHEET) Or LCase$(sourceWs.Name) = "data" Then
        MsgBox "Сначала откройте лист перечня приложений и запустите макрос ещё раз.", vbInformation
        GoTo CleanExit
    End If

    If Trim$(CStr(sourceWs.Cells(2, DataFull_COL_SECTION).Value)) <> "Раздел" Or _
       Trim$(CStr(sourceWs.Cells(2, DataFull_COL_RESPONSIBLE).Value)) <> "Ответственное лицо" Then
        MsgBox "Активный лист не похож на перечень приложений: ожидаются заголовки 'Ответственное лицо' в F2 и 'Раздел' в H2.", vbExclamation
        GoTo CleanExit
    End If

    lastRow = Application.Max(sourceWs.Cells(sourceWs.Rows.Count, DataFull_COL_SECTION).End(xlUp).Row, _
                               sourceWs.Cells(sourceWs.Rows.Count, DataFull_COL_RESPONSIBLE).End(xlUp).Row)
    Set pairs = CreateObject("Scripting.Dictionary")
    pairs.CompareMode = vbTextCompare

    For r = DataFull_SOURCE_FIRST_ROW To lastRow
        DoEvents
        sectionText = Trim$(CStr(sourceWs.Cells(r, DataFull_COL_SECTION).Value))
        personText = Trim$(CStr(sourceWs.Cells(r, DataFull_COL_RESPONSIBLE).Value))
        If Len(sectionText) > 0 And Len(personText) > 0 Then
            key = sectionText & Chr$(30) & personText
            If Not pairs.Exists(key) Then pairs.Add key, Array(sectionText, personText)
        End If
        Application.StatusBar = "Формирование data_full: строка " & r & " из " & lastRow
    Next r

    On Error Resume Next
    Set outputWs = GetRegisterWorkbook().Worksheets(DataFull_OUTPUT_SHEET)
    On Error GoTo Failed
    outputWasExisting = Not outputWs Is Nothing

    ' Keep a full worksheet copy until the rebuild has completed successfully.
    If outputWasExisting Then
        outputWs.Copy After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count)
        Set backupWs = ActiveSheet
        backupWs.Name = GetUniqueDataFullBackupName()
    Else
        Set outputWs = GetOrCreateDataFullSheet()
    End If

    outputWs.Cells.Clear
    outputWs.Range("A1:B1").Value = Array("Раздел", "Ответственное лицо")
    outRow = 2
    For Each item In pairs.Keys
        pairValue = pairs(item)
        outputWs.Cells(outRow, 1).Value = pairValue(0)
        outputWs.Cells(outRow, 2).Value = pairValue(1)
        outRow = outRow + 1
    Next item

    If outRow > 2 Then
        With outputWs.Sort
            .SortFields.Clear
            .SortFields.Add Key:=outputWs.Range("A2:A" & outRow - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=outputWs.Range("B2:B" & outRow - 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange outputWs.Range("A1:B" & outRow - 1)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If

    With outputWs.Range("A1:B" & Application.Max(1, outRow - 1))
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With
    With outputWs.Range("A1:B1")
        .Interior.Color = RGB(198, 224, 180)
        .Font.Bold = True
        .Font.Italic = True
        .WrapText = True
    End With
    outputWs.Columns("A:B").AutoFit
    outputWs.Activate
    operationCompleted = True
    If Not backupWs Is Nothing Then
        Application.DisplayAlerts = False
        backupWs.Delete
        Application.DisplayAlerts = oldAlerts
        Set backupWs = Nothing
    End If
    MsgBox "Лист data_full обновлён. Уникальных пар: " & pairs.Count, vbInformation
    GoTo CleanExit

Failed:
    Dim failureText As String
    Dim wasCancelled As Boolean
    failureText = Err.Description
    wasCancelled = (Err.Number = 18)

    ' Restore the previous output if an error or Ctrl+Break occurs mid-rebuild.
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not backupWs Is Nothing Then
        Set damagedWs = outputWs
        If damagedWs Is Nothing Then
            Err.Clear
            backupWs.Name = DataFull_OUTPUT_SHEET
            If Err.Number = 0 Then
                Set outputWs = backupWs
                Set backupWs = Nothing
            Else
                Err.Clear
                failureText = failureText & vbCrLf & "Не удалось вернуть резервной копии исходное имя. Резервная копия сохранена под именем " & backupWs.Name & "."
            End If
        Else
            recoverySheetName = GetUniqueDataFullRecoveryName()
            Err.Clear
            damagedWs.Name = recoverySheetName
            If Err.Number = 0 Then
                Err.Clear
                backupWs.Name = DataFull_OUTPUT_SHEET
                If Err.Number = 0 Then
                    Set outputWs = backupWs
                    Set backupWs = Nothing
                    Err.Clear
                    damagedWs.Delete
                    If Err.Number <> 0 Then
                        Err.Clear
                        failureText = failureText & vbCrLf & "Предыдущий лист восстановлен; неполный лист сохранён под именем " & recoverySheetName & "."
                    End If
                Else
                    Err.Clear
                    damagedWs.Name = DataFull_OUTPUT_SHEET
                    failureText = failureText & vbCrLf & "Не удалось переименовать резервную копию. Она сохранена под именем " & backupWs.Name & "."
                End If
            Else
                Err.Clear
                failureText = failureText & vbCrLf & "Не удалось безопасно переименовать неполный лист. Резервная копия сохранена под именем " & backupWs.Name & "."
            End If
        End If
    ElseIf Not outputWasExisting And Not outputWs Is Nothing And Not operationCompleted Then
        outputWs.Delete
    End If
    Application.DisplayAlerts = oldAlerts
    On Error GoTo 0

    If wasCancelled Then
        MsgBox "Формирование data_full отменено пользователем. Предыдущий лист восстановлен, если он существовал." & _
               IIf(InStr(1, failureText, "сохранён", vbTextCompare) > 0 Or InStr(1, failureText, "не удалось", vbTextCompare) > 0, vbCrLf & failureText, ""), vbInformation
    Else
        MsgBox "Не удалось сформировать data_full: " & failureText & _
               IIf(outputWasExisting, vbCrLf & "Предыдущий лист восстановлен, если операция успела изменить его.", ""), vbExclamation
    End If

CleanExit:
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
End Sub

Private Function GetOrCreateDataFullSheet() As Worksheet
    On Error Resume Next
    Set GetOrCreateDataFullSheet = GetRegisterWorkbook().Worksheets(DataFull_OUTPUT_SHEET)
    On Error GoTo 0
    If GetOrCreateDataFullSheet Is Nothing Then
        Set GetOrCreateDataFullSheet = GetRegisterWorkbook().Worksheets.Add(After:=GetRegisterWorkbook().Worksheets(GetRegisterWorkbook().Worksheets.Count))
        GetOrCreateDataFullSheet.Name = DataFull_OUTPUT_SHEET
    End If
End Function

Private Function GetUniqueDataFullRecoveryName() As String
    Dim candidate As String
    Dim suffix As Long
    candidate = "_failed_data_full"
    suffix = 1
    Do While DataFullWorksheetExists(candidate)
        candidate = "_failed_df_" & CStr(suffix)
        suffix = suffix + 1
    Loop
    GetUniqueDataFullRecoveryName = candidate
End Function

Private Function DataFullWorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetRegisterWorkbook().Worksheets(sheetName)
    DataFullWorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function GetUniqueDataFullBackupName() As String
    Dim baseName As String
    Dim candidate As String
    Dim suffix As Long
    Dim ws As Worksheet
    Dim exists As Boolean

    baseName = "df_bak_" & Format$(Now, "yymmddhhnnss")
    candidate = baseName
    suffix = 1
    Do
        exists = False
        For Each ws In GetRegisterWorkbook().Worksheets
            If StrComp(ws.Name, candidate, vbTextCompare) = 0 Then
                exists = True
                Exit For
            End If
        Next ws
        If Not exists Then Exit Do
        candidate = baseName & "_" & CStr(suffix)
        suffix = suffix + 1
    Loop
    GetUniqueDataFullBackupName = candidate
End Function


' ===== Source module: modSafetySupport.bas =====
' Common cancellation/error helpers for macros that use Excel and Word automation.
Public Function IsUserCancellation(ByVal errorNumber As Long) As Boolean
    IsUserCancellation = (errorNumber = 18)
End Function

Public Sub RestoreExcelApplicationState(ByVal screenUpdating As Boolean, _
                                        ByVal enableEvents As Boolean, _
                                        ByVal displayAlerts As Boolean, _
                                        ByVal calculationMode As XlCalculation, _
                                        ByVal oldStatusBar As Variant, _
                                        ByVal oldCancelKey As XlEnableCancelKey)
    On Error Resume Next
    Application.ScreenUpdating = screenUpdating
    Application.EnableEvents = enableEvents
    Application.DisplayAlerts = displayAlerts
    Application.Calculation = calculationMode
    Application.StatusBar = oldStatusBar
    Application.EnableCancelKey = oldCancelKey
    On Error GoTo 0
End Sub
